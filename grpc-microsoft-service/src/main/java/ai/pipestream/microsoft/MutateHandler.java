package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MutateRequest;
import ai.pipestream.microsoft.v1.MutateResponse;
import ai.pipestream.microsoft.v1.WriteError;
import ai.pipestream.microsoft.v1.WriteErrorCode;
import io.grpc.stub.StreamObserver;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.TimeUnit;
import java.util.function.BiConsumer;

/**
 * Receives the ops of one {@code Mutate} call, sends them to Graph in batches and
 * answers each op once.
 */
final class MutateHandler implements StreamObserver<MutateRequest> {

    /** How long a batch waits for more ops before it is sent, by default. */
    static final long DEFAULT_LINGER_MILLIS = 20;

    /**
     * One op with its Graph request.
     *
     * @param request the op
     * @param session the connection it runs on
     * @param plan the Graph request and response mapping
     * @param setter puts the response on the answer
     * @param <R> the op's response type
     */
    private record Planned<R>(MutateRequest request, MicrosoftSessions.Session session,
            WritePlans.Plan<R> plan, BiConsumer<MutateResponse.Builder, R> setter) {

        MutateResponse answer(GraphClient.Reply reply) {
            MutateResponse.Builder answer = MutateResponse.newBuilder().setOpId(request.getOpId());
            setter.accept(answer, plan.mapper().map(reply));
            return answer.build();
        }
    }

    private final MicrosoftSessions sessions;
    private final WritePlans plans;
    private final StreamObserver<MutateResponse> responses;
    private final BlockingQueue<Optional<MutateRequest>> queue =
            new ArrayBlockingQueue<>(2 * GraphClient.MAX_BATCH_SIZE);
    private final Set<String> failed = new HashSet<>();
    private final long lingerMillis;
    private final Thread worker;

    MutateHandler(MicrosoftSessions sessions, WritePlans plans,
            StreamObserver<MutateResponse> responses, long lingerMillis) {
        this.lingerMillis = lingerMillis;
        this.sessions = sessions;
        this.plans = plans;
        this.responses = responses;
        this.worker = Thread.ofVirtual().name("mutate").start(this::run);
    }

    @Override
    public void onNext(MutateRequest request) {
        enqueue(Optional.of(request));
    }

    @Override
    public void onError(Throwable t) {
        worker.interrupt();
    }

    @Override
    public void onCompleted() {
        enqueue(Optional.empty());
    }

    private void enqueue(Optional<MutateRequest> entry) {
        try {
            queue.put(entry);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            worker.interrupt();
        }
    }

    private void run() {
        try {
            boolean open = true;
            while (open) {
                Optional<MutateRequest> next = queue.take();
                List<MutateRequest> batch = new ArrayList<>();
                while (next != null && next.isPresent()
                        && batch.size() < GraphClient.MAX_BATCH_SIZE) {
                    batch.add(next.get());
                    next = batch.size() < GraphClient.MAX_BATCH_SIZE
                            ? queue.poll(lingerMillis, TimeUnit.MILLISECONDS) : null;
                }
                open = next == null || next.isPresent();
                send(batch);
            }
            responses.onCompleted();
        } catch (Throwable t) {
            responses.onError(GraphErrors.toException(t));
        }
    }

    private void send(List<MutateRequest> batch) throws InterruptedException {
        Map<String, List<Planned<?>>> byConnection = new LinkedHashMap<>();
        for (MutateRequest request : batch) {
            try {
                Optional<String> failedDependency = request.getDependsOnList().stream()
                        .filter(failed::contains).findFirst();
                if (failedDependency.isPresent()) {
                    refuse(request, WriteError.newBuilder()
                            .setCode(WriteErrorCode.WRITE_ERROR_CODE_FAILED_PRECONDITION)
                            .setMessage("op " + failedDependency.get() + " failed").build());
                    continue;
                }
                Planned<?> planned = plan(request);
                if (planned.plan().call() == null) {
                    responses.onNext(planned.answer(null));
                    continue;
                }
                byConnection.computeIfAbsent(planned.session().connectionId(),
                        id -> new ArrayList<>()).add(planned);
            } catch (InterruptedException e) {
                throw e;
            } catch (Exception e) {
                refuse(request, GraphErrors.toWriteError(e));
            }
        }
        for (List<Planned<?>> ops : byConnection.values()) {
            sendBatch(ops);
        }
    }

    private void sendBatch(List<Planned<?>> ops) throws InterruptedException {
        Map<String, Integer> indexes = new HashMap<>();
        List<GraphClient.BatchCall> calls = new ArrayList<>();
        for (int i = 0; i < ops.size(); i++) {
            MutateRequest request = ops.get(i).request();
            List<Integer> dependsOn = new ArrayList<>();
            for (String dependency : request.getDependsOnList()) {
                Integer index = indexes.get(dependency);
                if (index != null) {
                    dependsOn.add(index);
                }
            }
            calls.add(new GraphClient.BatchCall(ops.get(i).plan().call(), dependsOn));
            if (!request.getOpId().isEmpty()) {
                indexes.put(request.getOpId(), i);
            }
        }
        List<GraphClient.BatchReply> replies;
        try {
            replies = ops.get(0).session().graph().batch(calls);
        } catch (IOException | RuntimeException e) {
            WriteError error = GraphErrors.toWriteError(e);
            ops.forEach(op -> refuse(op.request(), error));
            return;
        }
        for (int i = 0; i < ops.size(); i++) {
            Planned<?> op = ops.get(i);
            GraphClient.BatchReply reply = replies.get(i);
            if (reply.error() != null) {
                refuse(op.request(), GraphErrors.toWriteError(reply.error()));
                continue;
            }
            try {
                responses.onNext(op.answer(reply.reply()));
            } catch (io.grpc.StatusRuntimeException e) {
                if (e.getStatus().getCode() == io.grpc.Status.Code.CANCELLED) {
                    throw e;
                }
                refuse(op.request(), GraphErrors.toWriteError(e));
            }
        }
    }

    private void refuse(MutateRequest request, WriteError error) {
        if (!request.getOpId().isEmpty()) {
            failed.add(request.getOpId());
        }
        responses.onNext(MutateResponse.newBuilder().setOpId(request.getOpId()).setError(error)
                .build());
    }

    private Planned<?> plan(MutateRequest request) throws IOException, InterruptedException {
        return switch (request.getOpCase()) {
            case CREATE_FOLDER -> planned(request, request.getCreateFolder().getConnectionId(),
                    session -> plans.createFolder(request.getCreateFolder()),
                    MutateResponse.Builder::setCreateFolder);
            case UPDATE_DRIVE_ITEM -> planned(request,
                    request.getUpdateDriveItem().getConnectionId(),
                    session -> plans.updateDriveItem(request.getUpdateDriveItem()),
                    MutateResponse.Builder::setUpdateDriveItem);
            case DELETE_DRIVE_ITEM -> planned(request,
                    request.getDeleteDriveItem().getConnectionId(),
                    session -> plans.deleteDriveItem(request.getDeleteDriveItem()),
                    MutateResponse.Builder::setDeleteDriveItem);
            case RESTORE_DRIVE_ITEM_VERSION -> planned(request,
                    request.getRestoreDriveItemVersion().getConnectionId(),
                    session -> plans.restoreDriveItemVersion(request.getRestoreDriveItemVersion()),
                    MutateResponse.Builder::setRestoreDriveItemVersion);
            case CREATE_LIST_ITEM -> planned(request, request.getCreateListItem().getConnectionId(),
                    session -> plans.createListItem(session, request.getCreateListItem()),
                    MutateResponse.Builder::setCreateListItem);
            case UPDATE_LIST_ITEM_FIELDS -> planned(request,
                    request.getUpdateListItemFields().getConnectionId(),
                    session -> plans.updateListItemFields(session,
                            request.getUpdateListItemFields()),
                    MutateResponse.Builder::setUpdateListItemFields);
            case DELETE_LIST_ITEM -> planned(request, request.getDeleteListItem().getConnectionId(),
                    session -> plans.deleteListItem(request.getDeleteListItem()),
                    MutateResponse.Builder::setDeleteListItem);
            case OP_NOT_SET -> throw new IllegalArgumentException("the op is empty");
        };
    }

    /**
     * Builds the plan of an op on its connection.
     *
     * @param <R> the op's response type
     */
    @FunctionalInterface
    private interface Planner<R> {
        WritePlans.Plan<R> plan(MicrosoftSessions.Session session)
                throws IOException, InterruptedException;
    }

    private <R> Planned<R> planned(MutateRequest request, String connectionId, Planner<R> planner,
            BiConsumer<MutateResponse.Builder, R> setter) throws IOException, InterruptedException {
        MicrosoftSessions.Session session = sessions.resolve(connectionId);
        return new Planned<>(request, session, planner.plan(session), setter);
    }
}
