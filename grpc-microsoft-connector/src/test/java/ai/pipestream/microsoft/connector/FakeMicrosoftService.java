package ai.pipestream.microsoft.connector;

import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.GetMeRequest;
import ai.pipestream.microsoft.v1.GetMeResponse;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import com.google.protobuf.Timestamp;
import io.grpc.Status;
import io.grpc.stub.StreamObserver;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

/** In-process MicrosoftService for connector tests. */
final class FakeMicrosoftService extends MicrosoftServiceGrpc.MicrosoftServiceImplBase {

    GraphUser me = GraphUser.newBuilder()
            .setId("user-1")
            .setDisplayName("Bot")
            .setUserPrincipalName("bot@contoso.com")
            .setMail("bot@contoso.com")
            .build();
    Status meError;
    final List<MicrosoftChange> changes = new ArrayList<>();
    Status syncError;
    String resumeCursor = "cursor-1";
    final AtomicInteger syncCalls = new AtomicInteger();

    @Override
    public void getMe(GetMeRequest request, StreamObserver<GetMeResponse> observer) {
        if (meError != null) {
            observer.onError(meError.asRuntimeException());
            return;
        }
        observer.onNext(GetMeResponse.newBuilder().setUser(me).build());
        observer.onCompleted();
    }

    @Override
    public void sync(SyncRequest request, StreamObserver<SyncResponse> observer) {
        syncCalls.incrementAndGet();
        if (syncError != null) {
            observer.onError(syncError.asRuntimeException());
            return;
        }
        for (MicrosoftChange change : changes) {
            observer.onNext(SyncResponse.newBuilder().setChange(change).build());
        }
        observer.onNext(SyncResponse.newBuilder().setResumeCursor(resumeCursor).build());
        observer.onCompleted();
    }

    static MicrosoftChange upsertFile(String id, String name) {
        DriveItem item = DriveItem.newBuilder()
                .setId(id)
                .setName(name)
                .setDriveId("drive-1")
                .setWebUrl("https://contoso.sharepoint.com/" + name)
                .setMimeType("text/plain")
                .setSize(12)
                .setCreatedAt(Timestamp.newBuilder().setSeconds(1_700_000_000))
                .setLastModifiedAt(Timestamp.newBuilder().setSeconds(1_700_000_100))
                .build();
        return MicrosoftChange.newBuilder()
                .setChangeId("c-" + id)
                .setOperation(ai.pipestream.microsoft.v1.ChangeOperation.CHANGE_OPERATION_UPSERT)
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId(id)
                        .setIngestedAt(Timestamp.newBuilder().setSeconds(1_700_000_200))
                        .setDriveItem(item))
                .build();
    }
}
