package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;

import java.util.List;

public final class CompositeMicrosoftChangeSink implements MicrosoftChangeSink {

    private final List<MicrosoftChangeSink> sinks;

    public CompositeMicrosoftChangeSink(List<MicrosoftChangeSink> sinks) {
        this.sinks = List.copyOf(sinks);
    }

    @Override
    public void emit(MicrosoftChange change) {
        for (MicrosoftChangeSink sink : sinks) {
            sink.emit(change);
        }
    }

    @Override
    public void snapshot(MicrosoftSnapshot snapshot) {
        for (MicrosoftChangeSink sink : sinks) {
            sink.snapshot(snapshot);
        }
    }

    @Override
    public void completeRun(String runId) {
        for (MicrosoftChangeSink sink : sinks) {
            sink.completeRun(runId);
        }
    }
}
