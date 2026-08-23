package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;

import java.util.List;

/**
 * Fans one emission out to several sinks, in list order. An exception from an
 * earlier sink skips the later ones and propagates to the caller: composing
 * sinks is wiring, not error policy, so each sink keeps its own failure
 * semantics.
 */
public final class CompositeMicrosoftChangeSink implements MicrosoftChangeSink {

    private final List<MicrosoftChangeSink> sinks;

    /**
     * Copies {@code sinks} and fans {@link #emit}, {@link #snapshot}, and
     * {@link #completeRun} out in that order.
     *
     * @param sinks destinations, in call order
     */
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
