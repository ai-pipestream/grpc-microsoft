package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class CompositeMicrosoftChangeSinkTest {

    @Test
    void fansOutEmitSnapshotAndCompleteRun() {
        InMemoryMicrosoftChangeSink first = new InMemoryMicrosoftChangeSink();
        InMemoryMicrosoftChangeSink second = new InMemoryMicrosoftChangeSink();
        CompositeMicrosoftChangeSink composite =
                new CompositeMicrosoftChangeSink(List.of(first, second));
        MicrosoftChange change = MicrosoftChange.newBuilder()
                .setChangeId("c1")
                .setOperation(ChangeOperation.CHANGE_OPERATION_DELETE)
                .build();
        MicrosoftSnapshot snapshot = MicrosoftSnapshot.newBuilder().setSnapshotId("s1").build();
        composite.emit(change);
        composite.snapshot(snapshot);
        composite.completeRun("run-9");
        assertThat(first.changes()).containsExactly(change);
        assertThat(second.snapshots()).containsExactly(snapshot);
        assertThat(first.completedRuns()).containsExactly("run-9");
        assertThat(second.completedRuns()).containsExactly("run-9");
    }

    @Test
    void earlyFailureSkipsLaterSinks() {
        MicrosoftChangeSink failing = new MicrosoftChangeSink() {
            @Override
            public void emit(MicrosoftChange change) {
                throw new IllegalStateException("boom");
            }

            @Override
            public void snapshot(MicrosoftSnapshot snapshot) {
            }
        };
        InMemoryMicrosoftChangeSink later = new InMemoryMicrosoftChangeSink();
        CompositeMicrosoftChangeSink composite =
                new CompositeMicrosoftChangeSink(List.of(failing, later));
        assertThatThrownBy(() -> composite.emit(MicrosoftChange.getDefaultInstance()))
                .isInstanceOf(IllegalStateException.class);
        assertThat(later.changes()).isEmpty();
    }
}
