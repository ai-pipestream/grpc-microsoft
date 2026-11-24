package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftEntity.EntityCase;
import org.junit.jupiter.api.Test;

import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

class RunMirrorTest {

    private static final Set<EntityCase> DRIVES = Set.of(EntityCase.DRIVE, EntityCase.DRIVE_ITEM);

    @Test
    void aRunReplacesOnlyTheKindsItEnumerated() {
        RunMirror<String> mirror = new RunMirror<>();
        mirror.put("items/a", EntityCase.DRIVE_ITEM, "a");
        mirror.put("items/b", EntityCase.DRIVE_ITEM, "b");
        mirror.put("lists/l", EntityCase.LIST, "l");
        mirror.put("lists/l/1", EntityCase.LIST_ITEM, "l1");
        mirror.complete(Set.of(EntityCase.DRIVE_ITEM, EntityCase.LIST, EntityCase.LIST_ITEM));

        // The next pass sweeps drives only, and file b is gone.
        mirror.put("items/a", EntityCase.DRIVE_ITEM, "a2");
        RunMirror.Completed<String> completed = mirror.complete(DRIVES);

        assertThat(completed.increment()).isFalse();
        assertThat(completed.values()).containsExactlyInAnyOrder("a2", "l", "l1");
    }

    @Test
    void aRunThatEnumeratedAKindDropsWhatItDidNotEmitOfThatKind() {
        RunMirror<String> mirror = new RunMirror<>();
        mirror.put("lists/l", EntityCase.LIST, "l");
        mirror.put("lists/m", EntityCase.LIST, "m");
        mirror.complete(Set.of(EntityCase.LIST));

        mirror.put("lists/m", EntityCase.LIST, "m2");

        assertThat(mirror.complete(Set.of(EntityCase.LIST)).values()).containsExactly("m2");
    }

    @Test
    void aRunThatEnumeratedNothingMergesIntoTheMirror() {
        RunMirror<String> mirror = new RunMirror<>();
        mirror.put("items/a", EntityCase.DRIVE_ITEM, "a");
        mirror.put("items/b", EntityCase.DRIVE_ITEM, "b");
        mirror.complete(DRIVES);

        mirror.put("items/b", EntityCase.DRIVE_ITEM, "b-deleted");
        mirror.put("items/c", EntityCase.DRIVE_ITEM, "c");
        RunMirror.Completed<String> completed = mirror.complete(Set.of());

        assertThat(completed.increment()).isFalse();
        assertThat(completed.values()).containsExactly("a", "b-deleted", "c");
    }

    @Test
    void beforeAnythingWasEnumeratedARunStandsAloneAsAnIncrement() {
        RunMirror<String> mirror = new RunMirror<>();
        mirror.put("items/a", EntityCase.DRIVE_ITEM, "a");

        RunMirror.Completed<String> first = mirror.complete(Set.of());
        assertThat(first.increment()).isTrue();
        assertThat(first.values()).containsExactly("a");

        // The increment was not kept as a mirror: the next one stands alone too.
        mirror.put("items/b", EntityCase.DRIVE_ITEM, "b");
        RunMirror.Completed<String> second = mirror.complete(Set.of());
        assertThat(second.increment()).isTrue();
        assertThat(second.values()).containsExactly("b");
    }
}
