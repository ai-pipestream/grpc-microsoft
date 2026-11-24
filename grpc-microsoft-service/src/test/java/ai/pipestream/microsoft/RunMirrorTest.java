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

    @Test
    void aFullRunOfOneConnectionKeepsTheEntriesOfAnother() {
        RunMirror<String> mirror = new RunMirror<>();
        RunMirror<String>.Pass x = mirror.pass("x");
        RunMirror<String>.Pass y = mirror.pass("y");
        x.put("x/1", EntityCase.DRIVE_ITEM, "x1");
        x.complete(MicrosoftChangeSink.ALL_KINDS);
        y.put("y/1", EntityCase.DRIVE_ITEM, "y1");

        RunMirror.Completed<String> completed = y.complete(MicrosoftChangeSink.ALL_KINDS);

        assertThat(completed.increment()).isFalse();
        assertThat(completed.values()).containsExactly("x1", "y1");
    }

    @Test
    void aFullRunDropsOnlyItsOwnConnectionsUnseenEntries() {
        RunMirror<String> mirror = new RunMirror<>();
        RunMirror<String>.Pass x = mirror.pass("x");
        RunMirror<String>.Pass y = mirror.pass("y");
        x.put("x/1", EntityCase.DRIVE_ITEM, "x1");
        x.put("x/2", EntityCase.DRIVE_ITEM, "x2");
        x.put("x/l", EntityCase.LIST_ITEM, "x-list");
        x.complete(MicrosoftChangeSink.ALL_KINDS);
        y.put("y/1", EntityCase.DRIVE_ITEM, "y1");
        y.complete(MicrosoftChangeSink.ALL_KINDS);
        // x/2 is gone; this run of x did not sweep lists.
        x.put("x/1", EntityCase.DRIVE_ITEM, "x1b");

        RunMirror.Completed<String> completed =
                x.complete(Set.of(EntityCase.DRIVE_ITEM));

        assertThat(completed.values()).containsExactlyInAnyOrder("x1b", "x-list", "y1");
    }

    @Test
    void aBaselineOfOneConnectionIsNotABaselineForAnother() {
        RunMirror<String> mirror = new RunMirror<>();
        RunMirror<String>.Pass x = mirror.pass("x");
        RunMirror<String>.Pass y = mirror.pass("y");
        x.put("x/1", EntityCase.DRIVE_ITEM, "x1");
        x.complete(MicrosoftChangeSink.ALL_KINDS);
        y.put("y/1", EntityCase.DRIVE_ITEM, "y1");

        // y has never been enumerated: its incremental run is not the
        // connection's bundle, so it is written apart.
        RunMirror.Completed<String> completed = y.complete(Set.of());

        assertThat(completed.increment()).isTrue();
        assertThat(completed.values()).containsExactly("y1");
        assertThat(x.complete(Set.of()).values()).containsExactly("x1");
    }

    @Test
    void twoPassesAtOnceCompleteWithTheirOwnValues() {
        RunMirror<String> mirror = new RunMirror<>();
        RunMirror<String>.Pass x = mirror.pass("x");
        RunMirror<String>.Pass alsoX = mirror.pass("x");
        x.put("a", EntityCase.DRIVE_ITEM, "a");
        alsoX.put("b", EntityCase.DRIVE_ITEM, "b");

        assertThat(x.complete(Set.of()).values()).containsExactly("a");
        assertThat(alsoX.complete(Set.of()).values()).containsExactly("b");
    }

    @Test
    void aDriveReadAgainReplacesItsOwnEntriesAndKeepsEveryOtherDrives() {
        RunMirror<String> mirror = new RunMirror<>();
        RunMirror<String>.Pass first = mirror.pass("x");
        first.put("drives/a", EntityCase.DRIVE, "", "drive-a");
        first.put("items/a/1", EntityCase.DRIVE_ITEM, "a", "a1");
        first.put("items/a/2", EntityCase.DRIVE_ITEM, "a", "a2");
        first.put("items/b/1", EntityCase.DRIVE_ITEM, "b", "b1");
        first.put("lists/s/l/1", EntityCase.LIST_ITEM, "a", "list-item-with-the-same-id");
        first.complete(MicrosoftChangeSink.ALL_KINDS);

        // Drive a was read again from scratch and a2 is gone; b had no changes.
        RunMirror<String>.Pass second = mirror.pass("x");
        second.put("items/a/1", EntityCase.DRIVE_ITEM, "a", "a1-again");
        RunMirror.Completed<String> completed =
                second.complete(Set.of(), Set.of(SweepScope.drive("a")));

        assertThat(completed.increment()).isFalse();
        assertThat(completed.values()).containsExactlyInAnyOrder("drive-a", "a1-again", "b1",
                "list-item-with-the-same-id");
    }

    @Test
    void aDriveReadAgainByOneConnectionKeepsAnotherConnectionsEntriesOfThatDrive() {
        RunMirror<String> mirror = new RunMirror<>();
        RunMirror<String>.Pass x = mirror.pass("x");
        x.put("x/items/a/1", EntityCase.DRIVE_ITEM, "a", "x-a1");
        x.complete(MicrosoftChangeSink.ALL_KINDS);
        RunMirror<String>.Pass y = mirror.pass("y");
        y.put("y/items/a/1", EntityCase.DRIVE_ITEM, "a", "y-a1");
        y.complete(MicrosoftChangeSink.ALL_KINDS);

        RunMirror.Completed<String> completed =
                mirror.pass("x").complete(Set.of(), Set.of(SweepScope.drive("a")));

        assertThat(completed.values()).containsExactly("y-a1");
    }

    @Test
    void aDriveReadInFullBeforeAnyBaselineStillStandsAloneAsAnIncrement() {
        RunMirror<String> mirror = new RunMirror<>();
        RunMirror<String>.Pass pass = mirror.pass("x");
        pass.put("items/a/1", EntityCase.DRIVE_ITEM, "a", "a1");

        RunMirror.Completed<String> completed =
                pass.complete(Set.of(), Set.of(SweepScope.drive("a")));

        assertThat(completed.increment()).isTrue();
        assertThat(completed.values()).containsExactly("a1");
    }

    @Test
    void theUnnamedPassIsTheProcessDefaultConnection() {
        RunMirror<String> mirror = new RunMirror<>();
        mirror.put("a", EntityCase.DRIVE_ITEM, "a");
        mirror.complete(MicrosoftChangeSink.ALL_KINDS);
        RunMirror<String>.Pass unnamed = mirror.pass(null);

        assertThat(unnamed.connectionId()).isEmpty();
        // Same owner: its full run replaces what the default pass wrote.
        unnamed.put("b", EntityCase.DRIVE_ITEM, "b");
        assertThat(unnamed.complete(MicrosoftChangeSink.ALL_KINDS).values()).containsExactly("b");
    }
}
