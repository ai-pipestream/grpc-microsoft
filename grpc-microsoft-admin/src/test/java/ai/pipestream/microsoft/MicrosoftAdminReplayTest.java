package ai.pipestream.microsoft;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.DynamicTest;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestFactory;

/**
 * Runs {@link AdminGrantScenarios} against the exchanges
 * {@link MicrosoftAdminLiveIT} recorded from the real tenant. No network: any
 * request that is not the next recorded one fails its scenario. A scenario
 * with no recording is skipped.
 */
class MicrosoftAdminReplayTest {

    @TestFactory
    List<DynamicTest> scenarios() throws Exception {
        Path recordings = AdminGrantScenarios.dir();
        List<DynamicTest> tests = new ArrayList<>();
        for (AdminGrantScenarios.Replayed replayed : AdminGrantScenarios.replay(recordings)) {
            tests.add(DynamicTest.dynamicTest(replayed.scenario().id(), () -> {
                assumeTrue(replayed.recorded(), "no recording of " + replayed.scenario().id()
                        + " yet under " + recordings + "; record one with liveAdminTest");
                assertThat(replayed.problems()).as("requests that differ from the recording").isEmpty();
                assertThat(replayed.outcome().status()).as(replayed.outcome().detail())
                        .isEqualTo(AdminGrantScenarios.Status.PASS);
            }));
        }
        return tests;
    }

    @Test
    void noRecordingHoldsTenantData() throws Exception {
        assertThat(AdminGrantScenarios.violations(AdminGrantScenarios.dir())).isEmpty();
    }
}
