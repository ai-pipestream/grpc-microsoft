package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.DeltaPreferences;
import ai.pipestream.microsoft.v1.SyncCursor;
import org.junit.jupiter.api.Test;

import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class GraphHelpersTest {

    private static final String BASE = "https://graph.microsoft.com/v1.0";

    @Test
    void deltaLinksMustStayOnTheGraphHostAndRoute() {
        for (String ok : new String[] {
                BASE + "/drives/d1/root/delta?token=abc",
                BASE + "/drives/d1/root/delta(token='abc')",
                BASE + "/drives/d1/root/microsoft.graph.delta()?$deltatoken=abc",
                "https://GRAPH.microsoft.com:443/v1.0/Drives/d1/root/delta?token=x"}) {
            assertThat(GraphLinks.requireDeltaLink(BASE, ok, "drives", "d1", "root")).isEqualTo(ok);
        }
        for (String bad : new String[] {
                "https://evil.example.com/v1.0/drives/d1/root/delta?token=abc",
                "http://graph.microsoft.com/v1.0/drives/d1/root/delta?token=abc",
                "https://graph.microsoft.com:8443/v1.0/drives/d1/root/delta",
                "https://user@graph.microsoft.com/v1.0/drives/d1/root/delta",
                "https://graph.microsoft.com/beta/drives/d1/root/delta",
                BASE + "/drives/d2/root/delta?token=abc",
                BASE + "/drives/d1/root/children",
                BASE + "/drives/d1/root/x/delta",
                "/drives/d1/root/delta",
                "not a url"}) {
            assertThatThrownBy(() -> GraphLinks.requireDeltaLink(BASE, bad, "drives", "d1", "root"))
                    .as(bad).isInstanceOf(IllegalArgumentException.class);
        }
        assertThat(GraphLinks.isUnderBase(BASE + "/", BASE + "/me")).isTrue();
        assertThat(GraphLinks.isUnderBase(BASE, "https://graph.microsoft.com/v1.0evil/me")).isFalse();
    }

    @Test
    void syncCursorRoundTripsAndRejectsForeignLinks() {
        SyncCursor cursor = SyncCursor.newBuilder()
                .putDriveDeltaLinks("d1", BASE + "/drives/d1/root/delta?token=a")
                .putListDeltaLinks(SyncCursors.listKey("host,s1,w1", "l1"),
                        BASE + "/sites/host,s1,w1/lists/l1/items/delta?token=b")
                .setSitesDeltaLink(BASE + "/sites/delta?token=c")
                .build();
        String encoded = SyncCursors.encode(cursor);
        assertThat(encoded).doesNotContain("=", "+", "/");
        assertThat(SyncCursors.decode(encoded, BASE)).isEqualTo(cursor);
        assertThat(SyncCursors.decode("", BASE)).isEqualTo(SyncCursor.getDefaultInstance());

        String foreign = SyncCursors.encode(SyncCursor.newBuilder()
                .putDriveDeltaLinks("d1", "https://evil.example.com/v1.0/drives/d1/root/delta").build());
        assertThatThrownBy(() -> SyncCursors.decode(foreign, BASE))
                .isInstanceOf(IllegalArgumentException.class);
        String swapped = SyncCursors.encode(SyncCursor.newBuilder()
                .putDriveDeltaLinks("d1", BASE + "/drives/d2/root/delta?token=a").build());
        assertThatThrownBy(() -> SyncCursors.decode(swapped, BASE))
                .isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> SyncCursors.decode("%%%", BASE))
                .isInstanceOf(IllegalArgumentException.class);
    }

    @Test
    void preferHeaderFollowsTypedPreferences() {
        assertThat(DeltaHeaders.prefer(DeltaPreferences.getDefaultInstance())).isEmpty();
        assertThat(DeltaHeaders.prefer(DeltaPreferences.newBuilder()
                .setHierarchicalSharing(true).setShowRemovedAsDeleted(true)
                .setTraversePermissionGaps(true).setShowSharingChanges(true)
                .setExcludeParent(true).build()))
                .isEqualTo(Map.of("Prefer", "hierarchicalsharing, deltashowremovedasdeleted, "
                        + "deltatraversepermissiongaps, deltashowsharingchanges, deltaExcludeParent"));
        assertThatThrownBy(() -> DeltaHeaders.prefer(DeltaPreferences.newBuilder()
                .setShowSharingChanges(true).build()))
                .isInstanceOf(IllegalArgumentException.class);
    }

    @Test
    void pathSegmentsAreEncodedAndTraversalRejected() {
        assertThat(GraphPaths.segment("a b/c", "id")).isEqualTo("a%20b%2Fc");
        assertThatThrownBy(() -> GraphPaths.segment("..", "id")).isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> GraphPaths.segment(" ", "id")).isInstanceOf(IllegalArgumentException.class);
        assertThat(GraphPaths.encodePath("\\Shared//./a b/r+1.docx/")).isEqualTo("/Shared/a%20b/r%2B1.docx");
        assertThat(GraphPaths.encodePath("/")).isEmpty();
        assertThatThrownBy(() -> GraphPaths.encodePath("/a/../b")).isInstanceOf(IllegalArgumentException.class);
    }
}
