package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.DeltaPreferences;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Turns the typed {@link DeltaPreferences} into Graph's {@code Prefer} request
 * header for {@code driveItem: delta}.
 */
public final class DeltaHeaders {

    /** {@code Prefer} value for {@link DeltaPreferences#getHierarchicalSharing()}. */
    public static final String HIERARCHICAL_SHARING = "hierarchicalsharing";
    /** {@code Prefer} value for {@link DeltaPreferences#getShowRemovedAsDeleted()}. */
    public static final String SHOW_REMOVED_AS_DELETED = "deltashowremovedasdeleted";
    /** {@code Prefer} value for {@link DeltaPreferences#getTraversePermissionGaps()}. */
    public static final String TRAVERSE_PERMISSION_GAPS = "deltatraversepermissiongaps";
    /** {@code Prefer} value for {@link DeltaPreferences#getShowSharingChanges()}. */
    public static final String SHOW_SHARING_CHANGES = "deltashowsharingchanges";
    /** {@code Prefer} value for {@link DeltaPreferences#getExcludeParent()}. */
    public static final String EXCLUDE_PARENT = "deltaExcludeParent";

    private DeltaHeaders() {
    }

    /**
     * The request headers for {@code prefs}.
     *
     * @param prefs the typed preferences
     * @return {@code Prefer: a, b, ...}, or no headers when nothing is set
     * @throws IllegalArgumentException when show_sharing_changes is set without
     *         show_removed_as_deleted and traverse_permission_gaps, which Graph
     *         requires alongside it
     */
    public static Map<String, String> prefer(DeltaPreferences prefs) {
        if (prefs.getShowSharingChanges()
                && (!prefs.getShowRemovedAsDeleted() || !prefs.getTraversePermissionGaps())) {
            throw new IllegalArgumentException("show_sharing_changes requires show_removed_as_deleted "
                    + "and traverse_permission_gaps (Graph rejects deltashowsharingchanges alone)");
        }
        List<String> values = new ArrayList<>();
        if (prefs.getHierarchicalSharing()) {
            values.add(HIERARCHICAL_SHARING);
        }
        if (prefs.getShowRemovedAsDeleted()) {
            values.add(SHOW_REMOVED_AS_DELETED);
        }
        if (prefs.getTraversePermissionGaps()) {
            values.add(TRAVERSE_PERMISSION_GAPS);
        }
        if (prefs.getShowSharingChanges()) {
            values.add(SHOW_SHARING_CHANGES);
        }
        if (prefs.getExcludeParent()) {
            values.add(EXCLUDE_PARENT);
        }
        return values.isEmpty() ? Map.of() : Map.of("Prefer", String.join(", ", values));
    }
}
