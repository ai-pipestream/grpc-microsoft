package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.LocalizedName;
import ai.pipestream.microsoft.v1.Term;
import ai.pipestream.microsoft.v1.TermGroup;
import ai.pipestream.microsoft.v1.TermGroupScope;
import ai.pipestream.microsoft.v1.TermLabel;
import ai.pipestream.microsoft.v1.TermSet;
import ai.pipestream.microsoft.v1.TermStore;
import com.fasterxml.jackson.databind.JsonNode;
import com.google.protobuf.Timestamp;
import java.util.function.Consumer;

/** Maps term store JSON to the contract's term store messages. */
final class TermStoreMapper {

    private TermStoreMapper() {
    }

    static TermStore toStore(JsonNode node) {
        TermStore.Builder store = TermStore.newBuilder()
                .setId(node.path("id").asText(""))
                .setDefaultLanguageTag(node.path("defaultLanguageTag").asText(""));
        for (JsonNode tag : node.path("languageTags")) {
            store.addLanguageTags(tag.asText(""));
        }
        return store.build();
    }

    static TermGroup toGroup(JsonNode node) {
        TermGroup.Builder group = TermGroup.newBuilder()
                .setId(node.path("id").asText(""))
                .setDisplayName(node.path("displayName").asText(""))
                .setDescription(node.path("description").asText(""))
                .setScope(scope(node.path("scope").asText("")));
        time(node, "createdDateTime", group::setCreatedAt);
        return group.build();
    }

    static TermSet toSet(JsonNode node, String groupId) {
        TermSet.Builder set = TermSet.newBuilder()
                .setId(node.path("id").asText(""))
                .setGroupId(groupId)
                .setDescription(node.path("description").asText(""));
        for (JsonNode name : node.path("localizedNames")) {
            set.addLocalizedNames(LocalizedName.newBuilder()
                    .setLanguageTag(name.path("languageTag").asText(""))
                    .setName(name.path("name").asText("")));
        }
        time(node, "createdDateTime", set::setCreatedAt);
        return set.build();
    }

    static Term toTerm(JsonNode node, String setId, String parentTermId) {
        Term.Builder term = Term.newBuilder()
                .setId(node.path("id").asText(""))
                .setSetId(setId)
                .setParentTermId(parentTermId);
        for (JsonNode label : node.path("labels")) {
            term.addLabels(TermLabel.newBuilder()
                    .setLanguageTag(label.path("languageTag").asText(""))
                    .setName(label.path("name").asText(""))
                    .setIsDefault(label.path("isDefault").asBoolean(false)));
        }
        for (JsonNode description : node.path("descriptions")) {
            term.addDescriptions(LocalizedName.newBuilder()
                    .setLanguageTag(description.path("languageTag").asText(""))
                    .setName(description.path("description").asText("")));
        }
        time(node, "createdDateTime", term::setCreatedAt);
        time(node, "lastModifiedDateTime", term::setLastModifiedAt);
        return term.build();
    }

    private static TermGroupScope scope(String scope) {
        return switch (scope) {
            case "global" -> TermGroupScope.TERM_GROUP_SCOPE_GLOBAL;
            case "system" -> TermGroupScope.TERM_GROUP_SCOPE_SYSTEM;
            case "siteCollection" -> TermGroupScope.TERM_GROUP_SCOPE_SITE_COLLECTION;
            default -> TermGroupScope.TERM_GROUP_SCOPE_UNSPECIFIED;
        };
    }

    private static void time(JsonNode node, String field, Consumer<Timestamp> setter) {
        Timestamp time = MicrosoftMapper.timestamp(node.path(field).asText(""));
        if (time != null) {
            setter.accept(time);
        }
    }
}
