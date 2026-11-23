package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ColumnDefinition;
import ai.pipestream.microsoft.v1.ColumnSpec;
import ai.pipestream.microsoft.v1.ColumnType;
import ai.pipestream.microsoft.v1.DateTimeColumnFormat;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.util.ArrayList;
import java.util.List;

/** Turns {@link ColumnSpec}s into Graph {@code columnDefinition} bodies and compares them. */
final class ColumnSpecs {

    private ColumnSpecs() {
    }

    /**
     * The Graph create body for a column.
     *
     * @param spec the column
     * @param siteId site whose term store a term column binds to
     * @return the {@code columnDefinition} JSON
     * @throws IllegalArgumentException when the spec has no name or an unsupported type
     */
    static ObjectNode toGraph(ColumnSpec spec, String siteId) {
        if (spec.getName().isBlank()) {
            throw new IllegalArgumentException("a column needs a name");
        }
        ObjectNode column = GraphClient.object();
        column.put("name", spec.getName());
        column.put("displayName",
                spec.getDisplayName().isBlank() ? spec.getName() : spec.getDisplayName());
        if (!spec.getDescription().isEmpty()) {
            column.put("description", spec.getDescription());
        }
        if (!spec.getColumnGroup().isEmpty()) {
            column.put("columnGroup", spec.getColumnGroup());
        }
        if (spec.getIndexed()) {
            column.put("indexed", true);
        }
        if (spec.getHidden()) {
            column.put("hidden", true);
        }
        if (spec.getRequired()) {
            column.put("required", true);
        }
        switch (spec.getType()) {
            case COLUMN_TYPE_TEXT -> {
                ObjectNode text = column.putObject("text").put("allowMultipleLines", false);
                if (spec.getMaxLength() > 0) {
                    text.put("maxLength", spec.getMaxLength());
                }
            }
            case COLUMN_TYPE_NOTE -> column.putObject("text").put("allowMultipleLines", true);
            case COLUMN_TYPE_CHOICE, COLUMN_TYPE_MULTICHOICE -> {
                ObjectNode choice = column.putObject("choice")
                        .put("allowTextEntry", spec.getAllowTextEntry())
                        .put("displayAs", spec.getType() == ColumnType.COLUMN_TYPE_MULTICHOICE
                                ? "checkBoxes" : "dropDownMenu");
                ArrayNode choices = choice.putArray("choices");
                spec.getChoicesList().forEach(choices::add);
            }
            case COLUMN_TYPE_NUMBER -> column.putObject("number");
            case COLUMN_TYPE_CURRENCY -> column.putObject("currency");
            case COLUMN_TYPE_DATE_TIME -> {
                ObjectNode dateTime = column.putObject("dateTime");
                if (spec.getDateTimeFormat()
                        == DateTimeColumnFormat.DATE_TIME_COLUMN_FORMAT_DATE_ONLY) {
                    dateTime.put("format", "dateOnly");
                } else if (spec.getDateTimeFormat()
                        == DateTimeColumnFormat.DATE_TIME_COLUMN_FORMAT_DATE_TIME) {
                    dateTime.put("format", "dateTime");
                }
            }
            case COLUMN_TYPE_LOOKUP -> {
                if (spec.getLookupListId().isBlank()) {
                    throw new IllegalArgumentException(
                            "lookup column " + spec.getName() + " needs lookup_list_id");
                }
                column.putObject("lookup")
                        .put("listId", spec.getLookupListId())
                        .put("columnName", spec.getLookupColumnName().isBlank() ? "Title"
                                : spec.getLookupColumnName())
                        .put("allowMultipleValues", spec.getMultiValued());
            }
            case COLUMN_TYPE_BOOLEAN -> column.putObject("boolean");
            case COLUMN_TYPE_USER -> column.putObject("personOrGroup")
                    .put("allowMultipleSelection", spec.getMultiValued());
            case COLUMN_TYPE_URL -> column.putObject("hyperlinkOrPicture").put("isPicture", false);
            case COLUMN_TYPE_TERM, COLUMN_TYPE_MULTITERM -> {
                if (spec.getTermSetId().isBlank()) {
                    throw new IllegalArgumentException(
                            "term column " + spec.getName() + " needs term_set_id");
                }
                column.putObject("term")
                        .put("allowMultipleValues",
                                spec.getType() == ColumnType.COLUMN_TYPE_MULTITERM)
                        .put("termSet@odata.bind", "sites/" + GraphPaths.site(siteId)
                                + "/termStore/sets/"
                                + GraphPaths.segment(spec.getTermSetId(), "term_set_id"));
            }
            default -> throw new IllegalArgumentException("column " + spec.getName()
                    + " has a type that cannot be created: " + spec.getType());
        }
        return column;
    }

    /**
     * Whether Graph refuses to create the column, so SharePoint's own API creates it.
     *
     * @param spec the column
     * @return true for a hyperlink or term column
     */
    static boolean needsSharePoint(ColumnSpec spec) {
        return spec.getType() == ColumnType.COLUMN_TYPE_URL || isTerm(spec);
    }

    /**
     * Whether the column holds terms.
     *
     * @param spec the column
     * @return true for a term or multi term column
     */
    static boolean isTerm(ColumnSpec spec) {
        return spec.getType() == ColumnType.COLUMN_TYPE_TERM
                || spec.getType() == ColumnType.COLUMN_TYPE_MULTITERM;
    }

    /**
     * The SharePoint {@code CreateFieldAsXml} body for a hyperlink or term column.
     *
     * @param spec the column
     * @param termStoreId the term store of a term column; unused for a hyperlink
     * @param noteColumnId id of the hidden note column of a term column; unused for a hyperlink
     * @return the request JSON
     */
    static ObjectNode toSharePoint(ColumnSpec spec, String termStoreId, String noteColumnId) {
        boolean multi = spec.getType() == ColumnType.COLUMN_TYPE_MULTITERM;
        StringBuilder xml = new StringBuilder(!isTerm(spec)
                ? "<Field Type=\"URL\" Format=\"Hyperlink\""
                : multi ? "<Field Type=\"TaxonomyFieldTypeMulti\" Mult=\"TRUE\" ShowField=\"Term1033\""
                        : "<Field Type=\"TaxonomyFieldType\" ShowField=\"Term1033\"");
        attribute(xml, "Name", spec.getName());
        attribute(xml, "StaticName", spec.getName());
        attribute(xml, "DisplayName",
                spec.getDisplayName().isBlank() ? spec.getName() : spec.getDisplayName());
        attribute(xml, "Description", spec.getDescription());
        attribute(xml, "Group", spec.getColumnGroup());
        attribute(xml, "Indexed", spec.getIndexed() ? "TRUE" : "");
        attribute(xml, "Hidden", spec.getHidden() ? "TRUE" : "");
        attribute(xml, "Required", spec.getRequired() ? "TRUE" : "");
        if (isTerm(spec)) {
            xml.append("><Customization><ArrayOfProperty>");
            property(xml, "SspId", termStoreId);
            property(xml, "TermSetId", spec.getTermSetId());
            property(xml, "TextField", "{" + noteColumnId + "}");
            xml.append("</ArrayOfProperty></Customization></Field>");
        } else {
            xml.append("/>");
        }
        return createField(xml);
    }

    /**
     * The SharePoint {@code CreateFieldAsXml} body for the hidden note column that holds
     * a term column's values.
     *
     * @param spec the term column
     * @return the request JSON
     */
    static ObjectNode toSharePointTermNote(ColumnSpec spec) {
        StringBuilder xml = new StringBuilder("<Field Type=\"Note\" Hidden=\"TRUE\""
                + " CanToggleHidden=\"TRUE\" ShowInViewForms=\"FALSE\"");
        attribute(xml, "Name", spec.getName() + "_0");
        attribute(xml, "StaticName", spec.getName() + "_0");
        attribute(xml, "DisplayName",
                (spec.getDisplayName().isBlank() ? spec.getName() : spec.getDisplayName()) + "_0");
        attribute(xml, "Group", spec.getColumnGroup());
        xml.append("/>");
        return createField(xml);
    }

    private static ObjectNode createField(StringBuilder xml) {
        ObjectNode body = GraphClient.object();
        // Options 8 keeps Name as the internal name.
        body.putObject("parameters").put("SchemaXml", xml.toString()).put("Options", 8);
        return body;
    }

    private static void property(StringBuilder xml, String name, String value) {
        xml.append("<Property><Name>").append(name).append("</Name>")
                .append("<Value xmlns:q1=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q1:string\"")
                .append(" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">")
                .append(escape(value)).append("</Value></Property>");
    }

    private static String escape(String value) {
        return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace("\"", "&quot;");
    }

    private static void attribute(StringBuilder xml, String name, String value) {
        if (value.isEmpty()) {
            return;
        }
        xml.append(' ').append(name).append("=\"").append(escape(value)).append('"');
    }

    /**
     * The spec fields an existing column differs in.
     *
     * @param spec the wanted column
     * @param actual the column as it exists
     * @return the differing {@code ColumnSpec} field names; empty when it matches
     */
    static List<String> drift(ColumnSpec spec, ColumnDefinition actual) {
        List<String> drift = new ArrayList<>();
        if (type(actual) != spec.getType()) {
            drift.add("type");
            return drift;
        }
        if (actual.hasChoice()
                && !actual.getChoice().getChoicesList().equals(spec.getChoicesList())) {
            drift.add("choices");
        }
        if (actual.hasLookup() && !spec.getLookupListId().isBlank()
                && !actual.getLookup().getListId().equalsIgnoreCase(spec.getLookupListId())) {
            drift.add("lookup_list_id");
        }
        return drift;
    }

    private static ColumnType type(ColumnDefinition column) {
        return switch (column.getKindCase()) {
            case TEXT -> column.getText().getAllowMultipleLines() ? ColumnType.COLUMN_TYPE_NOTE
                    : ColumnType.COLUMN_TYPE_TEXT;
            case CHOICE -> column.getChoice().getDisplayAs().equalsIgnoreCase("checkBoxes")
                    ? ColumnType.COLUMN_TYPE_MULTICHOICE : ColumnType.COLUMN_TYPE_CHOICE;
            case LOOKUP -> ColumnType.COLUMN_TYPE_LOOKUP;
            case PERSON_OR_GROUP -> ColumnType.COLUMN_TYPE_USER;
            case DATE_TIME -> ColumnType.COLUMN_TYPE_DATE_TIME;
            case NUMBER -> ColumnType.COLUMN_TYPE_NUMBER;
            case CURRENCY -> ColumnType.COLUMN_TYPE_CURRENCY;
            case CALCULATED -> ColumnType.COLUMN_TYPE_CALCULATED;
            case HYPERLINK_OR_PICTURE -> ColumnType.COLUMN_TYPE_URL;
            case TERM -> column.getTerm().getAllowMultipleValues()
                    ? ColumnType.COLUMN_TYPE_MULTITERM : ColumnType.COLUMN_TYPE_TERM;
            case BOOLEAN -> ColumnType.COLUMN_TYPE_BOOLEAN;
            case GEOLOCATION -> ColumnType.COLUMN_TYPE_GEOLOCATION;
            case THUMBNAIL -> ColumnType.COLUMN_TYPE_THUMBNAIL;
            case CONTENT_APPROVAL_STATUS -> ColumnType.COLUMN_TYPE_APPROVAL_STATUS;
            case KIND_NOT_SET -> column.getType();
        };
    }
}
