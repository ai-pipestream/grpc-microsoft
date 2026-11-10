package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.FieldOutcome;
import ai.pipestream.microsoft.v1.FieldValue;
import ai.pipestream.microsoft.v1.FieldWriteState;
import ai.pipestream.microsoft.v1.PersonValue;
import ai.pipestream.microsoft.v1.TermValue;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.google.protobuf.Timestamp;

import java.io.IOException;
import java.time.Instant;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/** Turns typed {@link FieldValue}s into the Graph {@code fields} object of a list item. */
final class ListFieldWriter {

    /** Looks up what a value needs from the list it is written to. */
    interface Lookups {

        /**
         * The user information list id of a site user.
         *
         * @param email the user's email address
         * @return the lookup id
         * @throws IOException if Graph returns an error
         * @throws InterruptedException if the HTTP call is interrupted
         */
        int personId(String email) throws IOException, InterruptedException;

        /**
         * The hidden note column behind a term column.
         *
         * @param termColumn internal name of the term column
         * @return internal name of its note column
         * @throws IOException if Graph returns an error
         * @throws InterruptedException if the HTTP call is interrupted
         */
        String termNoteColumn(String termColumn) throws IOException, InterruptedException;
    }

    /**
     * A translated patch.
     *
     * @param fields the Graph {@code fields} object
     * @param outcomes one outcome per value, in patch order
     */
    record Patch(ObjectNode fields, List<FieldOutcome> outcomes) {
    }

    private ListFieldWriter() {
    }

    /**
     * Translates column values.
     *
     * @param values the values
     * @param lookups resolves people and term columns
     * @param state the state to report for every column
     * @return the Graph fields and the outcomes
     * @throws IOException if a lookup fails
     * @throws InterruptedException if a lookup is interrupted
     * @throws IllegalArgumentException when a value is malformed or a name repeats
     */
    static Patch translate(List<FieldValue> values, Lookups lookups, FieldWriteState state)
            throws IOException, InterruptedException {
        ObjectNode fields = GraphClient.object();
        List<FieldOutcome> outcomes = new ArrayList<>();
        Set<String> seen = new HashSet<>();
        for (FieldValue value : values) {
            String name = value.getName();
            if (name.isBlank()) {
                throw new IllegalArgumentException("a field value needs a column name");
            }
            if (!seen.add(name)) {
                throw new IllegalArgumentException("column " + name + " is named twice");
            }
            String key = write(fields, value, lookups);
            outcomes.add(FieldOutcome.newBuilder().setName(name).setState(state).setGraphField(key)
                    .build());
        }
        return new Patch(fields, List.copyOf(outcomes));
    }

    private static String write(ObjectNode fields, FieldValue value, Lookups lookups)
            throws IOException, InterruptedException {
        String name = value.getName();
        String lookupKey = name + "LookupId";
        switch (value.getValueCase()) {
            case TEXT -> fields.put(name, value.getText());
            case NOTE -> fields.put(name, value.getNote());
            case NUMBER -> fields.put(name, value.getNumber());
            case BOOLEAN -> fields.put(name, value.getBoolean());
            case DATE_TIME -> fields.put(name, instant(value.getDateTime()));
            case CHOICE -> fields.put(name, value.getChoice());
            case MULTI_CHOICE -> {
                fields.put(name + "@odata.type", "Collection(Edm.String)");
                ArrayNode array = fields.putArray(name);
                value.getMultiChoice().getChoicesList().forEach(array::add);
            }
            case TERM -> {
                String note = noteColumn(value, lookups);
                fields.put(note, term(value.getTerm()));
                return note;
            }
            case MULTI_TERM -> {
                String note = noteColumn(value, lookups);
                List<String> terms = new ArrayList<>();
                for (TermValue term : value.getMultiTerm().getTermsList()) {
                    terms.add(term(term));
                }
                fields.put(note, String.join(";#", terms));
                return note;
            }
            case PERSON -> {
                fields.put(lookupKey, person(value.getPerson(), lookups));
                return lookupKey;
            }
            case MULTI_PERSON -> {
                fields.put(lookupKey + "@odata.type", "Collection(Edm.Int32)");
                ArrayNode array = fields.putArray(lookupKey);
                for (PersonValue person : value.getMultiPerson().getPeopleList()) {
                    array.add(person(person, lookups));
                }
                return lookupKey;
            }
            case HYPERLINK -> {
                String url = value.getHyperlink().getUrl();
                String description = value.getHyperlink().getDescription();
                fields.putObject(name).put("Url", url)
                        .put("Description", description.isEmpty() ? url : description);
            }
            case LOOKUP -> {
                fields.put(lookupKey, value.getLookup());
                return lookupKey;
            }
            case MULTI_LOOKUP -> {
                fields.put(lookupKey + "@odata.type", "Collection(Edm.Int32)");
                ArrayNode array = fields.putArray(lookupKey);
                value.getMultiLookup().getLookupIdsList().forEach(array::add);
                return lookupKey;
            }
            case CLEAR -> fields.putNull(name);
            case VALUE_NOT_SET ->
                throw new IllegalArgumentException("column " + name + " has no value");
        }
        return name;
    }

    private static String noteColumn(FieldValue value, Lookups lookups)
            throws IOException, InterruptedException {
        return value.getTermNoteColumn().isBlank() ? lookups.termNoteColumn(value.getName())
                : value.getTermNoteColumn();
    }

    private static String term(TermValue term) {
        if (term.getTermId().isBlank() || term.getLabel().isBlank()) {
            throw new IllegalArgumentException("a term needs term_id and label");
        }
        return (term.hasWssId() ? term.getWssId() : -1) + ";#" + term.getLabel() + "|"
                + term.getTermId();
    }

    private static int person(PersonValue person, Lookups lookups)
            throws IOException, InterruptedException {
        return switch (person.getPrincipalCase()) {
            case LOOKUP_ID -> person.getLookupId();
            case EMAIL -> lookups.personId(person.getEmail());
            case PRINCIPAL_NOT_SET ->
                throw new IllegalArgumentException("a person needs lookup_id or email");
        };
    }

    private static String instant(Timestamp timestamp) {
        return Instant.ofEpochSecond(timestamp.getSeconds(), timestamp.getNanos()).toString();
    }
}
