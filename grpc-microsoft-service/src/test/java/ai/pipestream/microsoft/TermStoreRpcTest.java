package ai.pipestream.microsoft;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import ai.pipestream.microsoft.v1.GetTermStoreRequest;
import ai.pipestream.microsoft.v1.ListTermGroupsRequest;
import ai.pipestream.microsoft.v1.ListTermSetsRequest;
import ai.pipestream.microsoft.v1.ListTermsRequest;
import ai.pipestream.microsoft.v1.LocalizedName;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.Term;
import ai.pipestream.microsoft.v1.TermGroup;
import ai.pipestream.microsoft.v1.TermGroupScope;
import ai.pipestream.microsoft.v1.TermLabel;
import ai.pipestream.microsoft.v1.TermSet;
import ai.pipestream.microsoft.v1.TermStore;
import com.google.protobuf.Timestamp;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

class TermStoreRpcTest {

    private static final String STORE = "/sites/s1/termStore";

    private FakeGraphServer fake;
    private Server server;
    private ManagedChannel channel;
    private final ExecutorService executor = Executors.newCachedThreadPool();
    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub;

    @BeforeEach
    void start() throws Exception {
        fake = FakeGraphServer.start();
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("tenant-1").clientId("client-1").clientSecret("env-secret")
                .graphBaseUrl(fake.baseUrl()).authority(fake.baseUrl()).build();
        MicrosoftGrpcService service = new MicrosoftGrpcService(config,
                new GraphFiles(new GraphClient(fake.baseUrl(), () -> "graph-token")), 0);
        String name = InProcessServerBuilder.generateName();
        server = InProcessServerBuilder.forName(name).executor(executor).addService(service).build().start();
        channel = InProcessChannelBuilder.forName(name).build();
        stub = MicrosoftServiceGrpc.newBlockingStub(channel);
    }

    @AfterEach
    void stop() {
        channel.shutdownNow();
        server.shutdownNow();
        executor.shutdownNow();
        fake.close();
    }

    private static <T> List<T> drain(Iterator<T> it) {
        List<T> out = new ArrayList<>();
        it.forEachRemaining(out::add);
        return out;
    }

    private static Timestamp at(String instant) {
        return Timestamp.newBuilder().setSeconds(Instant.parse(instant).getEpochSecond()).build();
    }

    @Test
    void theStoreIsReadFromTheSite() {
        fake.stub(STORE, """
                {"id":"store-1","defaultLanguageTag":"en-US","languageTags":["en-US","de-DE"]}""");

        TermStore store = stub.getTermStore(GetTermStoreRequest.newBuilder().setSiteId("s1").build())
                .getStore();

        assertThat(store).isEqualTo(TermStore.newBuilder().setId("store-1")
                .setDefaultLanguageTag("en-US").addLanguageTags("en-US").addLanguageTags("de-DE").build());
        assertThat(fake.only(STORE).authorization()).isEqualTo("Bearer graph-token");
    }

    @Test
    void groupsFollowTheNextLink() {
        fake.stub(STORE + "/groups", """
                {"value":[{"id":"g1","displayName":"Company","description":"ours","scope":"global",
                  "createdDateTime":"2026-11-01T10:00:00Z"}],
                 "@odata.nextLink":"%s%s/groups?$skiptoken=2"}""".formatted(fake.baseUrl(), STORE));
        fake.stub(STORE + "/groups?$skiptoken=2", """
                {"value":[{"id":"g2","displayName":"System","scope":"system"}]}""");

        List<TermGroup> groups = drain(stub.listTermGroups(
                ListTermGroupsRequest.newBuilder().setSiteId("s1").build()))
                .stream().map(r -> r.getGroup()).toList();

        assertThat(groups).containsExactly(
                TermGroup.newBuilder().setId("g1").setDisplayName("Company").setDescription("ours")
                        .setScope(TermGroupScope.TERM_GROUP_SCOPE_GLOBAL)
                        .setCreatedAt(at("2026-11-01T10:00:00Z")).build(),
                TermGroup.newBuilder().setId("g2").setDisplayName("System")
                        .setScope(TermGroupScope.TERM_GROUP_SCOPE_SYSTEM).build());
    }

    @Test
    void setsOfAGroupCarryTheirNames() {
        fake.stub(STORE + "/groups/g1/sets", """
                {"value":[{"id":"set-1","description":"where people work",
                  "localizedNames":[{"languageTag":"en-US","name":"Departments"}],
                  "createdDateTime":"2026-11-01T10:00:00Z"}]}""");

        List<TermSet> sets = drain(stub.listTermSets(
                ListTermSetsRequest.newBuilder().setSiteId("s1").setGroupId("g1").build()))
                .stream().map(r -> r.getSet()).toList();

        assertThat(sets).containsExactly(TermSet.newBuilder().setId("set-1").setGroupId("g1")
                .setDescription("where people work")
                .addLocalizedNames(LocalizedName.newBuilder().setLanguageTag("en-US").setName("Departments"))
                .setCreatedAt(at("2026-11-01T10:00:00Z")).build());
    }

    @Test
    void termsAreListedUnderTheSetAndUnderATerm() {
        fake.stub(STORE + "/sets/set-1/children", """
                {"value":[{"id":"t1","createdDateTime":"2026-11-01T10:00:00Z",
                  "lastModifiedDateTime":"2026-11-02T10:00:00Z",
                  "labels":[{"name":"Engineering","isDefault":true,"languageTag":"en-US"},
                            {"name":"Eng","isDefault":false,"languageTag":"en-US"}],
                  "descriptions":[{"description":"builders","languageTag":"en-US"}]}]}""");
        fake.stub(STORE + "/sets/set-1/terms/t1/children", """
                {"value":[{"id":"t2","labels":[{"name":"Search","isDefault":true,"languageTag":"en-US"}]}]}""");

        List<Term> roots = drain(stub.listTerms(
                ListTermsRequest.newBuilder().setSiteId("s1").setSetId("set-1").build()))
                .stream().map(r -> r.getTerm()).toList();
        List<Term> children = drain(stub.listTerms(ListTermsRequest.newBuilder().setSiteId("s1")
                .setSetId("set-1").setParentTermId("t1").build()))
                .stream().map(r -> r.getTerm()).toList();

        assertThat(roots).containsExactly(Term.newBuilder().setId("t1").setSetId("set-1")
                .addLabels(TermLabel.newBuilder().setLanguageTag("en-US").setName("Engineering").setIsDefault(true))
                .addLabels(TermLabel.newBuilder().setLanguageTag("en-US").setName("Eng"))
                .addDescriptions(LocalizedName.newBuilder().setLanguageTag("en-US").setName("builders"))
                .setCreatedAt(at("2026-11-01T10:00:00Z"))
                .setLastModifiedAt(at("2026-11-02T10:00:00Z")).build());
        assertThat(children).containsExactly(Term.newBuilder().setId("t2").setSetId("set-1")
                .setParentTermId("t1")
                .addLabels(TermLabel.newBuilder().setLanguageTag("en-US").setName("Search").setIsDefault(true))
                .build());
    }

    @Test
    void aMissingIdIsRefusedBeforeAnyCall() {
        assertThatThrownBy(() -> drain(stub.listTermSets(
                ListTermSetsRequest.newBuilder().setSiteId("s1").build())))
                .isInstanceOfSatisfying(StatusRuntimeException.class, e -> assertThat(e.getStatus().getCode())
                        .isEqualTo(Status.Code.INVALID_ARGUMENT));
        assertThatThrownBy(() -> drain(stub.listTerms(
                ListTermsRequest.newBuilder().setSiteId("s1").build())))
                .isInstanceOfSatisfying(StatusRuntimeException.class, e -> assertThat(e.getStatus().getCode())
                        .isEqualTo(Status.Code.INVALID_ARGUMENT));
    }
}
