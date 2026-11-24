package ai.pipestream.microsoft.admin;

import ai.pipestream.microsoft.GraphClient;
import ai.pipestream.microsoft.GraphErrors;
import ai.pipestream.microsoft.GraphPager;
import ai.pipestream.microsoft.GraphPaths;
import ai.pipestream.microsoft.v1.EnsureSiteAppGrantRequest;
import ai.pipestream.microsoft.v1.EnsureSiteAppGrantResponse;
import ai.pipestream.microsoft.v1.ListSiteAppGrantsRequest;
import ai.pipestream.microsoft.v1.ListSiteAppGrantsResponse;
import ai.pipestream.microsoft.v1.MicrosoftAdminServiceGrpc;
import ai.pipestream.microsoft.v1.RevokeSiteAppGrantRequest;
import ai.pipestream.microsoft.v1.RevokeSiteAppGrantResponse;
import ai.pipestream.microsoft.v1.SiteAppGrant;
import ai.pipestream.microsoft.v1.SiteAppGrantState;
import ai.pipestream.microsoft.v1.SiteAppRole;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.grpc.stub.StreamObserver;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/** Lists, ensures and revokes the grants applications hold on SharePoint sites. */
public final class MicrosoftAdminGrpcService extends MicrosoftAdminServiceGrpc.MicrosoftAdminServiceImplBase {

    private final GraphClient graph;
    private final GraphPager pager;

    /**
     * Creates the service.
     *
     * @param graph a Graph client signed in as the admin identity
     */
    public MicrosoftAdminGrpcService(GraphClient graph) {
        this.graph = graph;
        this.pager = new GraphPager(graph);
    }

    @Override
    public void listSiteAppGrants(ListSiteAppGrantsRequest request,
            StreamObserver<ListSiteAppGrantsResponse> observer) {
        try {
            for (SiteAppGrant grant : grants(request.getSiteId())) {
                observer.onNext(ListSiteAppGrantsResponse.newBuilder().setGrant(grant).build());
            }
        } catch (Throwable t) {
            observer.onError(GraphErrors.toException(t));
            return;
        }
        observer.onCompleted();
    }

    @Override
    public void ensureSiteAppGrant(EnsureSiteAppGrantRequest request,
            StreamObserver<EnsureSiteAppGrantResponse> observer) {
        EnsureSiteAppGrantResponse response;
        try {
            response = ensure(request);
        } catch (Throwable t) {
            observer.onError(GraphErrors.toException(t));
            return;
        }
        observer.onNext(response);
        observer.onCompleted();
    }

    @Override
    public void revokeSiteAppGrant(RevokeSiteAppGrantRequest request,
            StreamObserver<RevokeSiteAppGrantResponse> observer) {
        int revoked = 0;
        try {
            String appId = requireAppId(request.getAppId());
            String permissions = permissionsPath(request.getSiteId());
            List<String> ids = new ArrayList<>();
            for (SiteAppGrant grant : grants(request.getSiteId())) {
                if (grant.getAppId().equalsIgnoreCase(appId) && !ids.contains(grant.getId())) {
                    ids.add(grant.getId());
                }
            }
            for (String id : ids) {
                graph.delete(permissions + "/" + GraphPaths.segment(id, "permission id"));
                revoked++;
            }
        } catch (Throwable t) {
            observer.onError(GraphErrors.toException(t));
            return;
        }
        observer.onNext(RevokeSiteAppGrantResponse.newBuilder().setRevoked(revoked).build());
        observer.onCompleted();
    }

    private EnsureSiteAppGrantResponse ensure(EnsureSiteAppGrantRequest request)
            throws IOException, InterruptedException {
        String appId = requireAppId(request.getAppId());
        String role = graphRole(request.getRole());
        String permissions = permissionsPath(request.getSiteId());
        List<SiteAppGrant> held = grants(request.getSiteId()).stream()
                .filter(grant -> grant.getAppId().equalsIgnoreCase(appId)).toList();
        if (held.size() > 1) {
            throw new IllegalStateException("the application holds " + held.size()
                    + " grants on the site; revoke them and ensure again");
        }
        EnsureSiteAppGrantResponse.Builder response = EnsureSiteAppGrantResponse.newBuilder();
        if (held.isEmpty()) {
            if (request.getDryRun()) {
                return response.setState(SiteAppGrantState.SITE_APP_GRANT_STATE_WOULD_CREATE).build();
            }
            ObjectNode body = JsonNodeFactory.instance.objectNode();
            body.putArray("roles").add(role);
            body.putArray("grantedToIdentities").addObject().putObject("application")
                    .put("id", appId).put("displayName", request.getAppDisplayName().isBlank()
                            ? appId : request.getAppDisplayName().trim());
            JsonNode created = graph.post(permissions, body);
            return response.setState(SiteAppGrantState.SITE_APP_GRANT_STATE_CREATED)
                    .setGrant(grantOf(created, request.getSiteId(), appId)).build();
        }
        SiteAppGrant grant = held.get(0);
        if (grant.getRolesList().equals(List.of(request.getRole())) && grant.getOtherRolesCount() == 0) {
            return response.setState(SiteAppGrantState.SITE_APP_GRANT_STATE_EXISTS).setGrant(grant).build();
        }
        response.addAllPreviousRoles(grant.getRolesList());
        if (request.getDryRun()) {
            return response.setState(SiteAppGrantState.SITE_APP_GRANT_STATE_WOULD_UPDATE).setGrant(grant)
                    .build();
        }
        ObjectNode body = JsonNodeFactory.instance.objectNode();
        body.putArray("roles").add(role);
        JsonNode updated = graph.patch(permissions + "/" + GraphPaths.segment(grant.getId(),
                "permission id"), body);
        return response.setState(SiteAppGrantState.SITE_APP_GRANT_STATE_UPDATED)
                .setGrant(grantOf(updated, request.getSiteId(), appId)).build();
    }

    /** Every application grant of a site; the roles come from reading each permission. */
    private List<SiteAppGrant> grants(String siteId) throws IOException, InterruptedException {
        String permissions = permissionsPath(siteId);
        List<String> ids = new ArrayList<>();
        pager.forEach(List.of(permissions), Map.of(), node -> {
            ids.add(node.path("id").asText(""));
            return true;
        });
        List<SiteAppGrant> grants = new ArrayList<>();
        for (String id : ids) {
            JsonNode permission = graph.get(permissions + "/" + GraphPaths.segment(id, "permission id"));
            for (JsonNode identity : identities(permission)) {
                JsonNode application = identity.path("application");
                if (application.isObject()) {
                    grants.add(grant(permission, siteId, application));
                }
            }
        }
        return grants;
    }

    private static JsonNode identities(JsonNode permission) {
        JsonNode v2 = permission.path("grantedToIdentitiesV2");
        return v2.isArray() && !v2.isEmpty() ? v2 : permission.path("grantedToIdentities");
    }

    /** The grant of {@code appId} in a permission Graph just answered. */
    private static SiteAppGrant grantOf(JsonNode permission, String siteId, String appId) {
        for (JsonNode identity : identities(permission)) {
            if (identity.path("application").path("id").asText("").equalsIgnoreCase(appId)) {
                return grant(permission, siteId, identity.path("application"));
            }
        }
        return grant(permission, siteId, JsonNodeFactory.instance.objectNode().put("id", appId));
    }

    private static SiteAppGrant grant(JsonNode permission, String siteId, JsonNode application) {
        SiteAppGrant.Builder grant = SiteAppGrant.newBuilder()
                .setId(permission.path("id").asText(""))
                .setSiteId(siteId.trim())
                .setAppId(application.path("id").asText(""))
                .setAppDisplayName(application.path("displayName").asText(""));
        for (JsonNode role : permission.path("roles")) {
            SiteAppRole known = role(role.asText(""));
            if (known == SiteAppRole.SITE_APP_ROLE_UNSPECIFIED) {
                grant.addOtherRoles(role.asText(""));
            } else {
                grant.addRoles(known);
            }
        }
        return grant.build();
    }

    static SiteAppRole role(String graphRole) {
        return switch (graphRole) {
            case "read" -> SiteAppRole.SITE_APP_ROLE_READ;
            case "write" -> SiteAppRole.SITE_APP_ROLE_WRITE;
            case "manage" -> SiteAppRole.SITE_APP_ROLE_MANAGE;
            case "fullcontrol" -> SiteAppRole.SITE_APP_ROLE_FULL_CONTROL;
            case "owner" -> SiteAppRole.SITE_APP_ROLE_OWNER;
            default -> SiteAppRole.SITE_APP_ROLE_UNSPECIFIED;
        };
    }

    static String graphRole(SiteAppRole role) {
        return switch (role) {
            case SITE_APP_ROLE_READ -> "read";
            case SITE_APP_ROLE_WRITE -> "write";
            case SITE_APP_ROLE_MANAGE -> "manage";
            case SITE_APP_ROLE_FULL_CONTROL -> "fullcontrol";
            case SITE_APP_ROLE_OWNER -> "owner";
            case SITE_APP_ROLE_UNSPECIFIED, UNRECOGNIZED ->
                    throw new IllegalArgumentException("role is required");
        };
    }

    private static String permissionsPath(String siteId) {
        if (siteId.isBlank()) {
            throw new IllegalArgumentException("site_id is required");
        }
        return "/sites/" + GraphPaths.site(siteId) + "/permissions";
    }

    private static String requireAppId(String appId) {
        if (appId.isBlank()) {
            throw new IllegalArgumentException("app_id is required");
        }
        return appId.trim();
    }
}
