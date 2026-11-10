#!/usr/bin/env bash
# Boot a built release image without host ports or real upstream credentials.
# Requires Docker and GNU timeout. Helper images may be overridden with digests.
# Usage: scripts/container-smoke.sh MODE IMAGE
set -euo pipefail
mode=${1:-}
image=${2:-}
case "$mode" in
  microsoft|connector) ;;
  *) echo 'Usage: scripts/container-smoke.sh {microsoft|connector} IMAGE' >&2; exit 2 ;;
esac
[[ -n "$image" && $# == 2 ]] || { echo 'An image reference is required' >&2; exit 2; }
command -v timeout >/dev/null
probe_image=${GRPCURL_IMAGE:-fullstorydev/grpcurl:v1.9.3}
python_image=${PYTHON_IMAGE:-python:3.13-alpine}
prefix="split-smoke-${mode}-$(date +%s)-$$-${RANDOM}"
network="$prefix"
server="$prefix-server"
fixture="$prefix-fixture"
probe="$prefix-probe"
created_network=false
containers=()
# Bound every Docker operation as well as each RPC; no published host ports.
d() { timeout --kill-after=5s 30s docker "$@"; }
cleanup() {
  result=$?
  trap - EXIT
  if (( result != 0 )); then
    d logs --tail 80 "$server" >&2 || true
    d logs --tail 40 "$fixture" >&2 || true
  fi
  for container in "${containers[@]}"; do
    if d container inspect "$container" >/dev/null 2>&1; then
      if ! d rm -f "$container" >/dev/null; then
        echo "Cleanup failed for owned container $container" >&2
        result=1
      fi
    fi
  done
  if $created_network && ! d network rm "$network" >/dev/null; then
    echo "Cleanup failed for owned network $network" >&2
    result=1
  fi
  printf 'container-smoke mode=%s exit=%s\n' "$mode" "$result"
  exit "$result"
}
trap cleanup EXIT
trap 'exit 130' INT
trap 'exit 143' TERM
ensure_image() {
  if ! d image inspect "$1" >/dev/null 2>&1; then timeout --kill-after=5s 180s docker pull "$1"; fi
  d image inspect --format 'helper={{.Id}} repoDigests={{json .RepoDigests}}' "$1"
}
d image inspect --format 'image={{.Id}} repoDigests={{json .RepoDigests}}' "$image"
if [[ "$mode" == microsoft ]]; then ensure_image "$python_image"; fi
ensure_image "$probe_image"
created_network=true
d network create --internal "$network" >/dev/null
runtime=(--network "$network" --read-only --tmpfs /tmp:rw,exec,nosuid,size=128m
         --cap-drop ALL --security-opt no-new-privileges --cpus 2 --memory 768m
         -e JAVA_TOOL_OPTIONS=-XX:ActiveProcessorCount=2)
env_args=()
case "$mode" in
  connector) port=30303 ;;
  microsoft)
    port=9096
    # Startup performs client-credentials auth. Fake only that upstream,
    # while booting the unmodified production entrypoint and distribution.
    containers+=("$fixture")
    d run -d --name "$fixture" --network "$network" --network-alias auth \
      --read-only --cap-drop ALL --security-opt no-new-privileges \
      --cpus 1 --memory 128m "$python_image" python -u -c '
from http.server import BaseHTTPRequestHandler, HTTPServer
class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200); self.end_headers(); self.wfile.write(b"ready")
    def do_POST(self):
        if self.path != "/smoke/oauth2/v2.0/token":
            self.send_error(404); return
        self.rfile.read(int(self.headers.get("Content-Length", "0")))
        body = b"{\"access_token\":\"smoke-only\",\"expires_in\":3600,\"token_type\":\"Bearer\"}"
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers(); self.wfile.write(body)
HTTPServer(("0.0.0.0", 8080), Handler).serve_forever()
' >/dev/null
    # docker run is asynchronous; wait for the token fixture before server boot.
    ready=false
    for attempt in {1..30}; do
      if d exec "$fixture" python -c 'import urllib.request; urllib.request.urlopen("http://localhost:8080", timeout=2).read()' >/dev/null 2>&1; then ready=true; break; fi
      sleep 1
    done
    $ready || { echo 'OAuth fixture failed to start' >&2; exit 1; }
    env_args=(-e MICROSOFT_TENANT_ID=smoke -e MICROSOFT_CLIENT_ID=smoke
              -e MICROSOFT_CLIENT_SECRET=smoke-only -e MICROSOFT_AUTHORITY=http://auth:8080
              -e MICROSOFT_GRAPH_BASE_URL=http://auth:8080/v1.0)
    ;;
esac
containers+=("$server" "$probe")
d run -d --name "$server" --network-alias service "${runtime[@]}" "${env_args[@]}" "$image" >/dev/null
ready=false
for attempt in {1..30}; do
  [[ $(d inspect --format '{{.State.Running}}' "$server") == true ]] || {
    echo 'Service exited before readiness' >&2; exit 1;
  }
  if response=$(d run --name "$probe" --network "$network" --read-only --cap-drop ALL \
      --security-opt no-new-privileges --cpus 1 --memory 128m "$probe_image" \
      -plaintext -connect-timeout 2 -max-time 4 -d '{}' "service:$port" grpc.health.v1.Health/Check 2>&1) \
      && [[ "$response" == *'"SERVING"'* ]]; then
    printf '%s\n' "$response"
    ready=true; break
  fi
  printf 'readiness attempt %s: %s\n' "$attempt" "$response" >&2
  d rm -f "$probe" >/dev/null 2>&1 || true
  sleep 1
done
$ready || { echo 'Service did not become ready within 30 attempts' >&2; exit 1; }
d inspect --format 'runtime image={{.Image}} running={{.State.Running}} user={{.Config.User}} readonly={{.HostConfig.ReadonlyRootfs}} ports={{json .HostConfig.PortBindings}}' "$server"
d logs --tail 20 "$server"
