# Recorded Graph exchanges

`microsoft/` holds real Microsoft Graph requests and responses, recorded while
the write scenarios (`WriteScenarios`) ran against a real tenant and scrubbed
before they were written. The same scenarios run against them in the default
`test` task, with no network.

Do not write or edit these files by hand. Record them.

## Layout

```
microsoft/run.json                       suffix, Graph base URL, placeholder site and drive id
microsoft/<scenario>/NNN-METHOD-route.json   one HTTP exchange, in the order it happened
```

Each exchange file has the request (method, URL, whether it carried
`Authorization`, `If-Match` / `Content-Range` / `Content-Type` / `Prefer` /
`Range`, and the JSON body or the sha256 and length of a binary one) and the
response (status, `ETag` / `Location` / `Retry-After` / `Content-Range` /
`Content-Type`, and the body). File content the scenarios generated themselves
is stored as `blob` + sha256 and regenerated from the run suffix on replay.

## Record

Needs an app with write permission on a test site. It creates one folder
`pipestream-live-write-<suffix>` at the drive root and one list of the same
name, and deletes both at the end.

```
export MICROSOFT_TENANT_ID=... MICROSOFT_CLIENT_ID=... MICROSOFT_CLIENT_SECRET=...
export MICROSOFT_SITE_ID=...          # MICROSOFT_DRIVE_IDS optional
MICROSOFT_LIVE_WRITE=1 ./gradlew :grpc-microsoft-service:liveWriteTest
```

A run replaces everything under `microsoft/`. Only scenarios that passed are
saved. The manual CI job `live-microsoft-write` uploads the directory as the
`recordings` artifact. Read the diff before committing it.

## Scrubbing

`RecordingScrubber` replaces, with the same placeholder for the same real
value: GUIDs (tenant, client, list, user ids, eTags), drive ids, drive item
ids, tenant host names, email addresses and UPNs, display names of users,
groups and apps, access tokens, and every URL that is not on the Graph host
(web URLs, download, upload-session and monitor URLs). `Authorization` and
cookies are never recorded; token-endpoint calls do not pass through the
recorder. A file that still looks like it holds tenant data is not written,
and `RecordingsHygieneTest` checks every committed file again.

## Replay

`MicrosoftWriteReplayTest` runs each scenario against its directory. Every
request must equal the next recorded one (method, URL with query,
`Authorization` present or not, the headers above, body); anything else fails
the scenario, and so do recorded exchanges that were never requested. A
scenario without a directory is reported as skipped.

## Hand-written fake tests

`MicrosoftWriteRpcTest` and the other `FakeGraphServer` tests use made-up
responses. They stay until recordings of the same rpcs are committed here, and
are then to be replaced by replay.
