# Passage Web — self-hosted deployment guide

`Passage.Web` is a Blazor Server port of the Passage editor for self-hosting on
a NAS or home server. It reuses the same `Passage.Core`, `Passage.Parser`, and
`Passage.Export` libraries as the desktop apps, so parsing, page layout, and
PDF/text export behave identically.

## What it does

- **Script editor** with live Fountain syntax highlighting (scene headings,
  characters, dialogue, transitions, sections, synopses, notes, boneyard,
  lyrics, title page), powered by the real `FountainParser` on the server.
- **Outline sidebar** — Acts / Sequences / Scenes / Synopses, click to jump.
- **Notes sidebar** — every `[[note]]` with jump-to-line.
- **Beat Board** — read-only index cards mirroring the script structure.
- **Page preview** — paginated US-Letter pages rendered from the same layout
  engine the PDF exporter uses, with zoom.
- **Exports** — PDF, plain text, and raw `.fountain` downloads.
- **File library** stored server-side on the `/data` volume: create, open,
  save (Ctrl+S), delete, optional autosave.
- Markdown mode for `.md` files (plain editing plus a heading outline).

Scripts are plain `.fountain`/`.md`/`.txt` files in the volume, so they are
easy to back up and remain fully compatible with the desktop apps.

## Deploy with Portainer

### Option A — Stack from the git repository

1. Portainer → **Stacks** → **Add stack** → **Repository**.
2. Repository URL: this repo; Compose path: `docker-compose.yml`.
3. Deploy. Portainer builds the image on the host and starts the container.

### Option B — Build the image yourself

```bash
git clone <this repo> && cd <repo>
docker build -t passage-web .
docker run -d --name passage \
  -p 8095:8080 \
  -v passage-data:/data \
  --restart unless-stopped \
  passage-web
```

Then open `http://<nas-ip>:8095`.

The container listens on 8080 internally; the compose file publishes it on
host port **8095** because 8080 is commonly taken on a NAS (Nginx Proxy
Manager's default HTTP port, for one). Change the left-hand side of the
`ports:` mapping if 8095 clashes with something on your host.

### Option C — Behind Nginx Proxy Manager (no published port)

If you run NPM, skip the host port entirely and let NPM reach the container
over a shared Docker network:

1. Remove the `ports:` section from `docker-compose.yml` and attach the
   service to the network NPM uses, e.g.:

   ```yaml
   services:
     passage:
       # ...
       networks: [proxy]
   networks:
     proxy:
       external: true   # the network your NPM container is on
   ```

2. In NPM, add a Proxy Host: domain of your choice → forward to
   `passage` : `8080` (scheme `http`).
3. Enable **WebSockets Support** on that proxy host — Blazor Server requires
   it. Add NPM's access list or SSL as desired.

## Configuration

| Setting | Default | Notes |
| --- | --- | --- |
| `PASSAGE_DATA_DIR` | `/data` | Where scripts are stored. Mount a volume or a bind path (e.g. `/volume1/docker/passage:/data`). |
| `ASPNETCORE_URLS` | `http://+:8080` | Change the listen port if needed. |

To store scripts in an existing NAS share instead of a named volume, replace
the volume line in `docker-compose.yml` with a bind mount, e.g.
`- /volume1/docker/passage:/data`.

## Notes and limitations

- **No authentication.** Intended for a trusted LAN. If you expose it beyond
  your network, put it behind your reverse proxy's auth (Authelia, Tailscale,
  basic auth, etc.).
- Blazor Server needs a persistent WebSocket; reverse proxies must allow
  WebSocket upgrades for the app's path.
- Saving is last-write-wins; two browsers editing the same file will overwrite
  each other.
- The Beat Board is a read-only mirror of the script structure in this first
  version (the desktop apps remain the place for card editing), and writing
  goals/timers haven't been ported yet.
