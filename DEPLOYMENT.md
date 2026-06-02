# Deployment Guide — QB Combiner

This app self-hosts in two modes:

1. **Local on your Windows/Mac/Linux computer** — runs at `http://localhost:8501`
2. **On a Linux VM** (any provider: AWS EC2, DigitalOcean, Linode, Hetzner) — runs at `http://your-server-ip:8501` or behind a domain

Both modes use Docker. **You only need to install Docker once** — everything else is one command.

---

## Option A — Local (your own computer)

### 1. Install Docker Desktop

- **Windows**: https://docs.docker.com/desktop/install/windows-install/
- **Mac**: https://docs.docker.com/desktop/install/mac-install/
- **Linux**: https://docs.docker.com/desktop/install/linux-install/

After install, open Docker Desktop and wait for it to say "Docker is running".

### 2. Set your team password

In this folder (`qb_combiner_app`), copy the example secrets file:

**Windows (PowerShell):**
```powershell
Copy-Item .streamlit\secrets.toml.example .streamlit\secrets.toml
notepad .streamlit\secrets.toml
```

**Mac/Linux:**
```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
nano .streamlit/secrets.toml
```

Change `change-this-to-a-strong-password` to whatever password you want your team to use. Save and close.

### 3. Build and run

Open a terminal in this folder and run:

```bash
docker compose up -d --build
```

The `-d` flag runs it in the background. First build takes ~2 minutes; subsequent starts take ~5 seconds.

### 4. Open the app

Visit **http://localhost:8501** in your browser. Sign in with the password from step 2.

### 5. Stop / restart

```bash
docker compose down       # stop
docker compose up -d      # start again
docker compose logs -f    # follow logs
```

---

## Option B — Linux VM (cloud)

### 1. Provision a VM

Any Ubuntu 22.04 / Debian 12 box with at least 2 GB RAM works. Cheap options:

- DigitalOcean Droplet: $6/mo (1 vCPU, 1 GB RAM) — fine for ≤5 users
- Hetzner CX22: ~€4/mo (2 vCPU, 4 GB RAM) — preferred
- AWS Lightsail: $5/mo

### 2. Install Docker on the VM

```bash
ssh user@your-server-ip

# Install Docker
curl -fsSL https://get.docker.com | sudo sh
sudo usermod -aG docker $USER
newgrp docker  # re-login or run this to refresh group
```

### 3. Copy the app to the server

From your local machine:

```bash
# Replace user@your-server-ip with your actual SSH target
scp -r ./qb_combiner_app user@your-server-ip:~/qb_combiner_app
```

Or zip the folder and upload it via your provider's file manager.

### 4. Set the password on the server

```bash
ssh user@your-server-ip
cd ~/qb_combiner_app
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
nano .streamlit/secrets.toml   # change the password
```

### 5. Run

```bash
docker compose up -d --build
```

### 6. Open firewall port 8501

DigitalOcean / Hetzner / AWS Lightsail all have a firewall panel — add an inbound rule for **TCP port 8501** from `0.0.0.0/0`.

### 7. Open the app

Visit **http://your-server-ip:8501** and sign in.

---

## Optional: Custom domain + HTTPS

If you want `qb.yourcompany.com` with a real HTTPS cert:

### 1. Point DNS

In your DNS provider (Cloudflare, Namecheap, etc.), add an **A record** pointing `qb.yourcompany.com` to your server IP.

### 2. Install Caddy (zero-config HTTPS)

On the server:

```bash
sudo apt install -y caddy
sudo nano /etc/caddy/Caddyfile
```

Paste:

```
qb.yourcompany.com {
    reverse_proxy localhost:8501
}
```

Then:

```bash
sudo systemctl reload caddy
```

Caddy will automatically get a Let's Encrypt cert. Visit **https://qb.yourcompany.com** — that's it.

You can now close port 8501 on the firewall and leave only 80/443 open.

---

## Updating the app

When you change code:

```bash
docker compose down
docker compose up -d --build
```

---

## Troubleshooting

**"Port 8501 already in use"** — something else is bound to it. Either stop the other thing, or change the port in `docker-compose.yml`:

```yaml
ports:
  - "9000:8501"   # access at http://localhost:9000
```

**"Container exits immediately"** — view logs:

```bash
docker compose logs
```

The most common cause is a missing or malformed `secrets.toml`. Compare with `secrets.toml.example`.

**"Upload file too large"** — bump the limit in `.streamlit/config.toml`:

```toml
[server]
maxUploadSize = 500   # MB
```

Then rebuild.

**Multiple users editing at the same time** — Streamlit gives each browser session its own `session_state`. Two users will not see each other's files. Each user uploads, reviews, and downloads independently. Nothing is persisted server-side beyond the session.

---

## MongoDB — profile and mapping persistence

The app uses MongoDB to persist **mapping profiles** (saved mapping decisions per client/year). Without Mongo the app still works, but everything resets when the tab closes.

### When you `docker compose up`

The mongo container starts automatically alongside the app — nothing to configure. Mappings get stored in the named volume `mongo-data` and survive container restarts/rebuilds.

### Where the data lives

```bash
docker volume inspect qb_combiner_app_mongo-data
# Shows the host path of the persistent mongo storage
```

### Backups

Snapshot the mongo volume periodically. On the host:

```bash
docker compose exec mongo mongodump --db qb_combiner --out /tmp/backup
docker compose cp mongo:/tmp/backup ./backups/$(date +%Y%m%d)
```

Restore:

```bash
docker compose cp ./backups/20260518/qb_combiner mongo:/tmp/restore
docker compose exec mongo mongorestore --db qb_combiner /tmp/restore
```

You can also export individual profiles as JSON from the **🗂️ Profiles** page (clean, human-readable, easy to commit to git for change-tracking).

### Using an external/cloud MongoDB

Edit `.streamlit/secrets.toml`:

```toml
MONGODB_URI = "mongodb+srv://user:pass@cluster.mongodb.net"
MONGODB_DB  = "qb_combiner"
```

Then comment out the `mongo` service in `docker-compose.yml` and run `docker compose up qb-combiner` only.

---

## Backing up the rest

Beyond MongoDB, the only file worth keeping safe is `.streamlit/secrets.toml` (your password). Uploaded QB files and generated workbooks are in-memory only — closing the tab discards them.
