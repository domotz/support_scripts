# Grafana and Nginx Setup with Docker (HTTP Only - No SSL)

**Note**: This guide is designed for Ubuntu systems using the `ubuntu` user. Throughout this guide, you'll need to replace `my-ip` with your VM's IP address or FQDN (e.g., 192.168.1.100 or grafana.example.com). An automated replacement command using `sed` is provided in the configuration section below.

## Manage Grafana Docker Instance

### Folder

```bash
cd /home/ubuntu/grafana
```

### Configuration Files

- Docker Compose: `~/grafana/docker-compose-NO-SSL.yml`
- Nginx Configuration: `~/grafana/conf/default-NO-SSL.conf`

### Docker Management Commands

#### Creating with detach mode

```bash
docker compose -f ~/grafana/docker-compose-NO-SSL.yml up -d
```

#### Stopping the containers

```bash
docker compose -f ~/grafana/docker-compose-NO-SSL.yml down
```

### Logging

```bash
docker logs grafana -f
docker logs grafana-nginx -f
```

### Restart

```bash
docker restart grafana
docker restart grafana-nginx
```

---

## Docker Installation on Ubuntu 22.04 LTS

Reference: https://docs.docker.com/engine/install/ubuntu/

### Use ubuntu user with sudo

**Note**: Throughout this guide, all commands should be run as the `ubuntu` user unless explicitly stated otherwise.

```bash
sudo -i
```

### Remove old Docker packages (no sudo as sudo -i above)

```bash
for pkg in docker.io docker-doc docker-compose docker-compose-v2 podman-docker containerd runc; do apt-get remove $pkg; done
```

### Add Docker's official GPG key

```bash
apt-get update
apt-get install ca-certificates curl net-tools -y
install -m 0755 -d /etc/apt/keyrings
curl -fsSL https://download.docker.com/linux/ubuntu/gpg -o /etc/apt/keyrings/docker.asc
chmod a+r /etc/apt/keyrings/docker.asc
```

### Add the repository to Apt sources

```bash
echo \
  "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.asc] https://download.docker.com/linux/ubuntu \
  $(. /etc/os-release && echo "${UBUNTU_CODENAME:-$VERSION_CODENAME}") stable" | \
  tee /etc/apt/sources.list.d/docker.list > /dev/null
apt-get update
apt-get install docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin -y
docker run hello-world
systemctl enable docker.service
systemctl enable containerd.service
```

---

## Enable Ubuntu User to Run Docker

Reference: https://docs.docker.com/engine/install/linux-postinstall/

User ubuntu (or other non-root user) run:

```bash
# use ubuntu user
sudo groupadd docker
sudo usermod -aG docker $USER
newgrp docker

# test docker
docker run hello-world
```

---

## Docker Autocompletion

```bash
# use ubuntu user
sudo apt install bash-completion

cat <<EOT >> ~/.bashrc
if [ -f /etc/bash_completion ]; then
    . /etc/bash_completion
fi
EOT
```

---

## Insert Ubuntu User in Sudoers

```bash
ls -l /etc/sudoers
sudo EDITOR=vim visudo
```

Add at the end of the file:

```
%ubuntu   ALL=(ALL)       NOPASSWD: ALL
%domotz   ALL=(ALL)       NOPASSWD: ALL
```

---

## Check Ubuntu Firewall

```bash
sudo ufw status
```

If active, open port 80:

```bash
sudo ufw allow 80/tcp
```

---

## Deploy Grafana & Nginx Configuration

### Folder Structure Creation

```bash
# use ubuntu user
mkdir -p ~/grafana/{conf,data,plugins/datasource-domotz}


cd ~/grafana
```

### Copy Files to Ubuntu Machine

**Use the `ubuntu` user for all operations.**

Copy the following files from your local machine to the Ubuntu VM using `scp` or your preferred file transfer method:

#### From your local machine:

```bash
# Copy Docker Compose configuration
scp docker-compose-NO-SSL.yml ubuntu@<VM-IP>:~/grafana/

# Copy Nginx configuration
scp default-NO-SSL.conf ubuntu@<VM-IP>:~/grafana/conf/

# Copy Domotz Grafana plugin archive
scp dist-grafana-plugin_2025-06-05.tar.gz ubuntu@<VM-IP>:~/grafana/plugins/
```

Replace `<VM-IP>` with your actual VM IP address (e.g., 192.168.1.100).

#### On the Ubuntu VM (as ubuntu user):

Extract the Domotz plugin to the correct location:

```bash
# Navigate to the plugins directory
cd ~/grafana/plugins/datasource-domotz

# Extract the plugin archive
tar xzf ../dist-grafana-plugin_2025-06-05.tar.gz -C .

# Verify extraction
ls -la ~/grafana/plugins/datasource-domotz
```

#### Expected final directory structure:

```bash
# use ubuntu user
ubuntu@MyTestVM2-UK:~/grafana$ ll
total 24
drwxrwxr-x 5 ubuntu docker 4096 Jun  5 09:31 ./
drwxr-x--- 6 ubuntu ubuntu 4096 Jun  5 09:38 ../
drwxrwxr-x 2 ubuntu docker 4096 Jun  5 09:49 conf/
drwxrwxr-x 6 ubuntu docker 4096 Jun  5 11:55 data/
-rw-r--r-- 1 ubuntu ubuntu 1290 Jun  5 09:50 docker-compose-NO-SSL.yml
drwxrwxr-x 3 ubuntu docker 4096 Jun  5 09:28 plugins/

ubuntu@MyTestVM2-UK:~/grafana/conf$ ls -l
total 4
-rw-r--r-- 1 ubuntu ubuntu 1927 Jun  5 09:30 default-NO-SSL.conf

ubuntu@MyTestVM2-UK:~/grafana/plugins$ ll
total 12
drwxrwxr-x 3 ubuntu docker 4096 Jun  5 09:28 ./
drwxrwxr-x 5 ubuntu docker 4096 Jun  5 09:31 ../
drwxrwxr-x 3 ubuntu docker 4096 Jun  5 09:32 datasource-domotz/
```

### Important: Update Configuration Files

**Use the `ubuntu` user for these operations.**

Before starting the containers, update the configuration files to replace `my-ip` with your actual VM IP address or FQDN.

#### Automatic replacement using sed:

```bash
# Navigate to the grafana directory
cd ~/grafana

# Set your IP address or FQDN (replace with your actual value)
export MY_HOST="192.168.1.100"  # Or use FQDN like "grafana.example.com"

# Replace my-ip in docker-compose-NO-SSL.yml
sed -i "s/my-ip/$MY_HOST/g" docker-compose-NO-SSL.yml

# Verify the changes
grep -E "(GF_SERVER_ROOT_URL|GF_SERVER_DOMAIN)" docker-compose-NO-SSL.yml
```

#### Manual verification:

1. In `docker-compose-NO-SSL.yml`:
   
   - Verify that `my-ip` has been replaced with your actual IP/FQDN in these lines:
     - `GF_SERVER_ROOT_URL: http://<YOUR-IP>/grafana`
     - `GF_SERVER_DOMAIN: <YOUR-IP>`

2. The `default-NO-SSL.conf` file uses `server_name _;` which accepts requests on any IP/hostname. No changes needed unless you want to restrict access to a specific IP.

---

## Start the Services

**Use the `ubuntu` user for these operations.**

```bash
# Navigate to the grafana folder
cd ~/grafana

# Start the containers in detached mode
docker compose -f docker-compose-NO-SSL.yml up -d

# Check the logs to verify everything is working
docker logs grafana -f
docker logs grafana-nginx -f
```

---

## Access Grafana

Open your browser and navigate to:

```
http://<YOUR-IP-OR-FQDN>/grafana
```

Or simply access the root URL and you'll be automatically redirected:

```
http://<YOUR-IP-OR-FQDN>
```

Replace `<YOUR-IP-OR-FQDN>` with the actual IP address or FQDN you configured in the previous step (e.g., `http://192.168.1.100` or `http://grafana.example.com`)

**Note**: Accessing the root path (`/`) will automatically redirect to `/grafana/`.

Default Grafana credentials:

- Username: `admin`
- Password: `admin` (you'll be prompted to change this on first login)

---

## Testing

**Use the `ubuntu` user for these operations.**

### Test root redirect to Grafana

```bash
# Replace <YOUR-IP-OR-FQDN> with your actual IP address or FQDN
curl -I http://<YOUR-IP-OR-FQDN>
```

You should see a `301 Moved Permanently` redirect to `/grafana/`. Nginx is configured to automatically redirect from the root path (`/`) to `/grafana/`.

### Test Grafana access

```bash
# Replace <YOUR-IP-OR-FQDN> with your actual IP address or FQDN
curl http://<YOUR-IP-OR-FQDN>/grafana/
```

You should see HTML content from the Grafana login page.

---

## Troubleshooting

### Check if containers are running

```bash
docker ps
```

### Check container logs

```bash
docker logs grafana
docker logs grafana-nginx
```

### Restart containers

```bash
docker restart grafana
docker restart grafana-nginx
```

### Check firewall

```bash
sudo ufw status
```

Ensure port 80 is allowed.

### Check if port 80 is listening

```bash
sudo netstat -tlnp | grep :80
```

---

## Differences from SSL Version

This setup differs from the SSL version in the following ways:

1. **No SSL Certificates**: No Let's Encrypt certificates are required or used
2. **Port 80 Only**: Only HTTP port 80 is exposed (no port 443)
3. **No Certbot**: No need to install or configure certbot
4. **No DNS Required**: Works directly with IP addresses instead of domain names
5. **HTTP Protocol**: All communication is over HTTP (not encrypted)
6. **Simpler Nginx Config**: No SSL directives in the Nginx configuration
7. **Simplified Docker Compose**: No SSL certificate volume mounts

---

## Security Warning

**Important**: This setup uses HTTP without encryption. This means:

- All data transmitted between the browser and Grafana is **not encrypted**
- Passwords and sensitive data can be intercepted
- This setup is suitable for:
  - Internal networks only
  - Testing and development environments
  - Networks protected by VPN or other security measures

**For production environments or public-facing installations, always use the SSL version with HTTPS.**

---

## File Summary

Required files for this setup:

1. `docker-compose-NO-SSL.yml` - Docker Compose configuration for HTTP-only setup
2. `conf/default-NO-SSL.conf` - Nginx configuration for HTTP-only reverse proxy
3. `plugins/datasource-domotz/` - Domotz Grafana datasource plugin
