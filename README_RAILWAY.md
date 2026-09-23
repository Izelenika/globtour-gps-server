# GLOBTOUR GPS SERVER — Railway

## 1. GitHub
Create a new GitHub repository, for example:

`globtour-gps-server`

Upload all files from this folder to the repository.

## 2. Railway
Create a new Railway project and choose **Deploy from GitHub Repo**.
Select `globtour-gps-server`.

Railway will build the included Dockerfile.

## 3. HTTP dashboard
Railway service -> Settings -> Networking -> Public Networking -> Generate Domain.

The dashboard will then be available on the generated Railway URL.

## 4. TCP for FMC150
Railway service -> Settings -> Networking -> TCP Proxy.

Create a TCP proxy for the internal port:

`9000`

Railway will provide a TCP hostname and port, for example:

`xxxxx.proxy.rlwy.net:12345`

These two values are what go into FMC150:
- Server Domain = xxxxx.proxy.rlwy.net
- Port = 12345
- Protocol = TCP

Do NOT use the HTTP Railway domain for the FMC150 TCP connection.

## 5. Important storage note
The MVP currently uses SQLite. Railway service storage is ephemeral, so this is suitable for initial testing only.
For production, add a PostgreSQL database and migrate the GPS data to PostgreSQL.
Alternatively attach a Railway Volume for SQLite.

## 6. First test
Do not change FMC150 Server Settings until Railway shows:
- deployment successful
- HTTP dashboard opens
- TCP Proxy is created for internal port 9000

Then enter the generated TCP Proxy hostname/port in FMC150.
