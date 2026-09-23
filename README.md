# GLOBTOUR GPS SERVER — MVP

Ovaj projekt prima Teltonika FMC150 AVL podatke preko TCP-a i prikazuje zadnju poziciju vozila na web karti.

## Portovi
- TCP GPS listener: 9000
- Web dashboard/API: 8000

## Pokretanje na Windowsu

```powershell
cd C:\putanja\globtour_gps_server
py -3 -m venv .venv
.\.venv\Scripts\Activate
pip install -r requirements.txt
python server.py
```

Dashboard:
http://127.0.0.1:8000

GPS TCP server:
0.0.0.0:9000

## FMC150
Nakon što server bude dostupan s Interneta:
- GPRS Context: Enable
- APN: iot.1nce.net
- Authentication: Normal (PAP)
- Server Domain: javna IP adresa ili DNS servera
- Port: 9000
- Protocol: TCP

U FMC150 System -> Data Protocol preporučeno je koristiti Codec 8 Extended.

## Napomena
Za prvi lokalni test računalo mora biti dostupno s Interneta na TCP portu 9000 (npr. port forwarding na routeru) ili server treba biti postavljen na VPS. Za produkciju preporučujemo VPS s javnom IP adresom.
