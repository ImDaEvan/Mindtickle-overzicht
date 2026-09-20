# Mindtickle voortgangsrapport (MVP)

Genereert automatisch een opgemaakt Excel-overzicht van de Mindtickle-voortgang
van je team, inclusief kleurcodering en een "boven de X%"-check.

## 1. Installeren

```bash
cd mindtickle-report
python -m venv venv
venv\Scripts\activate        # Windows
# source venv/bin/activate   # macOS/Linux

pip install -r requirements.txt
```

## 2. Config invullen

Kopieer `.env.example` naar `.env`:

```bash
copy .env.example .env        # Windows
# cp .env.example .env        # macOS/Linux
```

Vul in `.env` het volgende in:

- **MINDTICKLE_SUBDOMAIN** — het stukje voor `.mindtickle.com` in je URL.
  Bij jou is dat `directresultmarketing`.
- **MINDTICKLE_MANAGER_ID** — staat al goed ingevuld (`1d01175d93e48000`),
  tenzij je een ander teamoverzicht wil ophalen.
- **MINDTICKLE_COOKIE** — zie stap 3 hieronder.
- **MINDTICKLE_HAR** — optioneel pad naar een geëxporteerde HAR uit Chrome.
  Als `MINDTICKLE_COOKIE` leeg is, leest het script de members-request en de
  cookie uit dit bestand.
- **SCORE_THRESHOLD** — de drempel in procenten (standaard 50).

## 3. Je cookie ophalen (nodig omdat Mindtickle geen publieke API-key geeft)

Dit script gebruikt dezelfde sessie als je browser, dus je moet elke keer dat
je sessie verloopt een verse cookie plakken:

1. Log in op Mindtickle en ga naar het teamoverzicht (Members-pagina).
2. Open de Developer Tools (F12) → tab **Network**.
3. Ververs de pagina en klik op het verzoek dat naar `.../members?...` gaat.
4. Scroll in de **Request Headers** naar `cookie:` en kopieer de **hele
   waarde** (dat is één lange regel met puntkomma's).
5. Plak die als waarde van `MINDTICKLE_COOKIE=` in je `.env` (op één regel,
   zonder aanhalingstekens).

⚠️ Deze cookie is persoonlijk en tijdelijk geldig — deel hem met niemand. Een
HAR bevat vaak dezelfde cookie en moet dus ook privé blijven. Zet `.env` en HAR-
bestanden nooit in git.

## 4. Draaien

```bash
python mindtickle_report.py
```

Met een HAR waarin cookies zijn meegeschreven kan het ook zonder de cookie
handmatig te plakken:

```env
MINDTICKLE_HAR=C:\Users\evanr\Downloads\directresultmarketing.mindtickle.com.har 07092026.har
```

Laat `MINDTICKLE_COOKIE` dan leeg. Het script gebruikt eerst expliciete waarden
uit `.env`; de HAR is een fallback. Let op: de aangeleverde HAR van 07-09-2026
bevat geen cookie of andere authenticatieheader. Gebruik daarom voor die HAR nog
steeds `MINDTICKLE_COOKIE`, of exporteer hem opnieuw met request cookies.

Je krijgt een bestand `ouput/voortgang_overzicht_<datum>.xlsx`, klaar om te
delen. De map `ouput` wordt automatisch aangemaakt als die nog niet bestaat.

## Bekende beperkingen (MVP)

- **Cookie verloopt.** Zodra je een `401`/`403`-foutmelding krijgt, moet je
  stap 3 herhalen. Er is geen automatisch inloggen ingebouwd — dat vraagt om
  een officiële API-key/OAuth-toegang bij Mindtickle, mocht dat beschikbaar
  zijn voor jullie account.
- **Eén manager-team per run.** Wil je meerdere teams combineren, voeg dan
  een lijst van manager-ID's toe en loop daaroverheen (kan ik voor je
  uitbreiden).
- Namen met een team-tag zoals `[V]Jan Jansen` worden automatisch
  schoongemaakt naar `Jan Jansen`.

## Volgende stappen (optioneel, later)

- Automatisch elke maandagochtend laten draaien (Windows Taakplanner / cron)
  en het bestand direct naar een vaste map of e-mail sturen.
- Rechtstreeks naar de groepsapp posten via een webhook, i.p.v. handmatig
  het Excel-bestand te versturen.
