# Charge Log — GitHub Pages PWA

Mobil webbloggbok för laddningar av din elbil, anpassad bland annat för iPhone 14 Pro Max (430 × 932), Dynamic Island och safe areas. Fungerar som en statisk webbplats utan backend: all data sparas enbart i webbläsarens `localStorage`. Repot innehåller bara appens kod — inga personliga data (kvitton, skärmdumpar, Excel med laddningar) finns eller kommer att finnas här: du laddar upp din egen logg via knappen **Importera** efter att du öppnat sajten, och datan hamnar aldrig i GitHub.

## Funktioner

- mobilanpassat gränssnitt och installation på hemskärmen som PWA;
- lägg till, redigera, ta bort och sök bland laddsessioner;
- eget fordonsnamn: byt namn under **Fordon** i inställningsmenyn — det används i rubriken samt i filnamn och rubriker vid export (standard: "XPENG G9");
- inbyggd hjälp: en **Hjälp & vanliga frågor**-dialog i inställningsmenyn förklarar fälten, import/export och var datan sparas;
- automatiska beräkningar: laddad SOC, snitteffekt, pris per kWh, tillagd räckvidd, kostnad / 100 km, kWh / 1 % SOC;
- analys för denna/förra månaden, i år och alla tider, med jämförelse mot föregående månads kostnad;
- daglig statistik: körsträcka, förbrukning, hastighet, regenerering, parkering och A/C (om sådan data importerats);
- kostnad / 100 km räknas på tillagd räckvidd från själva laddsessionerna — detta täcker alltid samma period som pengarna som spenderats; den faktiska dagliga körsträckan (om importerad) visas separat som referens;
- import av XLSX / CSV / JSON — enda sättet att lägga in egna data, repot innehåller inga;
- export av XLSX / CSV / JSON;
- säker återställning: innan lokal data raderas laddar appen alltid ner fullständiga kopior i JSON och XLSX;
- påminnelse om säkerhetskopia var 30:e dag;
- inget server- eller databasberoende — passar för GitHub Pages.

## Publicera på GitHub Pages

1. Skapa ett nytt GitHub-repo, till exempel `charge-log`.
2. Ladda upp alla filer och mappar från det här projektet till repots rot.
3. Öppna **Settings → Pages**.
4. Under **Build and deployment**, välj **Deploy from a branch**.
5. Branch: `main`, folder: `/ (root)`, klicka sedan **Save**.
6. Öppna den publicerade adressen via GitHub Pages. På iPhone/Android går det att lägga till sajten på hemskärmen.
7. Tryck på **Importera** och välj din fil (XLSX / CSV / JSON) — den stannar bara i webbläsarens/enhetens `localStorage` och skickas ingenstans. Committa aldrig en fil med riktig data till repot — mappen `data/` finns redan i `.gitignore`.

## Installation på iPhone 14 Pro Max

1. Öppna GitHub Pages-adressen i just Safari.
2. Tryck på **Dela**-knappen längst ner på skärmen.
3. Välj **Lägg till på hemskärmen** → **Lägg till**.
4. Starta appen från hemskärmen — den öppnas utan Safaris gränssnitt, som en vanlig app.

Gränssnittet tar hänsyn till det övre området med Dynamic Island och den nedre Home Indicator. Fälten är dimensionerade så att iOS inte zoomar in sidan vid inmatning.

## XLSX och offlineläge

SheetJS CE 0.20.3 finns lokalt i `vendor/xlsx.full.min.js` och cachas av PWA:n tillsammans med appen. Efter första öppningen på GitHub Pages fungerar import och export av XLSX utan internetuppkoppling.

## Data och integritet

GitHub-repot innehåller bara appens kod. Personliga data (Excel med laddningar, kvitton, skärmdumpar) committas aldrig — mappen `data/` och filen `XPENG-G9-Charge-Log-iPhone.zip` finns listade i `.gitignore`. Du laddar själv upp din logg via **Importera** efter att ha öppnat den publicerade sajten; därefter sparas datan enbart i den aktuella webbläsarens/enhetens `localStorage` och skickas varken till GitHub eller någon annan server. Använd Exportera → JSON eller XLSX regelbundet för säkerhetskopiering.

Vill du se om någon besöker sajten kan du använda GitHub:s inbyggda **Insights → Traffic** för repot — den kräver ingen kodändring och skickar ingen data till tredje part.
