# QGIS_Artfynd_Karta

## Översikt

Detta Python-script är avsett att köras i QGIS (PyQGIS) och automatiserar
processen att skapa en karta över artobservationer. Programmet läser in
artfynd från en Excel-fil, kopplar samman information om rödlistningsstatus,
symboliserar fyndpunkter och skapar en kartlayout med etiketter.

Scriptet används främst för naturvårdsanalys och dokumentation av
artobservationer, där det är viktigt att snabbt kunna generera tydliga
kartor från fältdata.

## Du behöver

- QGIS 3.40 rekommenderas
- Python-filen (`.py`) från detta repository
- En bakgrundskarta i GeoTIFF-format (`.tif`)
- En Excel-fil med artobservationer (`.xlsx`).
  - Koordinatsystem SWEREF 99 TM (EPSG:3006).
  - `Artnamn`, `Rödlistade`, `Ost` och `Nord` måste finnas.

Excel-filen kan vara en export från Artportalen och ska innehålla de kolumner som skriptet behöver, bland annat artnamn, rödlistningsklass och koordinater i SWEREF 99.

### Excel-stöd (`openpyxl`)

Skriptet använder Python-paketet `openpyxl` för att läsa Excel-filer (`.xlsx`). Paketet finns redan i vissa QGIS-installationer.

Om skriptet meddelar att `openpyxl` saknas:
Windows:
1. Öppna **Start-menyn** i Windows och sök efter **OSGeo4W Shell**.
2. Starta **OSGeo4W Shell**.
3. Kör: `python -m pip install openpyxl`
4. Starta om QGIS och kör skriptet igen.

macOS/Linux: Öppna en terminal med QGIS Python-miljö aktiverad och kör: `python3 -m pip install openpyxl`


## Så här kör du skriptet

1. Ladda ner Python-filen (`QGIS_Artfynd_ på_karta.py`) från detta repository: `https://github.com/TomasFalk64/QGIS_Artfyndkarta`
2. Starta QGIS.
3. Öppna **Tillägg → Python-konsol**.
4. Öppna **Script Editor** från Python-konsolen.
5. Öppna den nedladdade `.py`-filen i Script Editor. Du kan också klistra in skriptets innehåll direkt i redigeraren.
6. Klicka på **Kör** (den gröna play-symbolen).
7. Välj den GeoTIFF-fil som ska användas som bakgrundskarta.
8. Välj Excel-filen med artfynd när du får frågan.

Skriptet läser in fynden, skapar punktlager, symboliserar arterna och genererar kartlayouten.

Den färdiga kartan kan därefter granskas och justeras i QGIS innan den exporteras, exempelvis som PDF eller bild.

## Huvudfunktioner

### 1. Inläsning av data

Programmet läser två typer av data:

- Rasterkarta (bakgrundskarta)
- Excel-fil med artfynd

Excel-filen innehåller normalt följande kolumner:

- Artnamn
- Vetenskapligt namn
- Rödlistningsklass
- Koordinater (SWEREF 99)
- Datum
- ArtNr (valfritt)

Om ArtNr saknas genererar scriptet ett stabilt artnummer sorterat på
rödlistningsklass (sällsynt först) och artnamn (bokstavsordning).
Om ArtNr finns och har värden används de värdena.

### 2. Symbolisering

Punkter symboliseras efter rödlistningsklass:

- CR – röd
- EN – orange
- VU – gul
- NT – blå
- DD – grå
- LC – lila
- S – grön

Symbolnivåer används så att mer hotade arter ritas ovanpå mindre hotade.

### 3. Etiketter

Varje punkt kan få en etikett (ArtNr).

Två etikettlägen finns:

#### A) Standardläge

Etiketter ritas direkt vid punkten.

#### B) Täthetsbaserat läge (valbart)

Programmet analyserar punkttäthet och delar upp punkter i två grupper:

- `dense` = punkter med minst X grannar inom R meter
- `sparse` = övriga punkter

**dense-punkter:**

- etikett flyttas från punkten
- ledarlinje (callout) ritas till punkt

**sparse-punkter:**

- etikett visas direkt vid punkten

Syftet är att göra kartan läsbar även när flera observationer ligger nära
varandra.

### 4. Täthetsberäkning

Täthet beräknas genom:

1. Buffert runt varje punkt (radie R meter)
2. Räkna antal punkter i varje buffert
3. Resultatet sparas i fältet `pt_count`.

`pt_count` inkluderar punkten själv.

Standardvärden:

- R = 10 meter
- `dense` om `pt_count >= 2`

### 5. Kartlayout

Programmet skapar automatiskt en QGIS-layout som innehåller:

- bakgrundskarta
- fyndpunkter
- etiketter
- legend
- karttitel

Layouten kan exporteras eller redigeras manuellt efter körning.

## Programflöde

1. Användaren väljer rasterkarta
2. Användaren väljer Excel-fil med artfynd
3. Artdata konverteras till punktlager
4. ArtNr genereras eller används om det redan finns
5. Punkter symboliseras efter rödlistningsklass
6. (Valfritt) täthetsanalys körs
7. Punkter delas i dense / sparse
8. Etiketter skapas
9. Kartlayout genereras

## Viktiga designval

- Scriptet använder PyQGIS-processing där möjligt.
- Mellanresultat sparas till temporära eller projektfiler.
- Renderer klonas för att undvika att C++-objekt i QGIS raderas.
- Scriptet är skrivet för QGIS 3.x.

## Möjliga förbättringar

Följande utveckling är möjlig:

- automatisk skalning av etikettavstånd
- bättre klusterdetektion
- stöd för flera artlistor
- export till PDF automatiskt
- GUI istället för dialogrutor
- konvertering till QGIS-plugin

## Användning

Scriptet körs från QGIS Python-konsol eller Script Editor.

Krav:

- QGIS 3.x
- PyQGIS
- Excel-fil med artfynd
- koordinater i SWEREF 99
