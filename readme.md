PROGRAM: QGIS_Artfynd_Karta

ÖVERSIKT
Detta Python-script är avsett att köras i QGIS (PyQGIS) och automatiserar
processen att skapa en karta över artobservationer. Programmet läser in
artfynd från en Excel-fil, kopplar samman information om rödlistningsstatus,
symboliserar fyndpunkter och skapar en kartlayout med etiketter.

Scriptet används främst för naturvårdsanalys och dokumentation av
artobservationer, där det är viktigt att snabbt kunna generera tydliga
kartor från fältdata.

------------------------------------------------------------
HUVUDFUNKTIONER
------------------------------------------------------------

1. INLÄSNING AV DATA
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

------------------------------------------------------------

2. SYMBOLISERING
Punkter symboliseras efter rödlistningsklass:

CR - röd  
EN - orange  
VU - gul  
NT - blå  
DD - grå  
LC - lila  
S  - grön

Symbolnivåer används så att mer hotade arter ritas ovanpå mindre hotade.

------------------------------------------------------------

3. ETIKETTER
Varje punkt kan få en etikett (ArtNr).

Två etikettlägen finns:

A) Standardläge  
Etiketter ritas direkt vid punkten.

B) Täthetsbaserat läge (valbart)  
Programmet analyserar punkttäthet och delar upp punkter i två grupper:

    dense   = punkter med minst X grannar inom R meter
    sparse  = övriga punkter

dense-punkter:
    - etikett flyttas från punkten
    - ledarlinje (callout) ritas till punkt

sparse-punkter:
    - etikett visas direkt vid punkten

Syftet är att göra kartan läsbar även när flera observationer ligger nära
varandra.

------------------------------------------------------------

4. TÄTHETSBERÄKNING
Täthet beräknas genom:

1. Buffert runt varje punkt (radie R meter)
2. Räkna antal punkter i varje buffert
3. Resultatet sparas i fältet:

    pt_count

pt_count inkluderar punkten själv.

Standardvärden:
    R = 10 meter
    dense om pt_count >= 2

------------------------------------------------------------

5. KARTLAYOUT
Programmet skapar automatiskt en QGIS-layout som innehåller:

- bakgrundskarta
- fyndpunkter
- etiketter
- legend
- karttitel

Layouten kan exporteras eller redigeras manuellt efter körning.

------------------------------------------------------------

PROGRAMFLÖDE
------------------------------------------------------------

1. Användaren väljer rasterkarta
2. Användaren väljer Excel-fil med artfynd
3. Artdata konverteras till punktlager
4. ArtNr genereras eller används om det redan finns
5. Punkter symboliseras efter rödlistningsklass
6. (Valfritt) täthetsanalys körs
7. Punkter delas i dense / sparse
8. Etiketter skapas
9. Kartlayout genereras

------------------------------------------------------------

VIKTIGA DESIGNVAL
------------------------------------------------------------

- Scriptet använder PyQGIS-processing där möjligt.
- Mellanresultat sparas till temporära eller projektfiler.
- Renderer klonas för att undvika att C++-objekt i QGIS raderas.
- Scriptet är skrivet för QGIS 3.x.

------------------------------------------------------------

MÖJLIGA FÖRBÄTTRINGAR
------------------------------------------------------------

Följande utveckling är möjlig:

- automatisk skalning av etikettavstånd
- bättre klusterdetektion
- stöd för flera artlistor
- export till PDF automatiskt
- GUI istället för dialogrutor
- konvertering till QGIS-plugin

------------------------------------------------------------

ANVÄNDNING
------------------------------------------------------------

Scriptet körs från QGIS Python-konsol eller Script Editor.

Krav:
    QGIS 3.x
    PyQGIS
    Excel-fil med artfynd
    koordinater i SWEREF 99

------------------------------------------------------------
