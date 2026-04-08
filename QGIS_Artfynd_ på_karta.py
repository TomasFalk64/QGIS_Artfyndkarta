import os
import uuid
import processing
import csv
import time

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None

from qgis.PyQt.QtCore import Qt, QVariant
from qgis.PyQt.QtGui import QColor, QFont
from qgis.PyQt.QtWidgets import QFileDialog, QInputDialog
from qgis.core import (
    Qgis,
    QgsApplication,
    QgsCategorizedSymbolRenderer,
    QgsCoordinateReferenceSystem,
    QgsFields,
    QgsField,
    QgsFeature,
    QgsFillSymbol,
    QgsLayoutFrame,
    QgsLayoutItem,
    QgsLayoutItemAttributeTable,
    QgsLayoutItemLabel,
    QgsLayoutItemLegend,
    QgsLayoutItemMap,
    QgsLayoutItemPage,
    QgsLayoutItemPicture,
    QgsLayoutItemScaleBar,
    QgsLayoutItemShape,
    QgsLayoutPoint,
    QgsLayoutSize,
    QgsLayoutTableColumn,
    QgsLegendStyle,
    QgsLineSymbol,
    QgsMarkerSymbol,
    QgsPalLayerSettings,
    QgsPrintLayout,
    QgsProject,
    QgsProviderRegistry,
    QgsRasterLayer,
    QgsRendererCategory,
    QgsSingleSymbolRenderer,
    QgsSimpleLineCallout,
    QgsTextFormat,
    QgsUnitTypes,
    QgsVectorLayer,
    QgsVectorLayerSimpleLabeling
    
)

print("Script loaded successfully.")
print(f"__name__ = {__name__}")

# ----------------------------
# SETTINGS (defaults)
# ----------------------------
EPSG = 3006
X_FIELD = "Ost"
Y_FIELD = "Nord"
REDLIST_FIELD = "Rödlistade"
SPECIES_FIELD = "Artnamn"
ARTNR_FIELD = "ArtNr"
KNAEROT_NAME = "Knärot"
BUFFER_DIST = 50
LAYOUT_MAP_WIDTH_MM = 230
OUT_DIR = r"C:\GIS\output"
os.makedirs(OUT_DIR, exist_ok=True)
STABLE_FINAL_OUTPUTS = True
USE_DENSE_LABELS = True

# Global variables
project = None
root = None
run_id = None
RASTER_PATH = None
TABLE_PATH = None
SHEET_NAME = None
raster = None
excel_table = None
points = None
points_copy = None
knaerot_pts = None
knaerot_buf = None
knaerot_diss = None
has_knaerot = False
cats = []
renderer_template = None

# ----------------------------
# Helper functions
# ----------------------------
def p(name):
    return os.path.join(OUT_DIR, f"{name}_{run_id}.gpkg")

def final(name):
    if STABLE_FINAL_OUTPUTS:
        return os.path.join(OUT_DIR, f"{name}.gpkg")
    return p(name)

def add_layer(layer, name=None):
    if name:
        layer.setName(name)
    project.addMapLayer(layer)
    return layer

def require_valid(layer, msg):
    if layer is None or not layer.isValid():
        raise Exception(msg)

def move_to_top(layer):
    if layer is None:
        return

    project = QgsProject.instance()
    lyr = project.mapLayer(layer.id())
    if lyr is None:
        return  # layer no longer in project

    node = project.layerTreeRoot().findLayer(lyr.id())
    if node is None:
        return

    try:
        parent = node.parent()
        if parent is None:
            return
        # Remove + insert at top
        parent.removeChildNode(node)
        parent.insertChildNode(0, node)
    except RuntimeError:
        # Node got deleted/invalidated; just skip
        return

def apply_simple_labels(layer, field_name, size=8):
    pal = QgsPalLayerSettings()
    pal.enabled = True
    pal.fieldName = field_name
    pal.isExpression = False

    # QGIS 3.40: pal.placement vill ha enumen QgsPalLayerSettings.Placement
    placement_enum = getattr(QgsPalLayerSettings, "Placement", None) or getattr(QgsPalLayerSettings, "LabelPlacement", None)
    if placement_enum is not None:
        # Försök "OverPoint", annars "AroundPoint"
        if hasattr(placement_enum, "OverPoint"):
            pal.placement = placement_enum.OverPoint
        else:
            pal.placement = placement_enum.AroundPoint
    else:
        # sista fallback: låt QGIS default placement gälla
        pass

    fmt = QgsTextFormat()
    fmt.setSize(size)
    pal.setFormat(fmt)

    layer.setLabeling(QgsVectorLayerSimpleLabeling(pal))
    layer.setLabelsEnabled(True)
    layer.triggerRepaint()

def excel_to_clean_csv(xlsx_path, sheet_name, out_csv_path, required_headers):
    if load_workbook is None:
        raise Exception("Saknar beroende: openpyxl. Installera openpyxl i QGIS Python-miljo.")
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb[sheet_name]

    header_row_idx = None
    header = None

    # leta header inom första 30 rader
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), start=1):
        if not row:
            continue
        vals = [str(v).strip() if v is not None else "" for v in row]
        if all(h in vals for h in required_headers):
            header_row_idx = i
            header = vals
            break

    if header_row_idx is None:
        wb.close()
        raise Exception(f"Kunde inte hitta header-rad i {sheet_name}. Letade efter: {required_headers}")

    # skriv CSV från header och nedåt
    with open(out_csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=",")
        w.writerow(header)
        for row in ws.iter_rows(min_row=header_row_idx+1, values_only=True):
            if not row:
                continue
            w.writerow([("" if v is None else v) for v in row])

    wb.close()
    
def remove_layers_pointing_to(path):
    proj = QgsProject.instance()
    target_norm = os.path.normcase(os.path.normpath(path))
    target_base = os.path.normcase(os.path.basename(path))

    to_remove = []
    for lyr in proj.mapLayers().values():
        src = (lyr.source() or "")
        src_norm = os.path.normcase(os.path.normpath(src))

        # matcha antingen full path eller bara filnamn (för |layername=...)
        if target_norm in src_norm or target_base in os.path.normcase(src):
            to_remove.append(lyr.id())

    for lid in to_remove:
        proj.removeMapLayer(lid)

def delete_if_exists(path, retries=25, wait_s=0.25):
    if not path or not os.path.exists(path):
        return

    remove_layers_pointing_to(path)

    try:
        QgsProviderRegistry.instance().cleanProviderCaches()
    except Exception:
        pass

    last_err = None
    for _ in range(retries):
        try:
            os.remove(path)
            return
        except PermissionError as e:
            last_err = e
            time.sleep(wait_s)

    raise last_err

def resolve_field_name(layer, wanted_name):
    idx = layer.fields().indexFromName(wanted_name)
    if idx != -1:
        return layer.fields()[idx].name()

    wanted_norm = str(wanted_name).strip().casefold()
    for fld in layer.fields():
        if fld.name().strip().casefold() == wanted_norm:
            return fld.name()
    return None

def layer_has_field_with_values(layer, field_name, sample_limit=None):
    resolved = resolve_field_name(layer, field_name)
    if not resolved:
        return False
    idx = layer.fields().indexFromName(resolved)
    if idx == -1:
        return False
    n = 0
    for f in layer.getFeatures():
        v = f[resolved]
        if v is not None and str(v).strip() != "":
            return True
        n += 1
        if sample_limit is not None and n >= sample_limit:
            break
    return False

def apply_callout_labels(layer, field_name, dist_mm=2, size=8):
    pal = QgsPalLayerSettings()
    pal.enabled = True
    pal.fieldName = field_name

    # Placera etiketter runt punkten
    pal.placement = Qgis.LabelPlacement.AroundPoint
    pal.dist = dist_mm  # mm

    fmt = QgsTextFormat()
    fmt.setFont(QFont("Arial"))
    fmt.setSize(size)
    pal.setFormat(fmt)

    callout = QgsSimpleLineCallout()
    callout.setEnabled(True)

    # (valfritt) sätt linjesymbol för callouten
    line_symbol = QgsLineSymbol.createSimple({
        "color": "0,0,0,180",
        "width": "0.2"
    })
    callout.setLineSymbol(line_symbol)

    pal.setCallout(callout)

    layer.setLabeling(QgsVectorLayerSimpleLabeling(pal))
    layer.setLabelsEnabled(True)

def move_layer_name_to_project_top(layer_name: str) -> bool:
    root = QgsProject.instance().layerTreeRoot()

    # Hitta layer-id först (utan att hålla kvar tree-node referenser)
    target_layer = None
    for lyr in QgsProject.instance().mapLayers().values():
        if lyr.name() == layer_name:
            target_layer = lyr
            break
    if target_layer is None:
        return False

    # Hitta noden på nytt och klona den
    node = root.findLayer(target_layer.id())
    if node is None:
        return False

    try:
        clone = node.clone()
        # Lägg klonen överst i root
        root.insertChildNode(0, clone)
        # Ta bort originalnoden (från dess parent)
        parent = node.parent()
        if parent is not None:
            parent.removeChildNode(node)
        return True
    except RuntimeError:
        return False

def move_layer_to_bottom(layer_name: str) -> bool:
    root = QgsProject.instance().layerTreeRoot()

    # Hitta layer-id först (utan att hålla kvar tree-node referenser)
    target_layer = None
    for lyr in QgsProject.instance().mapLayers().values():
        if lyr.name() == layer_name:
            target_layer = lyr
            break
    if target_layer is None:
        return False

    # Hitta noden på nytt och klona den
    node = root.findLayer(target_layer.id())
    if node is None:
        return False

    try:
        clone = node.clone()
        # Lägg klonen underst i root
        root.insertChildNode(len(root.children()), clone)
        # Ta bort originalnoden (från dess parent)
        parent = node.parent()
        if parent is not None:
            parent.removeChildNode(node)
        return True
    except RuntimeError:
        return False

def build_unique_species_table(points_layer, artnr_field, red_field, name_field):
    mem = QgsVectorLayer("Point?crs=EPSG:3006", "Artlista (unik)", "memory")
    pr = mem.dataProvider()
    pr.addAttributes([
        QgsField("Artnr", QVariant.Int),
        QgsField("Rödlistade", QVariant.String),
        QgsField("Artnamn", QVariant.String),
    ])
    mem.updateFields()

    seen = {}
    for f in points_layer.getFeatures():
        artnr = f[artnr_field]
        red = f[red_field]
        art = f[name_field]
        key = (int(artnr) if artnr is not None else None, str(red or ""), str(art or ""))
        if key[0] is None:
            continue
        seen[key] = True

    rows = sorted(seen.keys(), key=lambda t: (t[0], t[2]))

    feats = []
    for (artnr, red, art) in rows:
        nf = QgsFeature(mem.fields())
        nf.setAttributes([artnr, red, art])
        # geometry not used in table, but provider requires one
        nf.setGeometry(points_layer.getFeature(next(points_layer.getFeatures()).id()).geometry())
        feats.append(nf)

    pr.addFeatures(feats)
    mem.updateExtents()
    QgsProject.instance().addMapLayer(mem, False)  # add but not visible
    return mem

# ----------------------------
# Main functions
# ----------------------------
def initialize_globals():
    global project, root, run_id
    project = QgsProject.instance()
    root = project.layerTreeRoot()
    run_id = uuid.uuid4().hex[:8]

def ask_user_for_files():
    global RASTER_PATH, TABLE_PATH, SHEET_NAME
    print('Välj rasterkarta')
    RASTER_PATH, _ = QFileDialog.getOpenFileName(
        None,
        "Välj rasterkarta",
        "",
        "Raster (*.tif *.tiff *.jpg *.jpeg *.png);;Alla filer (*.*)"
    )
    if not RASTER_PATH:
        raise Exception("Avbrutet: ingen raster valdes.")

    print('Välj excelfil')
    TABLE_PATH, _ = QFileDialog.getOpenFileName(
        None,
        "Välj Excel-fil",
        "",
        "Excel (*.xlsx);;Alla filer (*.*)"
    )
    if not TABLE_PATH:
        raise Exception("Avbrutet: ingen Excel-fil valdes.")

    # Pick sheet (tries to read sheetnames)
    SHEET_NAME = None
    try:
        if load_workbook is None:
            raise ImportError("openpyxl saknas")
        wb = load_workbook(TABLE_PATH, read_only=True, data_only=True)
        sheets = wb.sheetnames
        wb.close()

        if sheets:
            default_index = 0
            if "Test" in sheets:
                default_index = sheets.index("Test")
            chosen, ok = QInputDialog.getItem(
                None,
                "Välj blad i Excel",
                "Blad:",
                sheets,
                default_index,
                False
            )
            SHEET_NAME = chosen if ok and chosen else sheets[default_index]
    except Exception:
        # fallback: user must type sheet name or keep default
        chosen, ok = QInputDialog.getText(None, "Bladnamn", "Skriv bladnamn (t.ex. Test):")
        if ok and chosen.strip():
            SHEET_NAME = chosen.strip()

    if not SHEET_NAME:
        SHEET_NAME = "Test"  # final fallback

def load_raster():
    global raster
    raster = QgsRasterLayer(RASTER_PATH, os.path.basename(RASTER_PATH))
    require_valid(raster, f"Raster could not be loaded: {RASTER_PATH}")
    add_layer(raster, os.path.basename(RASTER_PATH))

def process_excel_to_table():
    global excel_table
    # Bygg en rensad CSV från Excel (hoppar metadata-rader automatiskt)
    clean_csv = os.path.join(OUT_DIR, f"clean_{run_id}.csv")
    excel_to_clean_csv(
        TABLE_PATH,
        SHEET_NAME,
        clean_csv,
        required_headers=[X_FIELD, Y_FIELD, SPECIES_FIELD, REDLIST_FIELD]
    )

    # Läs CSV istället för XLSX
    table_uri = f"file:///{clean_csv.replace('\\','/')}" \
                f"?delimiter=," \
                f"&crs=epsg:{EPSG}"

    excel_table = QgsVectorLayer(table_uri, "Excel (rensad CSV)", "delimitedtext")
    require_valid(excel_table, f"Could not load cleaned CSV as table: {clean_csv}")
    add_layer(excel_table)

    try:
        root.findLayer(excel_table.id()).setItemVisibilityChecked(False)
    except Exception:
        pass

    # Validate fields
    field_names = [f.name() for f in excel_table.fields()]
    missing = [f for f in [X_FIELD, Y_FIELD, REDLIST_FIELD, SPECIES_FIELD] if f not in field_names]
    if missing:
        raise Exception(f"Missing fields in Excel: {missing}\nAvailable: {field_names}")

def clean_and_create_points():
    global points, ARTNR_FIELD
    # Clean table: keep only numeric Ost/Nord
    clean_expr = f"to_real(\"{X_FIELD}\") IS NOT NULL AND to_real(\"{Y_FIELD}\") IS NOT NULL"
    clean_table_path = p("excel_clean")

    processing.run("native:extractbyexpression", {
        "INPUT": excel_table,
        "EXPRESSION": clean_expr,
        "OUTPUT": clean_table_path
    })

    clean_table = QgsVectorLayer(clean_table_path, "Excel (rensad)", "ogr")

    # Fill missing redlist values: NULL/empty -> "LC" (on the cleaned table) ---
    clean_table_lc_path = p("excel_clean_lc")
    processing.run("native:fieldcalculator", {
        "INPUT": clean_table,
        "FIELD_NAME": REDLIST_FIELD,
        "FIELD_TYPE": 2,          # String
        "FIELD_LENGTH": 10,
        "FIELD_PRECISION": 0,
        "FORMULA": f"coalesce(NULLIF(trim(\"{REDLIST_FIELD}\"), ''), 'LC')",
        "OUTPUT": clean_table_lc_path
    })

    clean_table = QgsVectorLayer(clean_table_lc_path, "Excel (rensad + LC)", "ogr")
    require_valid(clean_table, f"Could not load cleaned table with LC: {clean_table_lc_path}")
    add_layer(clean_table)

    try:
        root.findLayer(clean_table.id()).setItemVisibilityChecked(False)
    except Exception:
        pass

    # Create points (EPSG:3006)
    target_crs = QgsCoordinateReferenceSystem(f"EPSG:{EPSG}")
    points_raw_path = p("punkter_raw")

    processing.run("native:createpointslayerfromtable", {
        "INPUT": clean_table,
        "XFIELD": X_FIELD,
        "YFIELD": Y_FIELD,
        "ZFIELD": None,
        "MFIELD": None,
        "TARGET_CRS": target_crs,
        "OUTPUT": points_raw_path
    })

    points_raw = QgsVectorLayer(points_raw_path, "Punkter (raw)", "ogr")
    require_valid(points_raw, f"Could not load points_raw: {points_raw_path}")
    add_layer(points_raw)

    try:
        root.findLayer(points_raw.id()).setItemVisibilityChecked(False)
    except Exception:
        pass

    # Matcha ArtNr-fält robust (exakt + case/whitespace-insensitive)
    resolved_artnr = resolve_field_name(points_raw, ARTNR_FIELD)
    if resolved_artnr:
        ARTNR_FIELD = resolved_artnr

    # Stable species order + ArtNr (do NOT rely on provider order)
    if layer_has_field_with_values(points_raw, ARTNR_FIELD):
        points = points_raw
    else:
        # Rank order you want (threatened first)
        rank_map = {"CR": 1, "EN": 2, "VU": 3, "NT": 4, "LC": 5, "S": 6, "DD": 7}

        # Collect unique species rows from points_raw
        species_set = {}
        for f in points_raw.getFeatures():
            art = str(f[SPECIES_FIELD] or "").strip()
            rl = str(f[REDLIST_FIELD] or "LC").strip().upper()
            if not art:
                continue
            if rl == "":
                rl = "LC"
            species_set[art] = rl

        # Sort: by redlist rank, then by species name
        species_sorted = sorted(
            [(art, rl) for art, rl in species_set.items()],
            key=lambda t: (rank_map.get(t[1], 99), t[0])
        )

        # Build a memory layer species->ArtNr table
        species_layer = QgsVectorLayer("None", "Arter + ArtNr", "memory")
        pr = species_layer.dataProvider()
        pr.addAttributes([
            QgsField(SPECIES_FIELD, QVariant.String),
            QgsField(REDLIST_FIELD, QVariant.String),
            QgsField(ARTNR_FIELD, QVariant.Int),
        ])
        species_layer.updateFields()

        feats = []
        for i, (art, rl) in enumerate(species_sorted, start=1):
            nf = QgsFeature(species_layer.fields())
            nf.setAttributes([art, rl, i])
            feats.append(nf)
        pr.addFeatures(feats)

        # Save lookup table to disk (more stable for joins)
        species_with_nr_path = p("species_with_artnr")
        delete_if_exists(species_with_nr_path)  # viktigt: rensa ev. tidigare fil
        processing.run("native:savefeatures", {"INPUT": species_layer, "OUTPUT": species_with_nr_path})

        species_with_nr = QgsVectorLayer(species_with_nr_path, "Arter + ArtNr", "ogr")
        require_valid(species_with_nr, f"Could not load species_with_nr: {species_with_nr_path}")

        # Join ArtNr back to points
        points_final_path = final("punkter_med_artnr")
        delete_if_exists(points_final_path)  # rensa ev. tidigare
        processing.run("native:joinattributestable", {
            "INPUT": points_raw,
            "FIELD": SPECIES_FIELD,
            "INPUT_2": species_with_nr,
            "FIELD_2": SPECIES_FIELD,
            "FIELDS_TO_COPY": [ARTNR_FIELD],
            "METHOD": 1,
            "DISCARD_NONMATCHING": False,
            "PREFIX": "",
            "OUTPUT": points_final_path
        })

        points = QgsVectorLayer(points_final_path, "Rödlistningsklass", "ogr")
        require_valid(points, f"Could not load final points: {points_final_path}")

    # Optional: sort points layer by ArtNr for nicer attribute table browsing
    points_sorted_path = p("punkter_sorterade_artnr")
    delete_if_exists(points_sorted_path)
    processing.run("native:orderbyexpression", {
        "INPUT": points,
        "EXPRESSION": f"to_int(\"{ARTNR_FIELD}\")",
        "ASCENDING": True,
        "NULLS_FIRST": False,
        "OUTPUT": points_sorted_path
    })

    points = QgsVectorLayer(points_sorted_path, "Rödlistningsklass", "ogr")
    require_valid(points, f"Could not reload sorted points: {points_sorted_path}")
    add_layer(points)

def apply_symbology():
    global cats, renderer_template, points_copy
    # Symbology: color by rödlistning
    color_map = {
        "CR": QColor(220, 0, 0),
        "EN": QColor(255, 90, 0),
        "VU": QColor(240, 220, 70),
        "NT": QColor(100, 170, 255),
        "DD": QColor(160, 160, 160),
        "LC": QColor(190, 120, 220),
        "S":  QColor(70, 170, 70),
    }

    def make_cat(val, label, col, level):
        sym = QgsMarkerSymbol.createSimple({
            "name": "circle",
            "size": "3.4",
            "color": f"{col.red()},{col.green()},{col.blue()},255",
            "outline_color": "0,0,0,120",
            "outline_width": "0.2",
        })

        sl = sym.symbolLayer(0)
        if sl is not None:
            sl.setRenderingPass(level)

        return QgsRendererCategory(val, sym, label)

    levels = {"CR":7, "EN":6, "VU":5, "NT":4, "DD":3, "LC":2, "S":1}
    code_order = ["S", "LC", "DD", "NT", "VU", "EN", "CR"]
    present_codes = {
        str(f[REDLIST_FIELD] or "").strip().upper()
        for f in points.getFeatures()
    }
    cats = [
        make_cat(code, code, color_map[code], levels[code])
        for code in code_order
        if code in present_codes
    ]

    renderer_template = QgsCategorizedSymbolRenderer(REDLIST_FIELD, cats)
    renderer_template.setUsingSymbolLevels(True)

    points.setName("Rödlistningsklass")
    points.setRenderer(renderer_template.clone())
    points.triggerRepaint()

    points_copy_path = final("punkter_kopia")
    processing.run("native:savefeatures", {"INPUT": points, "OUTPUT": points_copy_path})

    points_copy = QgsVectorLayer(points_copy_path, "Rödlistningsklass (kopia)", "ogr")
    require_valid(points_copy, f"Could not load points copy: {points_copy_path}")
    add_layer(points_copy)

    # Hide the copy layer to avoid duplicate labels
    try:
        root.findLayer(points_copy.id()).setItemVisibilityChecked(False)
    except Exception:
        pass

    points_copy.setRenderer(renderer_template.clone())
    points_copy.triggerRepaint()

    # Labels: ArtNr
    points.setLabelsEnabled(False)
    points_copy.setLabelsEnabled(False)
    apply_simple_labels(points, ARTNR_FIELD, size=8)
    apply_simple_labels(points_copy, ARTNR_FIELD, size=8)

def create_knaerot_zones():
    global knaerot_pts, knaerot_buf, knaerot_diss, has_knaerot
    # 50 m circles around Knärot only + dissolve + blue fill
    knaerot_pts_path = final("knaerot_punkter")
    processing.run("native:extractbyexpression", {
        "INPUT": points,
        "EXPRESSION": f"\"{SPECIES_FIELD}\" = '{KNAEROT_NAME}'",
        "OUTPUT": knaerot_pts_path
    })
    knaerot_pts = QgsVectorLayer(knaerot_pts_path, "Knärot (punkter)", "ogr")
    require_valid(knaerot_pts, f"Could not load Knärot points: {knaerot_pts_path}")
    add_layer(knaerot_pts)

    try:
        root.findLayer(knaerot_pts.id()).setItemVisibilityChecked(False)
    except Exception:
        pass

    has_knaerot = knaerot_pts.featureCount() > 0

    knaerot_buf_path = final("knaerot_50m")
    buf_res = processing.run("native:buffer", {
        "INPUT": knaerot_pts,
        "DISTANCE": BUFFER_DIST,
        "SEGMENTS": 36,
        "END_CAP_STYLE": 0,
        "JOIN_STYLE": 0,
        "MITER_LIMIT": 2,
        "DISSOLVE": False,
        "OUTPUT": knaerot_buf_path
    })
    knaerot_buf = QgsVectorLayer(buf_res["OUTPUT"], f"Knärot {BUFFER_DIST} m", "ogr")
    require_valid(knaerot_buf, f"Could not load Knärot buffer: {buf_res['OUTPUT']}")
    add_layer(knaerot_buf)

    # Dissolve
    knaerot_diss_path = final("knaerot_50m_upplost")
    processing.run("native:dissolve", {
        "INPUT": knaerot_buf,
        "FIELD": [],
        "OUTPUT": knaerot_diss_path
    })
    knaerot_diss = QgsVectorLayer(knaerot_diss_path, f"Knärot {BUFFER_DIST} m skyddszon", "ogr")
    require_valid(knaerot_diss, f"Could not load dissolved layer: {knaerot_diss_path}")
    add_layer(knaerot_diss)

    # Style: light blue fill (~60% transparent) + outline
    diss_sym = QgsFillSymbol.createSimple({
        "outline_color": "0,0,0,220",
        "outline_width": "0.5",
        "color": "120,180,255,110"   # ~60% transparent fill
    })
    knaerot_diss.setRenderer(QgsSingleSymbolRenderer(diss_sym))
    knaerot_diss.triggerRepaint()

    # Hide non-dissolved buffers
    try:
        root.findLayer(knaerot_buf.id()).setItemVisibilityChecked(False)
    except Exception:
        pass

def arrange_layers():
    # Layer order
    move_layer_name_to_project_top("Rödlistningsklass")
    move_layer_to_bottom(os.path.basename(RASTER_PATH))

def create_print_layout():
    # Create / replace layout
    layout_name = "Artfynd"
    lm = QgsProject.instance().layoutManager()
    for l in lm.printLayouts():
        if l.name() == layout_name:
            lm.removeLayout(l)
            break

    layout = QgsPrintLayout(QgsProject.instance())
    layout.initializeDefaults()
    layout.setName(layout_name)
    lm.addLayout(layout)

    # A4 landscape
    page = layout.pageCollection().page(0)
    page.setPageSize("A4", QgsLayoutItemPage.Orientation.Landscape)

    ## ---- MAP item (left area) ----
    MAP_X, MAP_Y = 0, 0
    MAP_W, MAP_H = LAYOUT_MAP_WIDTH_MM, 210
    RIGHT_PANEL_GAP = 5
    RIGHT_PANEL_X = MAP_X + MAP_W + RIGHT_PANEL_GAP
    RIGHT_PANEL_W = max(40, 297 - RIGHT_PANEL_X - 7)

    map_item = QgsLayoutItemMap(layout)
    layout.addLayoutItem(map_item)

    # 1) Fixera storleken i mm 
    map_item.setReferencePoint(QgsLayoutItem.UpperLeft)
    map_item.setFixedSize(QgsLayoutSize(MAP_W, MAP_H, QgsUnitTypes.LayoutMillimeters))

    # 2) Placera objektet i mm
    map_item.attemptMove(
        QgsLayoutPoint(MAP_X, MAP_Y, QgsUnitTypes.LayoutMillimeters)
    )

    map_item.setFrameEnabled(True)

    # 3) Zooma ENDAST genom extent (påverkar inte storleken)
    ext = iface.mapCanvas().extent()
    map_item.setExtent(ext)
    map_item.setScale(iface.mapCanvas().scale())

    size = map_item.sizeWithUnits()
    pos = map_item.positionWithUnits()
    print("Map pos (mm):", pos.x(), pos.y(), " size (mm):", size.width(), size.height())

    # ---- TITLE box (right panel top)
    title = QgsLayoutItemLabel(layout)
    title.setText("Artfynd")
    title.setFont(QFont("Courier New", 18, QFont.Bold))
    title.setHAlign(Qt.AlignLeft)
    title.setVAlign(Qt.AlignVCenter)
    title.attemptMove(QgsLayoutPoint(RIGHT_PANEL_X, 10, QgsUnitTypes.LayoutMillimeters))
    title.attemptResize(QgsLayoutSize(RIGHT_PANEL_W, 18, QgsUnitTypes.LayoutMillimeters))
    title.setFrameEnabled(False)
    title.setBackgroundEnabled(True)
    title.setBackgroundColor(Qt.white)
    layout.addLayoutItem(title)

    # ---- LEGEND (right panel)
    legend = QgsLayoutItemLegend(layout)
    legend.setStyleFont(QgsLegendStyle.Title, QFont("Courier New", 10, QFont.Bold))
    legend.setStyleFont(QgsLegendStyle.Group, QFont("Courier New", 9))
    legend.setStyleFont(QgsLegendStyle.Subgroup, QFont("Courier New", 9))
    legend.setStyleFont(QgsLegendStyle.SymbolLabel, QFont("Courier New", 9))
    legend.setTitle("Teckenförklaring")
    LEGEND_Y = 32
    legend_item_count = len(cats) + (1 if has_knaerot else 0)
    LEGEND_BASE_H_MM = 16
    LEGEND_ROW_H_MM = 6
    legend_h = max(28, min(120, LEGEND_BASE_H_MM + legend_item_count * LEGEND_ROW_H_MM))
    legend.attemptMove(QgsLayoutPoint(RIGHT_PANEL_X, LEGEND_Y, QgsUnitTypes.LayoutMillimeters))
    legend.attemptResize(QgsLayoutSize(RIGHT_PANEL_W, legend_h, QgsUnitTypes.LayoutMillimeters))
    legend.setFrameEnabled(False)
    legend.setBackgroundEnabled(True)
    legend.setBackgroundColor(Qt.white)
    legend.setLinkedMap(map_item)
    legend.setAutoUpdateModel(False)

    # keep only: knaerot dissolved + points_copy (categorized)
    model = legend.model()
    rootg = model.rootGroup()

    # remove everything first
    for child in list(rootg.children()):
        rootg.removeChildNode(child)

    # add exactly the layers you want in the legend (order matters)
    # show dissolved protection zone only if Knärot exists
    if has_knaerot:
        rootg.addLayer(knaerot_diss)
    rootg.addLayer(points)

    legend.updateLegend()
    layout.addLayoutItem(legend)

    # ---- SCALE BAR (bottom-left)
    scalebar = QgsLayoutItemScaleBar(layout)
    scalebar.setLinkedMap(map_item)
    scalebar.applyDefaultSize()
    scalebar.setStyle("Single Box")
    scalebar.setUnits(QgsUnitTypes.DistanceMeters)
    scalebar.setNumberOfSegments(2)
    scalebar.setNumberOfSegmentsLeft(0)
    scalebar.setUnitsPerSegment(50)   # 0–50–100 m (adjust as you prefer)
    scalebar.attemptMove(QgsLayoutPoint(20, 188, QgsUnitTypes.LayoutMillimeters))
    scalebar.attemptResize(QgsLayoutSize(60, 8, QgsUnitTypes.LayoutMillimeters))
    layout.addLayoutItem(scalebar)

    # ---- NORTH ARROW (robust)
    north = QgsLayoutItemPicture(layout)
    north.setPicturePath(
        ":/images/north_arrows/layout_default_north_arrow.svg",
        QgsLayoutItemPicture.FormatSVG
    )
    north.attemptMove(QgsLayoutPoint(10, 180, QgsUnitTypes.LayoutMillimeters))
    north.attemptResize(QgsLayoutSize(12, 12, QgsUnitTypes.LayoutMillimeters))
    north.setLinkedMap(map_item)
    layout.addLayoutItem(north)

    # --- Build unique rows from the points layer (Rödlistningsklass) ---
    rows = set()
    for f in points.getFeatures():
        artnr = f[ARTNR_FIELD]
        art = f[SPECIES_FIELD]
        rl = f[REDLIST_FIELD]
        if artnr is None:
            continue
        rows.add((int(artnr), str(art or ""), str(rl or "")))

    rows = sorted(rows, key=lambda t: (t[0], t[1]))

    # --- Render as HTML table (preferred). Fallback to plain text if HTML mode not available ---
    html = ["<div style='font-family: Arial; font-size: 9pt;'>",
            "<b>Artlista</b>",
            "<table cellspacing='0' cellpadding='2' style='border-collapse:collapse;'>",
            "<tr><th align='left'>Artnr</th><th align='left'>Artnamn</th><th align='left'>Rödlistning</th></tr>"]
    for artnr, art, rl in rows:
        html.append(f"<tr><td>{artnr}</td><td>{art}</td><td>{rl}</td></tr>")
    html.append("</table></div>")
    html = "\n".join(html)

    plain_lines = ["Artlista",
                   "Artnr  Artnamn                          Rödlista"]
    for artnr, art, rl in rows:
        plain_lines.append(f"{artnr:<5}  {art[:28]:<28}  {rl}")
    plain_text = "\n".join(plain_lines)

    artlista = QgsLayoutItemLabel(layout)
    layout.addLayoutItem(artlista)

    # Position art list directly under legend (dynamic if legend height changes)
    layout.refresh()
    legend_pos = legend.positionWithUnits()
    legend_size = legend.sizeWithUnits()
    ARTLIST_GAP_MM = 3
    effective_legend_h = max(legend_size.height(), legend_h)
    artlista_y = legend_pos.y() + effective_legend_h + ARTLIST_GAP_MM
    artlista_h = max(20, 210 - artlista_y)
    artlista_w = min(RIGHT_PANEL_W + 6, 297 - RIGHT_PANEL_X - 2)
    artlista.attemptMove(QgsLayoutPoint(RIGHT_PANEL_X, artlista_y, QgsUnitTypes.LayoutMillimeters))
    artlista.attemptResize(QgsLayoutSize(artlista_w, artlista_h, QgsUnitTypes.LayoutMillimeters))

    # style: white background, no frame
    artlista.setBackgroundEnabled(True)
    artlista.setBackgroundColor(Qt.white)
    artlista.setFrameEnabled(False)
    artlista.setMargin(1.5)

    # HTML if supported, else plain text
    if hasattr(artlista, "setMode") and hasattr(QgsLayoutItemLabel, "ModeHtml"):
        artlista.setMode(QgsLayoutItemLabel.ModeHtml)
        artlista.setText(html)
    else:
        artlista.setFont(QFont("Courier New", 9))
        artlista.setText(plain_text)

    # final refresh
    layout.refresh()
    iface.openLayoutDesigner(layout)
    print(f"Layout created: {layout_name}")

def optional_dense_labels():
    # Always use dense labels as per settings
    USE_DENSE_LABELS = True

    if USE_DENSE_LABELS:
        # räkna grannar
        R = 10        # meter
        DENSE_MIN = 2 # minst 2 punkter inom R (inkl. sig själv)

        # 0) Skapa stabil nyckel för join (undvik provider-beroende fid)
        points_id_path = p("punkter_med_id")
        delete_if_exists(points_id_path)
        processing.run("native:fieldcalculator", {
            "INPUT": points,
            "FIELD_NAME": "pt_id",
            "FIELD_TYPE": 1,
            "FIELD_LENGTH": 20,
            "FIELD_PRECISION": 0,
            "FORMULA": "$id",
            "OUTPUT": points_id_path
        })
        points_for_density = QgsVectorLayer(points_id_path, "Rödlistningsklass (id)", "ogr")
        require_valid(points_for_density, f"Could not load points with id: {points_id_path}")

        # 1) Skapa buffert runt varje punkt (R meter)
        buf_path = p("punkter_buffer")
        delete_if_exists(buf_path)
        processing.run("native:buffer", {
            "INPUT": points_for_density,
            "DISTANCE": R,
            "SEGMENTS": 8,
            "END_CAP_STYLE": 0,
            "JOIN_STYLE": 0,
            "MITER_LIMIT": 2,
            "DISSOLVE": False,
            "OUTPUT": buf_path
        })
        buf = QgsVectorLayer(buf_path, "punkter_buffer", "ogr")
        require_valid(buf, f"Could not load buffer: {buf_path}")

        # 2) Räkna hur många punkter som faller i varje buffertpolygon
        counted_poly_path = p("buffer_med_antal")
        delete_if_exists(counted_poly_path)
        processing.run("native:countpointsinpolygon", {
            "POLYGONS": buf,
            "POINTS": points_for_density,
            "FIELD": "pt_count",   # nytt fält i polygonlagret
            "WEIGHT": "",
            "CLASSFIELD": "",
            "OUTPUT": counted_poly_path
        })
        buf_counted = QgsVectorLayer(counted_poly_path, "buffer_med_antal", "ogr")
        require_valid(buf_counted, f"Could not load counted buffers: {counted_poly_path}")
        print("buf_counted fields:", [f.name() for f in buf_counted.fields()])

        # 3) Join tillbaka pt_count till points via stabilt pt_id
        points_counted_path = p("punkter_med_grannar")
        delete_if_exists(points_counted_path)
        processing.run("native:joinattributestable", {
            "INPUT": points_for_density,
            "FIELD": "pt_id",
            "INPUT_2": buf_counted,
            "FIELD_2": "pt_id",
            "FIELDS_TO_COPY": ["pt_count"],
            "METHOD": 1,
            "DISCARD_NONMATCHING": False,
            "PREFIX": "",
            "OUTPUT": points_counted_path
        })

        points_counted = QgsVectorLayer(points_counted_path, "Rödlistningsklass", "ogr")
        require_valid(points_counted, f"Could not load points_counted: {points_counted_path}")
       

        print("points_counted fields:", [f.name() for f in points_counted.fields()])
        vals = []
        for i, ft in enumerate(points_counted.getFeatures()):
            vals.append(ft["pt_count"])
            if i >= 10:
                break
        print("First pt_count values:", vals)

        # Dela upp i dense / sparse
        COUNT_FIELD = "pt_count"
        dense_path = p("punkter_dense")
        sparse_path = p("punkter_sparse")
        delete_if_exists(dense_path)
        delete_if_exists(sparse_path)
        
        processing.run("native:extractbyexpression", {
            "INPUT": points_counted,
            "EXPRESSION": f"\"{COUNT_FIELD}\" >= {DENSE_MIN}",
            "OUTPUT": dense_path
        })

        processing.run("native:extractbyexpression", {
            "INPUT": points_counted,
            "EXPRESSION": f"\"{COUNT_FIELD}\" < {DENSE_MIN}",
            "OUTPUT": sparse_path
        })

        dense = QgsVectorLayer(dense_path, "Rödlistningsklass (täta)", "ogr")
        sparse = QgsVectorLayer(sparse_path, "Rödlistningsklass (glesa)", "ogr")
        
        dense_count = dense.featureCount()
        sparse_count = sparse.featureCount()
        total_count = points_counted.featureCount()
        print(f"Täta punkter (>= {DENSE_MIN} inom {R} m): {dense_count}")
        print(f"Glesa punkter (< {DENSE_MIN} inom {R} m): {sparse_count}")
        print(f"Totalt antal punkter: {total_count}")
        
        add_layer(sparse)
        add_layer(dense)

        # 
        points.setLabelsEnabled(False)

        # Sparse: 
        apply_simple_labels(sparse, ARTNR_FIELD, size=8)
        sparse.setRenderer(renderer_template.clone())
        sparse.triggerRepaint()
        # Dense: 
        apply_callout_labels(dense, ARTNR_FIELD, dist_mm=2, size=8)  
        dense.setRenderer(renderer_template.clone())
        dense.triggerRepaint()

def finalize():
    try:
        iface.mapCanvas().refresh()
    except Exception:
        pass

    print("Done. run_id =", run_id)
    print("Raster:", RASTER_PATH)
    print("Excel:", TABLE_PATH, "sheet:", SHEET_NAME)
    print("Points:", points.featureCount())
    print("Knärot points:", knaerot_pts.featureCount())

# ----------------------------
# Main execution
# ----------------------------
def main():
    initialize_globals()
    ask_user_for_files()
    load_raster()
    process_excel_to_table()
    clean_and_create_points()
    apply_symbology()
    create_knaerot_zones()
    # Zoom to all points before creating layout
    iface.mapCanvas().setExtent(points.extent())
    iface.mapCanvas().refresh()
    create_print_layout()
    optional_dense_labels()
    arrange_layers()
    finalize()

if __name__ in ("__main__", "__console__"):
    main()
