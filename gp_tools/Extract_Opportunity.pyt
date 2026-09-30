# -*- coding: utf-8 -*-
"""Extract_Opportunity.pyt

ArcGIS Pro Python Toolbox.

Purpose
-------
Extract the geographic overlap between an uploaded polygon layer and a
directory of raster files (the ``suitable`` rasters). All polygon features are
rasterized once into a single zone raster, then each ``suitable`` raster is
summarized in one pass with ZonalStatisticsAsTable. The per-zone COUNT is the
number of suitable pixels within each feature, converted to square kilometers,
and written into a copy of a local Excel (Gantt) template -- one output
workbook per polygon feature.

Performance
-----------
Zonal statistics computes all features at once, so the tool runs one raster
pass per suitable raster (N rasters) rather than N x M (rasters x features)
ExtractByMask operations. The zone raster is built once and snapped to the
value rasters (no resampling), and the Excel template layout is scanned once
and reused for every per-feature output copy.

Output (simplified, template-aligned)
-------------------------------------
- The tool reads the FIRST worksheet of the template to learn the ordered
  list of codes in column E (row 9 down).
- For each feature it writes a plain NEW workbook whose column E / column J
  values sit in the SAME row positions as the template (row 9 down), so the
  user can copy column J and paste it straight into column J of the template.
- The template's drawings/charts are never touched (editing them via openpyxl
  corrupts the Gantt shapes), so a simplified copy is produced instead.

Output naming
-------------
- One copy of the template is produced per polygon feature.
- Each copy is named ``<template_name>_<feature_name>.xlsx`` where the feature
  name comes from a user-chosen field on the input polygon features.

Assumptions (first shot)
------------------------
1. "Suitable" rasters are single-band. Cells that are part of the suitable
   area are cells with a value (i.e. not NoData). The "intersect value" for a
   raster is the count of suitable pixels that fall inside the polygon.
2. The rasters are in an Equal Earth (equal-area) projection where each pixel
   represents exactly 1 square kilometer. Therefore:
       suitable_area_km2 = suitable_pixel_count * 1
3. The template's column E holds an id that matches the raster file name
   (matched with and without file extension).
"""

import os
import re

import arcpy

# Each pixel is 1 square kilometer (rasters are in Equal Earth / equal-area).
KM2_PER_PIXEL = 1.0

# Fixed template layout.
ID_COLUMN = "E"          # column holding the raster id to match against
RESULT_COLUMN = "J"      # column that receives the km2 value
FIRST_DATA_ROW = 9       # first row of data in the template

# Raster file names look like ``AGR_01_suitable.tif``; the id used to match
# against column E is the leading ``<LETTERS>_<DIGITS>`` code (e.g. ``AGR_01``).
CODE_PATTERN = re.compile(r"^([A-Za-z]+_\d+)")


def raster_code(raster_name):
    """Extract the matching id/code from a raster file name.

    ``AGR_01_suitable.tif`` -> ``AGR_01``. Falls back to the extension-less
    stem if the expected pattern is not present.
    """
    match = CODE_PATTERN.match(str(raster_name).strip())
    if match:
        return match.group(1)
    return os.path.splitext(str(raster_name))[0]


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name is the .pyt file name)."""
        self.label = "Extract Opportunity"
        self.alias = "extractopportunity"
        self.tools = [ExtractOpportunity]


class ExtractOpportunity(object):
    def __init__(self):
        """Define the tool."""
        self.label = "Extract Opportunity"
        self.description = (
            "Extracts the intersecting suitable area between an uploaded "
            "polygon layer and a directory of 'suitable' raster files. Loops "
            "over polygon features with a search cursor, intersects each "
            "polygon against every raster in the directory, counts suitable "
            "pixels (1 pixel = 1 sq km, Equal Earth), and writes the km2 "
            "values into column J (starting row 9) of a copy of the Gantt "
            "template, matching rasters by id in column E. One output "
            "workbook is produced per polygon feature, named "
            "<template>_<feature name>."
        )
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define the tool parameters."""

        # 0 - Input polygon feature layer / class (the uploaded polygon).
        in_polygons = arcpy.Parameter(
            displayName="Input Polygon Features",
            name="in_polygons",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input",
        )
        in_polygons.filter.list = ["Polygon"]

        # 1 - Field on the polygons whose value names each output workbook.
        feature_name_field = arcpy.Parameter(
            displayName="Feature Name Field",
            name="feature_name_field",
            datatype="Field",
            parameterType="Required",
            direction="Input",
        )
        feature_name_field.parameterDependencies = [in_polygons.name]

        # 2 - Directory containing the 'suitable' rasters.
        suitable_dir = arcpy.Parameter(
            displayName="Suitable Raster Directory",
            name="suitable_dir",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="Input",
        )

        # 3 - Local Excel (Gantt) template file to copy and populate.
        template_file = arcpy.Parameter(
            displayName="Excel Gantt Template File",
            name="template_file",
            datatype="DEFile",
            parameterType="Required",
            direction="Input",
        )
        template_file.filter.list = ["xlsx"]

        # 4 - Output folder for the populated per-feature workbooks.
        out_folder = arcpy.Parameter(
            displayName="Output Folder",
            name="out_folder",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input",
        )

        # 5 - Raster file extensions to include.
        raster_extensions = arcpy.Parameter(
            displayName="Raster File Extensions",
            name="raster_extensions",
            datatype="GPString",
            parameterType="Optional",
            direction="Input",
            multiValue=True,
        )
        raster_extensions.filter.type = "ValueList"
        raster_extensions.filter.list = ["tif", "tiff", "img", "gdb"]
        raster_extensions.values = ["tif", "tiff", "img"]

        return [
            in_polygons,
            feature_name_field,
            suitable_dir,
            template_file,
            out_folder,
            raster_extensions,
        ]

    def isLicensed(self):
        """Spatial Analyst is required for raster extraction."""
        try:
            if arcpy.CheckExtension("Spatial") != "Available":
                return False
        except Exception:
            return False
        return True

    def updateParameters(self, parameters):
        return

    def updateMessages(self, parameters):
        return

    # ------------------------------------------------------------------ #
    # Helpers
    # ------------------------------------------------------------------ #
    @staticmethod
    def _list_suitable_rasters(suitable_dir, extensions):
        """Return a list of (name, full_path) for rasters in the directory.

        Uses an arcpy workspace listing so both file-based rasters (.tif,
        .img) and rasters inside a file geodataset are handled consistently.
        """
        rasters = []
        prev_ws = arcpy.env.workspace
        try:
            arcpy.env.workspace = suitable_dir
            for name in arcpy.ListRasters() or []:
                full = os.path.join(suitable_dir, name)
                # If extensions were supplied, filter file-based rasters by ext.
                ext = os.path.splitext(name)[1].lstrip(".").lower()
                if extensions and ext and ext not in extensions:
                    continue
                rasters.append((name, full))
        finally:
            arcpy.env.workspace = prev_ws
        return rasters

    @staticmethod
    def _safe_filename(value):
        """Make a value safe to use as part of a file name."""
        text = str(value).strip() if value is not None else ""
        if not text:
            text = "unnamed"
        # Replace characters that are invalid in Windows file names.
        return re.sub(r'[<>:"/\\|?*]+', "_", text)

    # ------------------------------------------------------------------ #
    # Execute
    # ------------------------------------------------------------------ #
    def execute(self, parameters, messages):
        in_polygons = parameters[0].valueAsText
        feature_name_field = parameters[1].valueAsText
        suitable_dir = parameters[2].valueAsText
        template_file = parameters[3].valueAsText
        out_folder = parameters[4].valueAsText
        raster_extensions = parameters[5].values or []
        raster_extensions = [str(e).lstrip(".").lower() for e in raster_extensions]

        arcpy.CheckOutExtension("Spatial")
        # Local import so the toolbox still loads if SA is unavailable.
        from arcpy.sa import ZonalStatisticsAsTable

        try:
            self._run(
                in_polygons,
                feature_name_field,
                suitable_dir,
                template_file,
                out_folder,
                raster_extensions,
                ZonalStatisticsAsTable,
                messages,
            )
        finally:
            arcpy.CheckInExtension("Spatial")

    def _run(
        self,
        in_polygons,
        feature_name_field,
        suitable_dir,
        template_file,
        out_folder,
        raster_extensions,
        ZonalStatisticsAsTable,
        messages,
    ):
        arcpy.env.overwriteOutput = True

        if not os.path.isfile(template_file):
            raise arcpy.ExecuteError(
                "Excel template file not found: {}".format(template_file)
            )
        if not os.path.isdir(out_folder):
            raise arcpy.ExecuteError(
                "Output folder not found: {}".format(out_folder)
            )

        rasters = self._list_suitable_rasters(suitable_dir, raster_extensions)
        if not rasters:
            raise arcpy.ExecuteError(
                "No rasters found in the suitable directory: {}".format(suitable_dir)
            )
        messages.addMessage("Found {} raster(s) in suitable directory.".format(len(rasters)))

        # Verify there is at least one feature before doing heavy work.
        if int(arcpy.management.GetCount(in_polygons)[0]) == 0:
            raise arcpy.ExecuteError("Input polygon layer has no features.")

        # feature_names (zone id -> feature name) is populated by
        # _build_zone_raster so its keys are exactly the zone ids stored in the
        # zone raster (avoids OID-renaming mismatches after the copy).
        feature_names = {}

        scratch = arcpy.env.scratchGDB
        zone_raster = None
        zone_field = "ZONE_OID"
        prev_snap = arcpy.env.snapRaster
        prev_cell = arcpy.env.cellSize
        prev_parallel = arcpy.env.parallelProcessingFactor

        try:
            # Let Spatial Analyst use multiple CPU cores for the 117 zonal
            # passes. Scoped here (restored in the finally) rather than set at
            # module load so importing the toolbox doesn't mutate global env.
            arcpy.env.parallelProcessingFactor = "66%"

            # ---- Build ONE zone raster from all features (zone = OID). ----
            # Snap and match cell size to the value rasters so counts align
            # 1:1 with the suitable pixels (no resampling).
            first_raster_path = rasters[0][1]
            arcpy.env.snapRaster = first_raster_path
            arcpy.env.cellSize = first_raster_path

            zone_raster, feature_names = self._build_zone_raster(
                in_polygons,
                zone_field,
                feature_name_field,
                first_raster_path,
                scratch,
                messages,
            )
            messages.addMessage(
                "Processing {} polygon feature(s).".format(len(feature_names))
            )

            # In the zone RASTER the OID values live in the standard "Value"
            # field (ZONE_OID was only the source feature field used to build
            # the raster). Zonal stats keys on the raster's Value field.
            zone_value_field = "Value"

            # ---- results[raster_code][oid] = suitable_area_km2 ----
            # One ZonalStatisticsAsTable pass per raster covers ALL features.
            results = {}
            zonal_tbl = os.path.join(scratch, "zs_tbl")
            for raster_index, (raster_name, raster_path) in enumerate(rasters, start=1):
                code = raster_code(raster_name)
                messages.addMessage(
                    "[{}/{}] {} ({})".format(raster_index, len(rasters), raster_name, code)
                )

                per_oid = {}
                try:
                    if arcpy.Exists(zonal_tbl):
                        arcpy.management.Delete(zonal_tbl)
                    # COUNT = number of value-raster (suitable) cells per zone.
                    ZonalStatisticsAsTable(
                        zone_raster,
                        zone_value_field,
                        raster_path,
                        zonal_tbl,
                        "DATA",
                        "SUM",  # requesting SUM also yields COUNT in the table
                    )
                    with arcpy.da.SearchCursor(zonal_tbl, [zone_value_field, "COUNT"]) as cur:
                        for zone_oid, count in cur:
                            per_oid[int(zone_oid)] = int(count or 0)
                except arcpy.ExecuteError as exc:
                    # Typically "no overlap" -> every zone is 0 for this raster.
                    messages.addWarningMessage(
                        "    Zonal stats failed for {}: {}".format(raster_name, exc)
                    )

                # 1 pixel == 1 sq km (Equal Earth, equal-area projection).
                results[code] = {
                    oid: per_oid.get(oid, 0) * KM2_PER_PIXEL
                    for oid in feature_names
                }

            if arcpy.Exists(zonal_tbl):
                arcpy.management.Delete(zonal_tbl)
        finally:
            arcpy.env.snapRaster = prev_snap
            arcpy.env.cellSize = prev_cell
            arcpy.env.parallelProcessingFactor = prev_parallel
            if zone_raster and arcpy.Exists(zone_raster):
                try:
                    arcpy.management.Delete(zone_raster)
                except Exception:
                    pass

        # ---- Write one output workbook per feature. ----
        # Cache the template id->row layout once; it is identical for every
        # copy, so we scan it a single time instead of per feature.
        template_stem = os.path.splitext(os.path.basename(template_file))[0]
        layout = self._read_template_layout(template_file)
        messages.addMessage(
            "Template has {} matchable ids in column {}.".format(
                len(layout["row_for_id"]), ID_COLUMN
            )
        )

        outputs = []
        for oid, feature_name in feature_names.items():
            safe_name = self._safe_filename(feature_name)
            # code -> km2 for this feature.
            feature_results = {code: vals.get(oid, 0) for code, vals in results.items()}
            out_path = os.path.join(
                out_folder, "{}_{}.xlsx".format(template_stem, safe_name)
            )
            self._populate_template(template_file, out_path, feature_results, layout, messages)
            outputs.append(out_path)
            messages.addMessage("Wrote {}".format(out_path))

        # Report the produced workbooks as a derived output (semicolon list).
        arcpy.SetParameterAsText(4, ";".join(outputs))
        messages.addMessage("Produced {} workbook(s).".format(len(outputs)))

    @staticmethod
    def _build_zone_raster(
        in_polygons, zone_field, name_field, value_raster, scratch, messages
    ):
        """Rasterize all polygons into a single zone raster (zone = OID).

        ``name_field`` is the field that names each output workbook. The zones
        are projected into the value raster's coordinate system and rasterized
        at the value raster's (numeric) cell size, so the zone grid aligns 1:1
        with the suitable pixels.

        The zone id is the copy's OID, resolved from the copied dataset rather
        than the input -- copying into a GDB renames the OID (e.g. shapefile
        FID -> OBJECTID), so the input's OID field name is not reused.

        Returns a tuple ``(zone_raster_path, feature_names)`` where
        ``feature_names`` maps the integer zone id (stored in the raster) to
        the feature name. The caller deletes the raster.
        """
        # Describe the value raster to get an explicit numeric cell size and
        # spatial reference. PolygonToRaster's cellsize argument must be a
        # number (passing a raster path raises ERROR 000353).
        rdesc = arcpy.Describe(value_raster)
        cell_size = float(rdesc.meanCellWidth)
        raster_sr = rdesc.spatialReference

        # Copy the OID into a plain long field so PolygonToRaster can use it
        # as the value field (the OID field itself is not always usable).
        zones_fc = os.path.join(scratch, "zones_fc")
        if arcpy.Exists(zones_fc):
            arcpy.management.Delete(zones_fc)

        # Project the polygons into the raster's coordinate system when they
        # differ, so the cell size (raster units) is meaningful.
        poly_sr = arcpy.Describe(in_polygons).spatialReference
        same_sr = (
            raster_sr is not None
            and poly_sr is not None
            and raster_sr.factoryCode != 0
            and poly_sr.factoryCode == raster_sr.factoryCode
        )
        if same_sr or raster_sr is None:
            arcpy.management.CopyFeatures(in_polygons, zones_fc)
        else:
            messages.addMessage(
                "Projecting features from '{}' to raster CS '{}'.".format(
                    getattr(poly_sr, "name", "Unknown"),
                    getattr(raster_sr, "name", "Unknown"),
                )
            )
            arcpy.management.Project(in_polygons, zones_fc, raster_sr)

        # Use the COPY's own OID field name (it may differ from the input's,
        # e.g. shapefile FID -> GDB OBJECTID after CopyFeatures/Project).
        calc_source = arcpy.Describe(zones_fc).OIDFieldName

        if zone_field not in [f.name for f in arcpy.ListFields(zones_fc)]:
            arcpy.management.AddField(zones_fc, zone_field, "LONG")
        arcpy.management.CalculateField(
            zones_fc, zone_field, "!{}!".format(calc_source), "PYTHON3"
        )

        # Read the zone id -> name map from the COPY, keyed on ZONE_OID so it
        # matches exactly the integer values that end up in the zone raster.
        feature_names = {}
        with arcpy.da.SearchCursor(zones_fc, [zone_field, name_field]) as cur:
            for zone_id, name_value in cur:
                if zone_id is not None:
                    feature_names[int(zone_id)] = name_value

        zone_raster = os.path.join(scratch, "zone_raster")
        if arcpy.Exists(zone_raster):
            arcpy.management.Delete(zone_raster)
        messages.addMessage(
            "Building zone raster from features (cell size {}).".format(cell_size)
        )
        arcpy.conversion.PolygonToRaster(
            zones_fc,
            zone_field,
            zone_raster,
            "CELL_CENTER",
            "NONE",
            cell_size,  # explicit numeric cell size (raster units)
        )
        try:
            arcpy.management.Delete(zones_fc)
        except Exception:
            pass
        return zone_raster, feature_names

    @staticmethod
    def _read_template_layout(template_file):
        """Scan the template ONCE and return its reusable layout.

        Returns a dict with:
          - ``row_for_id``: {column-E id (str) -> row index}
          - ``ordered_ids``: list of column-E ids in template row order
            (row 9 down), so a simplified output can mirror the template rows.
          - ``result_col``: integer index of the result column (J)
          - ``max_row``: the last scanned row index in the template.

        The layout is identical for every per-feature copy, so scanning it a
        single time avoids re-reading the workbook structure per feature.
        """
        try:
            import openpyxl
            from openpyxl.utils import column_index_from_string
        except ImportError:
            raise arcpy.ExecuteError(
                "openpyxl is required to populate the Excel template but is "
                "not available in this environment."
            )

        id_col = column_index_from_string(ID_COLUMN)
        result_col = column_index_from_string(RESULT_COLUMN)

        # read_only is fast and low-memory for the layout scan.
        wb = openpyxl.load_workbook(template_file, read_only=True)
        try:
            ws = wb.worksheets[0]
            max_row = ws.max_row
            row_for_id = {}
            # ordered_ids[i] is the column-E id at row FIRST_DATA_ROW + i, or
            # None for a blank row -- preserving exact template positions.
            ordered_ids = []
            for r in range(FIRST_DATA_ROW, max_row + 1):
                key = ws.cell(row=r, column=id_col).value
                if key is None:
                    ordered_ids.append(None)
                    continue
                key = str(key).strip()
                row_for_id[key] = r
                ordered_ids.append(key)
        finally:
            wb.close()

        return {
            "row_for_id": row_for_id,
            "ordered_ids": ordered_ids,
            "result_col": result_col,
            "max_row": max_row,
        }

    @staticmethod
    def _populate_template(template_file, out_path, feature_results, layout, messages):
        """Write a SIMPLIFIED results workbook that mirrors the template rows.

        Rather than editing the Gantt template directly (openpyxl's save cycle
        corrupts the template's drawings), this produces a plain new workbook
        whose column E / column J values line up row-for-row with the template
        starting at row 9. The end user can select column J in this output and
        paste it straight into column J of the real template.

        ``feature_results`` maps raster code -> km2 for a single feature.
        ``layout`` is the cached output of ``_read_template_layout``.
        """
        try:
            import openpyxl
            from openpyxl.utils import column_index_from_string
        except ImportError:
            raise arcpy.ExecuteError(
                "openpyxl is required to write the results workbook but is "
                "not available in this environment."
            )

        result_col = layout["result_col"]
        ordered_ids = layout["ordered_ids"]
        id_col = column_index_from_string(ID_COLUMN)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "results"

        # A small header (above the data start) noting how to use the file.
        ws.cell(row=1, column=id_col, value="Code (col E)")
        ws.cell(row=1, column=result_col, value="Suitable km2 (paste into col J)")

        # Emit one row per template row, in the SAME positions (row 9 down),
        # so column J here aligns 1:1 with column J in the template.
        matched = 0
        unmatched = []
        seen_codes = set()
        for offset, code in enumerate(ordered_ids):
            out_row = FIRST_DATA_ROW + offset
            if code is None:
                continue  # keep blank rows blank to preserve alignment
            seen_codes.add(code)
            ws.cell(row=out_row, column=id_col, value=code)
            if code in feature_results:
                ws.cell(row=out_row, column=result_col, value=feature_results[code])
                matched += 1

        # Any computed codes that are NOT present in the template are reported
        # so nothing is silently dropped.
        for code in feature_results:
            if code not in seen_codes:
                unmatched.append(code)

        if unmatched:
            messages.addWarningMessage(
                "    {} computed code(s) not found in template column {}: {}".format(
                    len(unmatched), ID_COLUMN, ", ".join(sorted(unmatched))
                )
            )

        wb.save(out_path)
