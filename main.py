import shutil
import uuid

from fastapi import (
    FastAPI,
    UploadFile,
    File,
    Request,
    Form
)

from fastapi.staticfiles import StaticFiles

from fastapi.responses import (
    FileResponse,
    HTMLResponse
)

from fastapi.templating import (
    Jinja2Templates
)

from config import (
    UPLOAD_DIR,
    OUTPUT_DIR,
    MAPPING_FILE,
    BASE_DIR
)

from kml_parser import KMLParser
from boq_mapper import BOQMapper
from excel_generator import ExcelGenerator


# =====================================
# APP
# =====================================

app = FastAPI(
    title="FTTH BOQ Generator"
)

# Static Folder
app.mount(
    "/static",
    StaticFiles(directory="static"),
    name="static"
)

templates = Jinja2Templates(
    directory="templates_html"
)


# =====================================
# HOME
# =====================================

@app.get(
    "/",
    response_class=HTMLResponse
)
async def home(
    request: Request
):

    return templates.TemplateResponse(
        "index.html",
        {
            "request": request
        }
    )


# =====================================
# GENERATE BOQ
# =====================================

@app.post("/generate")
async def generate_boq(
    request: Request,
    kml_file: UploadFile = File(...),
    network_type: str = Form("perumahan")
):
    unique_id = str(uuid.uuid4())
    upload_path = UPLOAD_DIR / f"{unique_id}_{kml_file.filename}"

    try:
        # =====================================
        # SAVE UPLOAD
        # =====================================
        with open(upload_path, "wb") as buffer:
            shutil.copyfileobj(kml_file.file, buffer)

        # =====================================
        # PARSE KML
        # =====================================
        parser = KMLParser(str(upload_path))
        parsed_data = parser.parse()
        raw_objects = parsed_data["raw_objects"]
        cable_spans = parsed_data["cable_spans"]

        # =====================================
        # TEMPLATE FILE (MASTER)
        # =====================================
        template_file = BASE_DIR / "FORMAT BOQ PLAN.xlsx"
        if not template_file.exists():
            raise FileNotFoundError("Berkas master template 'FORMAT BOQ PLAN.xlsx' tidak ditemukan di root direktori.")
        template_name = "Master"

        # =====================================
        # BOQ MAPPING
        # =====================================
        mapper = BOQMapper(
            MAPPING_FILE,
            template_name,
            network_type=network_type,
            template_path=str(template_file)
        )
        boq_result = mapper.map_objects(raw_objects)

        # =====================================
        # GENERATE EXCEL
        # =====================================
        generator = ExcelGenerator(str(template_file))
        generator.fill_boq(boq_result)
        generator.update_boq_header(kml_file.filename, template_name)
        generator.create_data_kml_sheet(raw_objects, cable_spans, kml_file.filename)

        # =====================================
        # OUTPUT FILE NAME
        # =====================================
        clean_name = (
            kml_file.filename
            .replace(".kml", "")
            .replace(" ", "_")
        )
        output_file = OUTPUT_DIR / f"BOQ_{clean_name}.xlsx"
        generator.save(str(output_file))
        print(f"[INFO] Output : {output_file}")

        # =====================================
        # DOWNLOAD
        # =====================================
        return FileResponse(
            path=str(output_file),
            filename=f"BOQ_{clean_name}.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        # Clean up uploaded file if it exists
        if upload_path.exists():
            try:
                upload_path.unlink()
            except Exception:
                pass
        return templates.TemplateResponse(
            "index.html",
            {
                "request": request,
                "error_message": f"Gagal memproses file KML: {str(e)}"
            }
        )


# =====================================
# HEALTH CHECK
# =====================================

@app.get("/health")
async def health():

    return {
        "status": "ok"
    }