import os
import time
from typing import List

import redis
from fastapi import FastAPI, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pptx import Presentation
from fastapi import Request
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt, Inches
import hashlib
from pydantic import BaseModel
from starlette.responses import JSONResponse

FREE_LIMIT = 3

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"]
)


def get_redis_client() -> redis.Redis | None:
    url = os.getenv("REDIS_URL")
    if not url:
        print("REDIS_URL not set, guest limiting disabled")
        return None
    try:
        client = redis.from_url(url)
        # Light ping to validate connection
        client.ping()
        print("Connected to Redis successfully")
        return client
    except Exception as e:
        print(f"Redis connection failed: {e}")
        return None


# Redis connection
r = get_redis_client()


def create_guest_id(request: Request) -> str:
    """Fingerprint = IP + User-Agent (privacy-safe)"""
    ip = request.client.host.replace("::ffff:", "")
    ua = request.headers.get("user-agent", "")[:100]
    fingerprint = f"{ip}:{ua}"
    return hashlib.md5(fingerprint.encode()).hexdigest()


@app.middleware("http")
async def guest_limiter(request: Request, call_next):
    # If Redis not available, skip limiting (but app still works)
    if r is None:
        return await call_next(request)

    if request.url.path == "/api/generate-pptx" and request.method == "POST":
        guest_id = create_guest_id(request)
        key = f"guest:{guest_id}"

        # Track usage
        count = r.incr(key)
        r.expire(key, 2592000)  # 30 days

        # Store in request state for PPT generation
        request.state.guest_count = count
        request.state.guest_id = guest_id

        if count > FREE_LIMIT:
            return JSONResponse(
                status_code=402,
                content={
                    "status": "limit_reached",
                    "used": count,
                    "limit": FREE_LIMIT,
                    "guest_id": guest_id,
                    "upgrade_url": "/upgrade",
                    "message": "You've created 3 amazing presentations!"
                }
            )

    response = await call_next(request)
    return response


@app.get("/api/templates")
async def get_templates():
    """
    Fetch available templates grouped by type
    """
    return {
        "status": "success",
        "templates": [
            {
                "id": "geometric",
                "name": "geometric",
                "description": "",
                "thumbnails": ["/thumbnails/thumbnail_geometric_1.png", "/thumbnails/thumbnail_geometric_2.png"]
            },
            {
                "id": "streamline",
                "name": "Streamline",
                "description": "",
                "thumbnails": ["/thumbnails/thumbnail_streamline_1.png", "/thumbnails/thumbnail_streamline_2.png"]
            },
            {
                "id": "swiss",
                "name": "swiss",
                "description": "",
                "thumbnails": ["/thumbnails/thumbnail_swiss_1.png", "/thumbnails/thumbnail_swiss_2.png"]
            },
            {
                "id": "momentum",
                "name": "momentum",
                "description": "",
                "thumbnails": ["/thumbnails/thumbnail_momentum_1.png", "/thumbnails/thumbnail_momentum_2.png"]
            },
            {
                "id": "material",
                "name": "material",
                "description": "",
                "thumbnails": ["/thumbnails/thumbnail_material_1.png", "/thumbnails/thumbnail_material_2.png"]
            },
            {
                "id": "slate",
                "name": "slate",
                "description": "",
                "thumbnails": ["/thumbnails/thumbnail_slate_1.png", "/thumbnails/thumbnail_slate_2.png"]
            },
            {
                "id": "paperback",
                "name": "paperback",
                "description": "",
                "thumbnails": ["/thumbnails/thumbnail_paperback_1.png", "/thumbnails/thumbnail_paperback_2.png"]
            }
        ]
    }


class Slide(BaseModel):
    slideNumber: int
    title: str
    type: str
    format: str
    content: List[str]


class SlidesRequest(BaseModel):
    templateId: str
    title: str
    slides: List[Slide]


@app.post("/api/generate-pptx")
async def generate_pptx(request: SlidesRequest = Body(...), request_obj: Request = None):
    try:
        # Get guest info from middleware (if exists)
        guest_count = getattr(request_obj.state, 'guest_count', 0) if request_obj else 0
        is_guest = not request_obj.headers.get("authorization") if request_obj else True

        slides_data = request.dict()
        template_id = slides_data.get("templateId")

        print(f" - Guest count: {guest_count}")

        template_files = {
            "geometric": "templates/geometric.pptx",
            "streamline": "templates/streamline.pptx",
            "swiss": "templates/swiss.pptx",
            "momentum": "templates/momentum.pptx",
            "material": "templates/material.pptx",
            "slate": "templates/slate.pptx",
            "paperback": "templates/paperback.pptx"
        }
        template_path = template_files.get(template_id, template_files["slate"])

        prs = Presentation(template_path)

        # Get template slides
        title_template = prs.slides[0]  # First slide = title template
        content_template = prs.slides[1]  # Second slide = content template

        for slide_data in slides_data["slides"]:
            slide_type = slide_data.get("type", "content")

            # Choose correct template slide
            if slide_type == "title":
                template = title_template
            else:
                template = content_template

            # Use template's layout (SAFE)
            slide = prs.slides.add_slide(template.slide_layout)
            add_slide_content(slide, slide_data)

        r_id_list = prs.slides._sldIdLst

        # Delete slide at index 1 first (to avoid index shift)
        del r_id_list[1]

        # Then delete slide at index 0
        del r_id_list[0]

        # Add watermark for free users (PPT #3+)
        if is_guest and guest_count >= 3:
            add_watermark(prs, "Made with Voice-to-PPT Free")

        filename = f"presentation-{template_id}-{int(time.time())}.pptx"
        os.makedirs("presentations", exist_ok=True)
        output_path = f"presentations/{filename}"
        prs.save(output_path)

        return {
            "status": "success",
            "downloadUrl": f"/download/{filename}",
            "is_guest": is_guest,
            "guest_count": guest_count,
            "show_upgrade": is_guest and guest_count >= 3,
            "watermark": is_guest and guest_count >= 3
        }

    except Exception as e:
        return {"status": "error", "message": str(e)}


def add_watermark(prs, text: str):
    """Add subtle watermark to all slides for free users"""
    for slide in prs.slides:
        left = Inches(0.5)
        top = Inches(6.5)
        width = Inches(7)
        height = Inches(0.5)

        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(12)
        p.font.color.rgb = RGBColor(0xF5, 0xA6, 0x23)  # Gold color
        p.alignment = PP_ALIGN.CENTER


def get_correct_layout(prs, slide_type):
    """Get correct layout for slide type"""

    # Title slide - always first
    if slide_type == "title":
        return prs.slide_layouts[0]  # Title Slide

    # Content slide - look for layouts with title + content placeholders
    for layout in prs.slide_layouts[1:]:  # Skip title
        # Check for title (placeholder 0) + content (placeholder 1)
        has_title = False
        has_content = False

        for shape in layout.placeholders:
            if shape.placeholder_format.idx == 0:  # Title placeholder
                has_title = True
            elif shape.placeholder_format.idx == 1:  # Content placeholder
                has_content = True

        if has_title and has_content:
            return layout

    # Fallback: first layout with content
    for layout in prs.slide_layouts:
        if any(shape.placeholder_format.idx == 1 for shape in layout.placeholders):
            return layout

    return prs.slide_layouts[1]  # Final fallback


def add_slide_content(slide, slide_data):
    """Update text in placeholders - NEVER clear text_frame"""

    slide_type = slide_data.get("type", "content")
    title = slide_data.get("title", "")
    content_format = slide_data.get("format", "bullets")  # NEW: Get format type
    content = slide_data.get("content", [])

    print(f" - Content: {content_format}")

    # Get ALL text shapes in order
    text_shapes = [shape for shape in slide.shapes if shape.has_text_frame]
    text_shapes.sort(key=lambda s: s.top)

    if slide_type == "title":
        # Title slide
        if len(text_shapes) >= 1:
            # Update first shape (title)
            p = text_shapes[0].text_frame.paragraphs[0]
            p.text = title

        if len(text_shapes) >= 2 and content:
            # Update second shape (subtitle)
            p = text_shapes[1].text_frame.paragraphs[0]
            if content_format == "paragraph":
                p.text = content[0]  # Single string for paragraph
            else:
                p.text = content[0]  # First item for bullets

    else:  # content slide
        # Content slide
        if len(text_shapes) >= 1:
            # Update first shape (title)
            p = text_shapes[0].text_frame.paragraphs[0]
            p.text = title

        if len(text_shapes) >= 2:
            # Update second shape (content)
            content_frame = text_shapes[1].text_frame

            # IMPORTANT: Don't clear! Just update/add paragraphs
            # Remove ALL paragraphs except first
            while len(content_frame.paragraphs) > 1:
                # Get the paragraph element
                p_element = content_frame.paragraphs[1].element
                # Remove from XML
                p_element.getparent().remove(p_element)

            # NEW: Handle based on format type
            if content_format == "paragraph":
                # Single paragraph content
                p = content_frame.paragraphs[0]
                p.text = content[0]  # content is a string
                p.level = 0
            else:  # bullets format
                # Multiple bullet points (content is array)
                for i, bullet in enumerate(content):
                    if i == 0:
                        # Update first paragraph (already exists)
                        p = content_frame.paragraphs[0]
                        p.text = bullet
                    else:
                        # Add new paragraph
                        p = content_frame.add_paragraph()
                        p.text = bullet

                    p.level = 0


def debug_slide(slide):
    """Print all shapes in slide"""
    for i, shape in enumerate(slide.shapes):
        print(f"Shape {i}: {shape.name}")
        if shape.has_text_frame:
            print(f" - Has text: {shape.text_frame.text[:50]}")
        else:
            print(f" - No text (image/design)")


@app.get("/debug")
async def debug():
    import os
    return {
        "cwd": os.getcwd(),
        "thumbnails_exists": os.path.exists("thumbnails"),
        "thumbnails_contents": os.listdir("thumbnails") if os.path.exists("thumbnails") else "NO FOLDER",
        "mount_status": "Check if app.mount line exists above"
    }


@app.get("/download/{filename}")
async def download_presentation(filename: str):
    return FileResponse(
        path=f"presentations/{filename}",
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )


app.mount("/thumbnails", StaticFiles(directory="thumbnails"), name="thumbnails")

if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)  # ← host="0.0.0.0"
