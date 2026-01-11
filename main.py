import os
import time
from typing import List

from fastapi import FastAPI, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pptx import Presentation
from pptx.util import Pt
from pydantic import BaseModel

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"]
)


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
    content: List[str]


class SlidesRequest(BaseModel):
    templateId: str
    title: str
    slides: List[Slide]


@app.post("/api/generate-pptx")
async def generate_pptx(request: SlidesRequest = Body(...)):
    try:
        slides_data = request.dict()
        template_id = slides_data.get("templateId")

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

        filename = f"presentation-{template_id}-{int(time.time())}.pptx"
        os.makedirs("presentations", exist_ok=True)
        output_path = f"presentations/{filename}"
        prs.save(output_path)

        return {"status": "success", "downloadUrl": f"/download/{filename}"}

    except Exception as e:
        return {"status": "error", "message": str(e)}


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
                p.text = content  # Single string for paragraph
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
                p.text = content  # content is a string
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
