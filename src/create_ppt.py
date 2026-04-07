# src/create_ppt.py
import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

from pptx import Presentation
import theme
import content
import builders

OUTPUT_PATH = os.path.join(
    os.path.dirname(__file__), "..", "output", "network-card-csi-v1.0.pptx"
)

def create_presentation():
    prs = Presentation()
    prs.slide_width = theme.SLIDE_WIDTH
    prs.slide_height = theme.SLIDE_HEIGHT

    blank_layout = prs.slide_layouts[6]  # completely blank layout

    for idx, slide_data in enumerate(content.SLIDES):
        enriched = {**slide_data, "_slide_num": idx + 1}
        slide = prs.slides.add_slide(blank_layout)
        slide_type = enriched["type"]
        builder_fn = builders.BUILDERS.get(slide_type)
        if builder_fn is None:
            print(f"WARNING: unknown slide type '{slide_type}', skipping")
            continue
        builder_fn(slide, enriched)
        if enriched.get("speaker_notes"):
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = enriched["speaker_notes"]

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    prs.save(OUTPUT_PATH)
    print(f"Saved: {OUTPUT_PATH}")
    print(f"Total slides: {len(prs.slides)}")

if __name__ == "__main__":
    create_presentation()
