# tests/test_ppt.py
import sys
import os
import pytest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from pptx import Presentation
import theme

PPT_PATH = os.path.join(os.path.dirname(__file__), '..', 'output', 'network-card-csi.pptx')

@pytest.fixture(scope="module")
def prs():
    assert os.path.exists(PPT_PATH), f"PPT not found: {PPT_PATH}"
    return Presentation(PPT_PATH)

def test_slide_count(prs):
    assert len(prs.slides) == 25

def test_slide_dimensions(prs):
    assert prs.slide_width == theme.SLIDE_WIDTH
    assert prs.slide_height == theme.SLIDE_HEIGHT

def test_slide_17_has_table(prs):
    slide = prs.slides[16]  # 0-indexed, slide 17
    tables = [s for s in slide.shapes if s.has_table]
    assert len(tables) >= 1, "Slide 17 should have a table"

def test_slide_21_has_table(prs):
    slide = prs.slides[20]
    tables = [s for s in slide.shapes if s.has_table]
    assert len(tables) >= 1, "Slide 21 should have a table"

def test_slide_22_has_table(prs):
    slide = prs.slides[21]
    tables = [s for s in slide.shapes if s.has_table]
    assert len(tables) >= 1, "Slide 22 should have a table"

def test_all_slides_have_shapes(prs):
    for i, slide in enumerate(prs.slides):
        assert len(slide.shapes) > 0, f"Slide {i+1} has no shapes"
