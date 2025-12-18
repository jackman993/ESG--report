"""
Component for arranging SDG icon images in a grid on a PPT slide.
"""
from pptx.util import Cm
from pathlib import Path


class SDGIconGrid:
    def __init__(self, prs):
        self.prs = prs

    def add_to_slide(
        self,
        slide,
        icon_paths,
        left_cm: float = 3.0,
        top_cm: float = 3.0,
        cell_size_cm: float = 3.5,
        gap_cm: float = 0.4,
        columns: int = 3,
    ):
        if not icon_paths:
            return
        rows = (len(icon_paths) + columns - 1) // columns
        for idx, icon in enumerate(icon_paths):
            if not Path(icon).exists():
                continue
            col = idx % columns
            row = idx // columns
            left = Cm(left_cm + col * (cell_size_cm + gap_cm))
            top = Cm(top_cm + row * (cell_size_cm + gap_cm))
            slide.shapes.add_picture(str(icon), left, top, width=Cm(cell_size_cm))


