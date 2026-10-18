from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class PageTransform:
    crop_x: float = 0.0
    crop_y: float = 0.0
    rotation: int = 0
    scale: float = 1.0
    page_width: float = 0.0
    page_height: float = 0.0
    deskew_degrees: float = 0.0

    def _rotate(self, x: float, y: float) -> tuple[float, float]:
        rotation = self.rotation % 360
        if rotation == 0:
            return x, y
        if rotation == 90:
            return self.page_height - y, x
        if rotation == 180:
            return self.page_width - x, self.page_height - y
        if rotation == 270:
            return y, self.page_width - x
        raise ValueError("rotation must be a multiple of 90; deskew is recorded separately")

    def _derotate(self, x: float, y: float) -> tuple[float, float]:
        rotation = self.rotation % 360
        if rotation == 0:
            return x, y
        if rotation == 90:
            return y, self.page_height - x
        if rotation == 180:
            return self.page_width - x, self.page_height - y
        if rotation == 270:
            return self.page_width - y, x
        raise ValueError("rotation must be a multiple of 90; deskew is recorded separately")

    @staticmethod
    def _bounds(points: tuple[tuple[float, float], ...]) -> tuple[float, float, float, float]:
        return min(p[0] for p in points), min(p[1] for p in points), max(p[0] for p in points), max(p[1] for p in points)

    def to_render(self, box: tuple[float, float, float, float]) -> tuple[float, float, float, float]:
        x0, y0, x1, y1 = box
        rotated = self._bounds(tuple(self._rotate(x, y) for x, y in ((x0, y0), (x1, y0), (x0, y1), (x1, y1))))
        return (
            rotated[0] * self.scale - self.crop_x,
            rotated[1] * self.scale - self.crop_y,
            rotated[2] * self.scale - self.crop_x,
            rotated[3] * self.scale - self.crop_y,
        )

    def to_source(self, box: tuple[float, float, float, float]) -> tuple[float, float, float, float]:
        x0, y0, x1, y1 = box
        rendered = tuple(((x + self.crop_x) / self.scale, (y + self.crop_y) / self.scale) for x, y in ((x0, y0), (x1, y0), (x0, y1), (x1, y1)))
        return self._bounds(tuple(self._derotate(x, y) for x, y in rendered))

    def roundtrip_ok(self, box: tuple[float, float, float, float], tolerance: float = 0.01) -> bool:
        return all(abs(a - b) <= tolerance for a, b in zip(box, self.to_source(self.to_render(box))))
