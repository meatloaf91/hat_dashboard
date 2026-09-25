"""PDF artwork detection, cropping, and image export.

The Qt UI lives in ``new_ui_window.py``. This module intentionally contains
only the artwork processing pipeline so it can also be tested or reused by a
future batch workflow.
"""

from __future__ import annotations

import re
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

OutputFormat = Literal["tiff", "png", "jpeg"]


@dataclass(frozen=True)
class Bounds:
	left: float
	top: float
	right: float
	bottom: float

	@property
	def width(self) -> float:
		return max(0.0, self.right - self.left)

	@property
	def height(self) -> float:
		return max(0.0, self.bottom - self.top)


@dataclass
class DetectionResult:
	bounds: Bounds
	confidence: float
	source: str
	evidence: list[str] = field(default_factory=list)
	warnings: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class ProcessOptions:
	dpi: int = 300
	output_format: OutputFormat = "tiff"
	include_diecut: bool = True
	white_tolerance: int = 8


@dataclass
class ProcessResult:
	input_path: Path
	diecut_path: Path | None
	clean_path: Path | None
	detection: DetectionResult
	warnings: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class CutOption:
	"""A user-selectable layer or vector-color cut candidate."""

	label: str
	kind: Literal["layer", "spot_color", "vector"]
	bounds: Bounds
	detail: str = ""
	color_rgb: tuple[int, int, int] | None = None
	xref: int | None = None


@dataclass(frozen=True)
class CutSelection:
	bounds: Bounds
	kind: Literal["layer", "spot_color", "manual"] = "manual"
	name: str = ""
	color_rgb: tuple[int, int, int] | None = None
	xref: int | None = None
	hide_kind: Literal["layer", "spot_color", ""] = ""
	hide_name: str = ""
	hide_color_rgb: tuple[int, int, int] | None = None
	hide_xref: int | None = None
	hide_items: tuple[tuple[str, str, tuple[int, int, int] | None, int | None], ...] = ()
	manual_crop: bool = False
	stroke_shape: Literal["rectangle", "ellipse", "custom"] = "rectangle"
	corners: Literal["sharp", "rounded", "beveled"] = "sharp"
	corner_amount: float = 0.0
	custom_path: tuple[tuple[str, tuple[float, ...]], ...] = ()
	stroke_weight: float = 1.0
	stroke_color: str = "#B8F35A"


@dataclass
class ArtworkInspection:
	input_path: Path
	preview: Any
	page_bounds: Bounds
	detected_bounds: Bounds
	options: list[CutOption] = field(default_factory=list)
	warnings: list[str] = field(default_factory=list)


DIECUT_TERMS = ("die", "cut", "dieline", "die-line", "knife", "contour", "cutline")
CAD_TERMS = ("cad", "technical", "construction", "crease", "fold", "score", "bleed")


def _require_dependencies() -> tuple[Any, Any]:
	try:
		try:
			import pymupdf as fitz
		except ImportError:
			import fitz
		from PIL import Image
	except ImportError as error:
		raise RuntimeError(
			"Install artwork dependencies with: python -m pip install PyMuPDF Pillow"
		) from error
	return fitz, Image


def _bounds_from_rect(rect: Any) -> Bounds:
	return Bounds(float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1))


def _document_layers(document: Any) -> list[dict[str, Any]]:
	"""Return PDF optional-content layers across PyMuPDF layer APIs."""
	layers = document.get_layers() or []
	if layers:
		return list(layers)
	ocgs = document.get_ocgs() or {}
	return [{"xref": xref, **dict(info)} for xref, info in ocgs.items()]


def _color_label(color: Any) -> str | None:
	if not isinstance(color, (tuple, list)) or len(color) < 3:
		return None
	values = tuple(max(0, min(255, round(float(value) * 255))) for value in color[:3])
	return "RGB #{:02X}{:02X}{:02X}".format(*values)


def _cmyk_to_rgb(cyan: float, magenta: float, yellow: float, black: float) -> tuple[int, int, int]:
	return tuple(round(255 * (1 - min(1.0, component / 100 + black / 100))) for component in (cyan, magenta, yellow))


def _pdf_spot_colors(document: Any) -> list[tuple[str, tuple[int, int, int]]]:
	"""Read named spot plates and CMYK swatches from Illustrator XMP metadata."""
	xml = document.get_xml_metadata() or ""
	spots: list[tuple[str, tuple[int, int, int]]] = []
	for match in re.finditer(
		r"<xmpG:swatchName>(.*?)</xmpG:swatchName>(.*?)</rdf:li>", xml, re.S
	):
		name, block = match.groups()
		if re.search(r"<xmpG:type>SPOT</xmpG:type>", block, re.I):
			values = []
			for tag in ("cyan", "magenta", "yellow", "black"):
				value = re.search(rf"<xmpG:{tag}>([-0-9.]+)</xmpG:{tag}>", block)
				values.append(float(value.group(1)) if value else 0.0)
			spots.append((name.strip(), _cmyk_to_rgb(*values)))
	return spots


def _distance(first: tuple[int, int, int], second: tuple[int, int, int]) -> int:
	return sum((a - b) ** 2 for a, b in zip(first, second))


def inspect_artwork(input_path: Path, dpi: int = 72, preview_width: int = 720) -> ArtworkInspection:
	"""Render the first page and return named-layer/vector-color cut choices."""
	fitz, _ = _require_dependencies()
	if dpi < 36:
		raise ValueError("Preview resolution must be at least 36 DPI.")
	document = fitz.open(input_path)
	try:
		if document.page_count == 0:
			raise ValueError(f"PDF has no pages: {input_path.name}")
		page = document[0]
		page_bounds = _bounds_from_rect(page.rect)
		detected_bounds = detect_diecut(page).bounds
		options: list[CutOption] = []
		seen: set[tuple[str, str]] = set()

		layer_drawings: dict[str, list[Any]] = {}
		for drawing in page.get_drawings():
			layer_name = str(drawing.get("layer") or "").strip().lower()
			if layer_name:
				layer_drawings.setdefault(layer_name, []).append(drawing)

		for layer in _document_layers(document):
			name = str(layer.get("name", "")).strip()
			if not name:
				continue
			layer_bounds = [
				_bounds_from_rect(drawing["rect"])
				for drawing in layer_drawings.get(name.lower(), [])
				if drawing.get("rect")
			]
			if layer_bounds:
				bounds = Bounds(
					min(item.left for item in layer_bounds), min(item.top for item in layer_bounds),
					max(item.right for item in layer_bounds), max(item.bottom for item in layer_bounds),
				)
			else:
				bounds = page_bounds
			key = ("layer", name.lower())
			if key not in seen:
				seen.add(key)
				options.append(CutOption(name, "layer", bounds, "PDF layer", xref=layer.get("xref")))

		drawings = [drawing for drawing in page.get_drawings() if drawing.get("rect")]
		named_spots = _pdf_spot_colors(document)
		if named_spots:
			for name, spot_rgb in named_spots:
				matching = []
				for drawing in drawings:
					color = drawing.get("color") or drawing.get("fill")
					if isinstance(color, (tuple, list)) and len(color) >= 3:
						rgb = tuple(round(float(value) * 255) for value in color[:3])
						matching.append((_distance(spot_rgb, rgb), drawing))
				best = [drawing for distance, drawing in sorted(matching, key=lambda item: item[0])[:8] if distance < 9000]
				if best:
					bounds = Bounds(
						min(float(d["rect"].x0) for d in best), min(float(d["rect"].y0) for d in best),
						max(float(d["rect"].x1) for d in best), max(float(d["rect"].y1) for d in best),
					)
				else:
					bounds = page_bounds
				options.append(CutOption(name, "spot_color", bounds, "PDF spot color", spot_rgb))

		for drawing in drawings:
			rect = drawing.get("rect")
			if not rect or rect.width <= 0 or rect.height <= 0:
				continue
			color = _color_label(drawing.get("color") or drawing.get("fill"))
			if named_spots:
				color = None
			if color:
				key = ("color", color)
				if key not in seen:
					seen.add(key)
					options.append(CutOption(color, "spot_color", _bounds_from_rect(rect), "Vector stroke/fill", tuple(max(0, min(255, round(float(value) * 255))) for value in (drawing.get("color") or drawing.get("fill"))[:3])))
			width = float(drawing.get("width") or 0)
			if width <= 1.5:
				options.append(CutOption("Thin vector rectangle", "vector", _bounds_from_rect(rect), "Thin vector geometry"))

		preview_scale = min(1.0, preview_width / max(page.rect.width, 1))
		preview = page.get_pixmap(matrix=fitz.Matrix(preview_scale, preview_scale), alpha=False, colorspace=fitz.csRGB)
		warnings = []
		if not options:
			warnings.append("No named die-cut layer or vector color was found; page bounds are the only cut area.")
		return ArtworkInspection(input_path, preview, page_bounds, detected_bounds, options, warnings)
	finally:
		document.close()


def _name_score(name: str) -> tuple[int, bool]:
	normalized = re.sub(r"[^a-z0-9]+", " ", name.lower())
	has_diecut = any(term in normalized for term in DIECUT_TERMS)
	has_cad = any(term in normalized for term in CAD_TERMS)
	return (3 if has_diecut else 0) - (2 if has_cad else 0), has_diecut


def detect_diecut(page: Any) -> DetectionResult:
	"""Estimate a trim rectangle from PDF layers and vector geometry."""
	page_rect = page.rect
	candidates: list[tuple[float, Bounds, str, str]] = []

	for drawing in page.get_drawings():
		rect = drawing.get("rect")
		if not rect or rect.width <= 0 or rect.height <= 0:
			continue
		width = float(drawing.get("width") or 0)
		score = 0.25
		evidence = "vector rectangle"
		if drawing.get("color") is not None or drawing.get("fill") is not None:
			score += 0.1
			evidence += ", named color unavailable"
		if width <= 1.5:
			score += 0.1
			evidence += ", thin stroke"
		area_ratio = (rect.width * rect.height) / (page_rect.width * page_rect.height)
		if 0.2 <= area_ratio <= 0.98:
			score += 0.25
		candidates.append((score, _bounds_from_rect(rect), "vector geometry", evidence))

	layer_info = getattr(page.parent, "get_layers", lambda: [])()
	for layer in _document_layers(page.parent):
		name = str(layer.get("name", ""))
		name_score, is_diecut = _name_score(name)
		if is_diecut:
			score = min(0.98, 0.6 + max(0, name_score) * 0.1)
			candidates.append((score, _bounds_from_rect(page_rect), f"layer: {name}", "die-cut layer name"))

	if not candidates:
		return DetectionResult(
			_bounds_from_rect(page_rect), 0.0, "page bounds fallback",
			warnings=["No die-cut vector or named layer was detected; page bounds were used."],
		)

	score, bounds, source, evidence = max(candidates, key=lambda item: item[0])
	warnings: list[str] = []
	if score < 0.7:
		warnings.append("Die-cut detection is uncertain; review the crop before delivery.")
	if bounds.width >= page_rect.width * 0.98 and bounds.height >= page_rect.height * 0.98:
		warnings.append("Detected bounds are almost the full page; the PDF trim box may be wrong.")
	return DetectionResult(bounds, round(min(score, 0.99), 2), source, [evidence], warnings)


def _crop_pixmap(page: Any, bounds: Bounds, dpi: int) -> Any:
	fitz, _ = _require_dependencies()
	scale = dpi / 72.0
	clip = fitz.Rect(bounds.left, bounds.top, bounds.right, bounds.bottom)
	return page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip, alpha=False, colorspace=fitz.csRGB)


def _normalize_white(image: Any, tolerance: int) -> Any:
	from PIL import Image
	from PIL import ImageChops

	white = Image.new("RGB", image.size, (255, 255, 255))
	difference = ImageChops.difference(image, white)
	mask = difference.convert("L").point(lambda value: 255 if value <= tolerance else 0)
	image.paste(white, mask=mask)
	return image


def _save_pixmap(pixmap: Any, target: Path, output_format: OutputFormat, white_tolerance: int) -> None:
	_, Image = _require_dependencies()
	if hasattr(pixmap, "samples"):
		mode = "RGBA" if pixmap.alpha else "RGB"
		image = Image.frombytes(mode, (pixmap.width, pixmap.height), pixmap.samples)
	else:
		image = pixmap
	if image.mode == "RGB":
		image = _normalize_white(image, white_tolerance)
	target.parent.mkdir(parents=True, exist_ok=True)
	if output_format == "jpeg":
		image.save(target, format="JPEG", quality=95, dpi=(300, 300))
	elif output_format == "png":
		image.save(target, format="PNG", dpi=(300, 300))
	else:
		image.save(target, format="TIFF", compression="tiff_lzw", dpi=(300, 300))


def _remove_color_from_pixmap(pixmap: Any, color_rgb: tuple[int, int, int], tolerance: int = 24) -> Any:
	"""Erase a selected separation color from a rendered clean copy."""
	_, Image = _require_dependencies()
	if hasattr(pixmap, "samples"):
		mode = "RGBA" if pixmap.alpha else "RGB"
		image = Image.frombytes(mode, (pixmap.width, pixmap.height), pixmap.samples)
	else:
		image = pixmap.copy()
	pixels = image.load()
	red, green, blue = color_rgb
	for y in range(image.height):
		for x in range(image.width):
			current = pixels[x, y]
			if max(abs(current[0] - red), abs(current[1] - green), abs(current[2] - blue)) <= tolerance:
				pixels[x, y] = (current[0], current[1], current[2], 0) if image.mode == "RGBA" else (255, 255, 255)
	return image


class ArtworkProcessor:
	def __init__(self, options: ProcessOptions | None = None) -> None:
		self.options = options or ProcessOptions()

	def process_file(
		self,
		input_path: Path,
		output_dir: Path,
		cut_bounds: Bounds | CutSelection | None = None,
	) -> ProcessResult:
		fitz, _ = _require_dependencies()
		if self.options.dpi < 72:
			raise ValueError("Resolution must be at least 72 DPI.")
		output_dir.mkdir(parents=True, exist_ok=True)
		document = fitz.open(input_path)
		try:
			if document.page_count == 0:
				raise ValueError(f"PDF has no pages: {input_path.name}")
			page = document[0]
			detection = detect_diecut(page)
			selection = cut_bounds if isinstance(cut_bounds, CutSelection) else None
			if selection is not None:
				cut_bounds = selection.bounds
			if selection is not None and selection.kind == "manual" and not selection.manual_crop:
				cut_bounds = _bounds_from_rect(page.rect)
			if cut_bounds is not None:
				detection = DetectionResult(
					cut_bounds, 1.0, "manual selection", ["User-selected manual cut"], []
				)
			if selection and selection.kind == "layer":
				layer_xrefs = [
					layer.get("xref") for layer in _document_layers(document)
					if layer.get("xref") and (
						selection.xref == layer.get("xref")
						or str(layer.get("name", "")) == selection.name
					)
				]
				if layer_xrefs:
					document.set_layer(-1, on=layer_xrefs)
			extension = self.options.output_format
			stem = input_path.stem
			diecut_path = output_dir / f"{stem}_with_diecut.{extension}" if self.options.include_diecut else None
			clean_path = output_dir / f"{stem}_clean.{extension}"
			if diecut_path:
				_save_pixmap(_crop_pixmap(page, detection.bounds, self.options.dpi), diecut_path, extension, self.options.white_tolerance)
			clean_pixmap, clean_warning = self._render_clean(document, page, detection, selection)
			warnings = list(detection.warnings)
			if clean_warning:
				warnings.append(clean_warning)
			_save_pixmap(clean_pixmap, clean_path, extension, self.options.white_tolerance)
			return ProcessResult(input_path, diecut_path, clean_path, detection, warnings)
		finally:
			document.close()

	def _render_clean(
		self,
		document: Any,
		page: Any,
		detection: DetectionResult,
		selection: CutSelection | None = None,
	) -> tuple[Any, str | None]:
		fitz, _ = _require_dependencies()
		layers = _document_layers(document)
		hide_kind = selection.hide_kind if selection else ""
		hide_name = selection.hide_name if selection else ""
		hide_xref = selection.hide_xref if selection else None
		hide_color_rgb = selection.hide_color_rgb if selection else None
		hide_items = list(getattr(selection, "hide_items", ()) or [])
		if selection and hide_kind and hide_name:
			hide_items.append((hide_kind, hide_name, hide_color_rgb, hide_xref))
		layer_targets = [
			layer.get("xref") for layer in layers
			if layer.get("xref") and (
				(selection and selection.kind == "layer" and str(layer.get("name", "")) == selection.name)
				or (selection and hide_kind == "layer" and (hide_xref == layer.get("xref") or str(layer.get("name", "")) == hide_name))
				or any(
					(item_kind == "layer" and (item_xref == layer.get("xref") or str(layer.get("name", "")) == item_name))
					for item_kind, item_name, _, item_xref in hide_items
				)
				or (selection is None and any(term in str(layer.get("name", "")).lower() for term in DIECUT_TERMS + CAD_TERMS))
			)
		]
		base_pixmap = _crop_pixmap(page, detection.bounds, self.options.dpi)
		pixmap = base_pixmap
		if layer_targets:
			try:
				document.set_layer(-1, basestate="ON", off=layer_targets)
				with tempfile.TemporaryDirectory(prefix="hat_artwork_") as temp_dir:
					temp_path = Path(temp_dir) / "clean_render.pdf"
					document.save(temp_path)
					clean_document = fitz.open(temp_path)
					try:
						clean_page = clean_document[0]
						scale = self.options.dpi / 72.0
						clip = fitz.Rect(detection.bounds.left, detection.bounds.top, detection.bounds.right, detection.bounds.bottom)
						pixmap = clean_page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip, alpha=True, colorspace=fitz.csRGB)
					finally:
						clean_document.close()
			except (TypeError, RuntimeError, ValueError, OSError):
				pass
		warning = "Clean export uses the cropped render because this PDF is flattened or has no controllable OCG layers."
		for item_kind, item_name, item_color, item_xref in hide_items:
			if item_kind == "spot_color" and item_color:
				pixmap = _remove_color_from_pixmap(pixmap, item_color)
				warning = None
		if selection and hide_kind == "spot_color" and hide_color_rgb:
			pixmap = _remove_color_from_pixmap(pixmap, hide_color_rgb)
			warning = None
		if selection and hide_kind == "spot_color" and not hide_color_rgb:
			warning = "The selected spot color could not be masked because its color metadata was not available."
		return pixmap, warning

	def process_batch(self, input_dir: Path, output_dir: Path) -> list[ProcessResult]:
		paths = sorted(path for path in input_dir.glob("*.pdf") if path.is_file())
		if not paths:
			raise FileNotFoundError(f"No PDF files found in {input_dir}")
		return [self.process_file(path, output_dir) for path in paths]


__all__ = [
		"ArtworkInspection", "ArtworkProcessor", "Bounds", "CutOption", "CutSelection",
	"CutSelection", "DetectionResult", "ProcessOptions", "ProcessResult", "detect_diecut",
	"inspect_artwork",
]
