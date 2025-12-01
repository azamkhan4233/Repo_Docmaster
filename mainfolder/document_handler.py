# ==============================
# document_handler.py
# ==============================

import re
import os
import tempfile
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P


class DocumentHandler:
    """
    Central parser that converts DOCX or plain text into a structured model:

    sections = {
        "Introduction": [
            {"type": "text", "content": "..."},
            {"type": "image", "path": "...", "caption": "...", "alt_text": "...", "virtual": False},
            {"type": "table", "data": [["..."], ["..."]]}
        ],
        ...
    }

    detected_order = ["Uncategorized", "Abstract", "Introduction", ...]
    """

    # Recognized headings for section detection
    SECTION_PATTERNS = [
        r'^\s*(abstract)\s*$',
        r'^\s*(introduction)\s*$',
        r'^\s*(literature\s+review)\s*$',
        r'^\s*(background)\s*$',
        r'^\s*(methodology|materials\s+and\s+methods|methods)\s*$',
        r'^\s*(experiment|experimental\s+setup)\s*$',
        r'^\s*(results|findings|analysis)\s*$',
        r'^\s*(discussion|observations)\s*$',
        r'^\s*(conclusion|summary|closing\s+remarks)\s*$',
        r'^\s*(future\s+work|recommendations)\s*$',
        r'^\s*(references|bibliography|works\s+cited)\s*$'
    ]

    # Figure caption pattern: "Figure 1: ...", "Fig. 2 - ..."
    FIGURE_CAPTION_PATTERN = re.compile(
        r'^(fig(ure)?\s*\d+[:.\-)]\s+).*', re.IGNORECASE
    )

    def __init__(self):
        self.sections = {}
        self.detected_order = []
        self._last_text = None  # last paragraph text, used for caption detection

    # ====================================================
    # PLAIN TEXT / PDF TEXT PARSER  (no tables/images)
    # ====================================================
    def parse_text(self, text: str):
        """
        Parse raw text into sections based on heading patterns.
        Used for TXT input and PDF-extracted text.
        """
        self.sections = {}
        self.detected_order = []

        current_section = "Uncategorized"
        self.sections[current_section] = []
        self.detected_order.append(current_section)

        buffer = []

        def flush():
            nonlocal buffer
            if buffer:
                para = " ".join(buffer).strip()
                if para:
                    self.sections[current_section].append({
                        "type": "text",
                        "content": para
                    })
                buffer = []

        for line in text.splitlines():
            stripped = line.strip()
            if not stripped:
                flush()
                continue

            heading = self._detect_heading(stripped)
            if heading:
                flush()
                current_section = heading
                if heading not in self.sections:
                    self.sections[heading] = []
                    self.detected_order.append(heading)
            else:
                buffer.append(stripped)

        flush()
        self._cleanup()
        return self.sections

    # ====================================================
    # DOCX PARSER (text + tables + inline images + captions)
    # ====================================================
    def parse_docx(self, path: str):
        """
        Parse a DOCX file into sections, preserving order of text, tables, images.
        """
        doc = Document(path)
        self.sections = {"Uncategorized": []}
        self.detected_order = ["Uncategorized"]
        current_section = "Uncategorized"
        self._last_text = None

        # Image extraction directory
        img_dir = os.path.join(tempfile.gettempdir(), "docmaster_images")
        os.makedirs(img_dir, exist_ok=True)

        # Map relationship id -> image file path
        image_map = self._extract_all_images(doc, img_dir)

        # Iterate in XML body order for correct sequence
        for block in doc.element.body:

            # Paragraph
            if isinstance(block, CT_P):
                para = Paragraph(block, doc)
                txt = para.text.strip()

                # Heading detection
                heading = self._detect_heading(txt) if txt else None
                if heading:
                    current_section = heading
                    if heading not in self.sections:
                        self.sections[heading] = []
                        self.detected_order.append(heading)
                    self._last_text = txt
                    continue

                # Inline images within this paragraph
                imgs = self._extract_images_from_paragraph(para, image_map)
                if imgs:
                    for img in imgs:
                        img["caption"] = self._detect_caption(self._last_text, txt)
                        self.sections[current_section].append(img)
                    self._last_text = txt
                    continue

                # Normal text paragraph
                if txt:
                    self.sections[current_section].append({
                        "type": "text",
                        "content": txt
                    })
                self._last_text = txt

            # Table
            elif isinstance(block, CT_Tbl):
                tbl = Table(block, doc)
                rows = []
                for r in tbl.rows:
                    rows.append([c.text.strip() for c in r.cells])
                self.sections[current_section].append({
                    "type": "table",
                    "data": rows
                })

        self._cleanup()
        return self.sections

    # ====================================================
    # IMAGE HANDLING HELPERS
    # ====================================================
    def _extract_all_images(self, doc, image_dir):
        """
        Save all embedded images to disk and return rId -> file_path mapping.
        """
        image_map = {}
        for rel in list(doc.part.rels.values()):
            try:
                if "image" in rel.reltype:
                    part = rel.target_part
                    filename = os.path.basename(part.partname)
                    out_path = os.path.join(image_dir, filename)
                    with open(out_path, "wb") as f:
                        f.write(part.blob)
                    image_map[rel.rId] = out_path
            except Exception:
                continue
        return image_map

    def _extract_images_from_paragraph(self, para, image_map):
        """
        Find images inside runs (inline pictures) for a paragraph.
        """
        results = []
        for run in para.runs:
            nodes = run._r.xpath(".//*")
            for node in nodes:
                for attr in node.attrib:
                    if "embed" in attr:
                        rId = node.attrib[attr]
                        results.append({
                            "type": "image",
                            "path": image_map.get(rId),
                            "caption": None,
                            "alt_text": node.get("descr") or "",
                            "virtual": False
                        })
        return results

    def _detect_caption(self, previous, current):
        """
        Caption is either just before or just after an image if it matches
        'Figure 1: ...' or 'Fig. 2 - ...'.
        """
        if previous and self.FIGURE_CAPTION_PATTERN.match(previous):
            return previous
        if current and self.FIGURE_CAPTION_PATTERN.match(current):
            return current
        return None

    # ====================================================
    # HEADING & CLEANUP HELPERS
    # ====================================================
    def _detect_heading(self, text: str):
        """
        Identify section headings like 'Introduction' or '1. Introduction'.
        """
        if not text:
            return None
        l = text.lower().strip()

        # Handle numbered headings like "1. Introduction"
        numbered = re.match(r'^\s*\d+[\.\)]\s*(.+)$', l)
        candidate = numbered.group(1).strip() if numbered else l

        for patt in self.SECTION_PATTERNS:
            if re.match(patt, candidate):
                # Strip numbers and return cleaned original text
                return re.sub(r'^\s*\d+[\.\)]\s*', '', text).strip().title()
        return None

    def _cleanup(self):
        """
        Drop empty sections and sync detected_order.
        """
        self.sections = {k: v for k, v in self.sections.items() if v}
        self.detected_order = [s for s in self.detected_order if s in self.sections]

    # ====================================================
    # VIRTUAL GRAPH/CHART DETECTION (for PDF text)
    # ====================================================
    def detect_virtual_image(self, paragraph_text):
        """
        Detect keywords that imply presence of a graph/chart/figure
        that cannot be extracted as an actual image (PDF flattened, etc.).
        Returns a virtual image element if matched, else None.
        """
        if not paragraph_text:
            return None

        text_lower = paragraph_text.lower()
        keywords = ["figure", "fig.", "graph", "chart", "plot", "diagram"]

        if any(k in text_lower for k in keywords):
            return {
                "type": "image",
                "path": None,
                "caption": paragraph_text,
                "virtual": True,
                "alt_text": ""
            }
        return None

    # For debugging / console inspection
    def get_structured_text(self):
        out = []
        for sec in self.detected_order:
            out.append(f"=== {sec} ===")
            for el in self.sections[sec]:
                if el["type"] == "text":
                    out.append(el["content"])
                elif el["type"] == "table":
                    out.append("[TABLE]")
                    for row in el["data"]:
                        out.append(" | ".join(row))
                    out.append("[/TABLE]")
                elif el["type"] == "image":
                    out.append(
                        f"[IMAGE] path={el.get('path')} VIRTUAL={el.get('virtual')} "
                        f"CAPTION={el.get('caption')} ALT={el.get('alt_text')}"
                    )
            out.append("")
        return "\n".join(out)


# ==============================
# document_handler.py  (AI + Rules + Captions)
# ==============================

# import re
# import os
# import tempfile
# from docx import Document
# from docx.table import Table
# from docx.text.paragraph import Paragraph
# from docx.oxml.table import CT_Tbl
# from docx.oxml.text.paragraph import CT_P

# # === AI Hybrid Classifier Import ===
# from section_predictor import predict_section


# class DocumentHandler:
#     """
#     Central parser that converts DOCX or plain text into a structured model:

#     sections = {
#         "Introduction": [
#             {"type": "text", "content": "..."},
#             {"type": "image", "path": "...", "caption": "...", "alt_text": "...", "virtual": False},
#             {"type": "table", "data": [["..."], ["..."]]}
#         ],
#         ...
#     }

#     detected_order = ["Uncategorized", "Abstract", "Introduction", ...]
#     """

#     # Regex-based headings (manual detection first priority)
#     SECTION_PATTERNS = [
#         r'^\s*(abstract)\s*$',
#         r'^\s*(introduction)\s*$',
#         r'^\s*(literature\s+review)\s*$',
#         r'^\s*(background)\s*$',
#         r'^\s*(methodology|materials\s+and\s+methods|methods)\s*$',
#         r'^\s*(experiment|experimental\s+setup)\s*$',
#         r'^\s*(results|findings|analysis)\s*$',
#         r'^\s*(discussion|observations)\s*$',
#         r'^\s*(conclusion|summary|closing\s+remarks)\s*$',
#         r'^\s*(future\s+work|recommendations)\s*$',
#         r'^\s*(references|bibliography|works\s+cited)\s*$'
#     ]

#     # Captions like "Figure 1: ..."
#     FIGURE_CAPTION_PATTERN = re.compile(r'^(fig(ure)?\s*\d+[:.\-)]+.*)', re.IGNORECASE)

#     def __init__(self):
#         self.sections = {}
#         self.detected_order = []
#         self._last_text = None  # track previous paragraph for caption detection

#     # ====================================================
#     # PLAIN TEXT/PDF INPUT (NO TABLE/IMAGE support here)
#     # ====================================================
#     def parse_text(self, text: str):
#         self.sections = {}
#         self.detected_order = []

#         current_section = "Uncategorized"
#         self.sections[current_section] = []
#         self.detected_order.append(current_section)

#         buffer = []

#         def flush():
#             nonlocal buffer, current_section
#             if buffer:
#                 para = " ".join(buffer).strip()
#                 if para:
#                     # Try AI prediction when no explicit heading
#                     predicted = predict_section(para)
#                     current_section = predicted.title()
#                     if current_section not in self.sections:
#                         self.sections[current_section] = []
#                         self.detected_order.append(current_section)

#                     self.sections[current_section].append({"type": "text", "content": para})
#                 buffer = []

#         for line in text.splitlines():
#             stripped = line.strip()
#             if not stripped:
#                 flush()
#                 continue

#             # Manual regex heading wins first
#             heading = self._detect_heading(stripped)
#             if heading:
#                 flush()
#                 current_section = heading
#                 if heading not in self.sections:
#                     self.sections[heading] = []
#                     self.detected_order.append(heading)
#             else:
#                 buffer.append(stripped)

#         flush()
#         self._cleanup()
#         return self.sections

#     # ====================================================
#     # DOCX PARSER (AI + manual headings + tables + images)
#     # ====================================================
#     def parse_docx(self, path: str):
#         doc = Document(path)
#         self.sections = {"Uncategorized": []}
#         self.detected_order = ["Uncategorized"]
#         current_section = "Uncategorized"
#         self._last_text = None

#         # Save all embedded images
#         img_dir = os.path.join(tempfile.gettempdir(), "docmaster_images")
#         os.makedirs(img_dir, exist_ok=True)
#         image_map = self._extract_all_images(doc, img_dir)

#         # Parse document in exact body order
#         for block in doc.element.body:

#             # ---------- Text/Paragraph ----------
#             if isinstance(block, CT_P):
#                 para = Paragraph(block, doc)
#                 txt = para.text.strip()

#                 # Heading detection
#                 heading = self._detect_heading(txt) if txt else None
#                 if heading:
#                     current_section = heading
#                     if heading not in self.sections:
#                         self.sections[heading] = []
#                         self.detected_order.append(heading)
#                     self._last_text = txt
#                     continue

#                 # Inline images
#                 imgs = self._extract_images_from_paragraph(para, image_map)
#                 if imgs:
#                     for img in imgs:
#                         # Custom chosen caption format: "[IMAGE] – caption text"
#                         caption = self._detect_caption(self._last_text, txt)
#                         if caption:
#                             img["caption"] = f"[IMAGE] – {caption}"
#                         self.sections[current_section].append(img)
#                     self._last_text = txt
#                     continue

#                 # Normal text → ask AI to classify section if no heading before
#                 if txt:
#                     # If last section was auto or no heading, detect via model
#                     predicted = predict_section(txt)
#                     current_section = predicted.title()

#                     if current_section not in self.sections:
#                         self.sections[current_section] = []
#                         self.detected_order.append(current_section)

#                     self.sections[current_section].append({"type": "text", "content": txt})

#                 self._last_text = txt

#             # ---------- Table ----------
#             elif isinstance(block, CT_Tbl):
#                 tbl = Table(block, doc)
#                 rows = [[c.text.strip() for c in r.cells] for r in tbl.rows]
#                 self.sections[current_section].append({"type": "table", "data": rows})

#         self._cleanup()
#         return self.sections

#     # ====================================================
#     # IMAGE HELPERS
#     # ====================================================
#     def _extract_all_images(self, doc, image_dir):
#         image_map = {}
#         for rel in list(doc.part.rels.values()):
#             try:
#                 if "image" in rel.reltype:
#                     filename = os.path.basename(rel.target_part.partname)
#                     out_path = os.path.join(image_dir, filename)
#                     with open(out_path, "wb") as f:
#                         f.write(rel.target_part.blob)
#                     image_map[rel.rId] = out_path
#             except Exception:
#                 continue
#         return image_map

#     def _extract_images_from_paragraph(self, para, image_map):
#         results = []
#         for run in para.runs:
#             nodes = run._r.xpath(".//*")
#             for node in nodes:
#                 for attr in node.attrib:
#                     if "embed" in attr:
#                         rId = node.attrib[attr]
#                         results.append({"type": "image", "path": image_map.get(rId), "caption": None,
#                                         "alt_text": node.get("descr") or "", "virtual": False})
#         return results

#     def _detect_caption(self, previous, current):
#         if previous and self.FIGURE_CAPTION_PATTERN.match(previous):
#             return previous
#         if current and self.FIGURE_CAPTION_PATTERN.match(current):
#             return current
#         return None

#     # ====================================================
#     # HEADING CLEANUP
#     # ====================================================
#     def _detect_heading(self, text: str):
#         if not text:
#             return None
#         l = text.lower().strip()
#         numbered = re.match(r'^\s*\d+[\.\)]\s*(.+)$', l)
#         candidate = numbered.group(1).strip() if numbered else l
#         for patt in self.SECTION_PATTERNS:
#             if re.match(patt, candidate):
#                 return re.sub(r'^\s*\d+[\.\)]\s*', '', text).strip().title()
#         return None

#     def _cleanup(self):
#         self.sections = {k: v for k, v in self.sections.items() if v}
#         self.detected_order = [s for s in self.detected_order if s in self.sections]

#     # Debug
#     def get_structured_text(self):
#         out = []
#         for sec in self.detected_order:
#             out.append(f"=== {sec} ===")
#             for el in self.sections[sec]:
#                 out.append(str(el))
#             out.append("")
#         return "\n".join(out)
