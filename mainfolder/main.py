# ==============================
# main.py
# ==============================

import sys
import os
import tempfile
import docx2txt
import pdfplumber

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QTabWidget, QPushButton, QTextEdit, QLabel,
    QFileDialog, QHBoxLayout, QComboBox, QSpinBox, QMessageBox, QListWidget,
    QDoubleSpinBox, QLineEdit, QGroupBox
)

from document_handler import DocumentHandler
from export_service import ExportService, STYLE_PRESETS


class DocMasterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DocMaster Professional - Research Formatter")
        self.resize(1050, 780)

        # Core state
        self.doc_handler = DocumentHandler()
        self.selected_style = "APA"
        self.custom_style = {}
        self.author = ""
        self.institution = ""

        # Layout + tabs
        layout = QVBoxLayout()
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        self._init_upload_tab()
        self._init_structured_preview_tab()
        self._init_style_tab()
        self._init_formatted_preview_tab()
        self._init_export_tab()

        self.setLayout(layout)

    # ====================================================
    # TAB 1: Upload & Detect
    # ====================================================
    def _init_upload_tab(self):
        tab = QWidget()
        v = QVBoxLayout()

        self.text_input = QTextEdit()
        self.text_input.setPlaceholderText("Paste research text, or import .txt / .docx / .pdf ...")
        v.addWidget(self.text_input)

        btn_txt = QPushButton("Import .txt")
        btn_txt.clicked.connect(self.import_txt)

        btn_docx = QPushButton("Import .docx")
        btn_docx.clicked.connect(self.import_docx)

        btn_pdf = QPushButton("Import .pdf")
        btn_pdf.clicked.connect(self.import_pdf)

        btn_detect = QPushButton("Detect Sections (from text)")
        btn_detect.clicked.connect(self.detect_sections_from_text)

        h = QHBoxLayout()
        h.addWidget(btn_txt)
        h.addWidget(btn_docx)
        h.addWidget(btn_pdf)
        h.addWidget(btn_detect)

        v.addLayout(h)
        tab.setLayout(v)
        self.tabs.addTab(tab, "Upload & Detect")

    def import_txt(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select TXT File", "", "Text Files (*.txt)")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                text = f.read()
            self.text_input.setText(text)
            self.doc_handler = DocumentHandler()
            self.doc_handler.parse_text(text)
            self.refresh_structured_preview()
            QMessageBox.information(self, "TXT Imported", "Text file imported and parsed.")
        except Exception as e:
            QMessageBox.critical(self, "TXT Import Error", str(e))

    def import_docx(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select DOCX File", "", "Word Document (*.docx)")
        if not path:
            return
        try:
            self.doc_handler = DocumentHandler()
            self.doc_handler.parse_docx(path)
            self.text_input.clear()
            self.refresh_structured_preview()
            QMessageBox.information(self, "DOCX Imported", "DOCX imported with tables & images preserved.")
        except Exception as e:
            QMessageBox.critical(self, "DOCX Import Error", str(e))

    def import_pdf(self):
        """
        Smart PDF import (Option B + T1):

        - Extract text per page (clean basic headers/footers)
        - Detect a dominant section heading per page
        - Extract tables where possible
        - Extract images where possible
        - Attach tables/images to nearest section
        - Parse text into sections with DocumentHandler.parse_text
        - Also add virtual image placeholders for caption-only 'Figure/Graph' lines
        """
        path, _ = QFileDialog.getOpenFileName(self, "Select PDF File", "", "PDF Files (*.pdf)")
        if not path:
            return

        try:
            page_texts = []
            page_sections = []
            tables_with_section = []
            images_with_section = []

            img_dir = os.path.join(tempfile.gettempdir(), "docmaster_pdf_images")
            os.makedirs(img_dir, exist_ok=True)

            temp_handler = DocumentHandler()
            last_section = "Uncategorized"

            with pdfplumber.open(path) as pdf:
                for page_index, page in enumerate(pdf.pages):
                    txt = page.extract_text()
                    cleaned_lines = []
                    page_section = None

                    if txt:
                        for line in txt.split("\n"):
                            stripped = line.strip()
                            if not stripped:
                                continue

                            # Heading detection
                            heading = temp_handler._detect_heading(stripped)
                            if heading:
                                page_section = heading

                            # Virtual image detection (caption-only line)
                            virt = temp_handler.detect_virtual_image(stripped)
                            if virt:
                                images_with_section.append((last_section if page_section is None else page_section, virt))

                            cleaned_lines.append(stripped)

                    cleaned_text = "\n".join(cleaned_lines).strip() if cleaned_lines else ""
                    page_texts.append(cleaned_text)

                    # Determine section for this page
                    if page_section:
                        last_section = page_section
                    page_sections.append(last_section)

                    # Extract tables
                    try:
                        tables = page.extract_tables()
                        if tables:
                            for t in tables:
                                tables_with_section.append((last_section, t))
                    except Exception:
                        pass

                    # Extract images (real ones) from this page
                    try:
                        if page.images:
                            page_image = None
                            for im in page.images:
                                try:
                                    if page_image is None:
                                        page_image = page.to_image(resolution=150)
                                    bbox = (im["x0"], im["top"], im["x1"], im["bottom"])
                                    cropped = page_image.crop(bbox)
                                    out_path = os.path.join(
                                        img_dir,
                                        f"pdf_img_page{page_index}_xref{im['xref']}.png"
                                    )
                                    cropped.save(out_path, format="PNG")
                                    images_with_section.append((last_section, {
                                        "type": "image",
                                        "path": out_path,
                                        "caption": None,
                                        "alt_text": "",
                                        "virtual": False
                                    }))
                                except Exception:
                                    continue
                    except Exception:
                        pass

            raw_text = "\n\n".join([t for t in page_texts if t]).strip()
            if not raw_text:
                QMessageBox.warning(self, "PDF Import", "No readable text found in PDF.")
                return

            # Parse combined text into sections
            self.doc_handler = DocumentHandler()
            self.doc_handler.parse_text(raw_text)

            # Attach tables & images to closest sections (T1 behavior)
            if tables_with_section or images_with_section:
                for sec_name, tbl in tables_with_section:
                    target = sec_name if sec_name in self.doc_handler.sections else self.doc_handler.detected_order[-1]
                    self.doc_handler.sections[target].append({
                        "type": "table",
                        "data": tbl
                    })

                for sec_name, img in images_with_section:
                    target = sec_name if sec_name in self.doc_handler.sections else self.doc_handler.detected_order[-1]
                    if isinstance(img, dict):
                        self.doc_handler.sections[target].append(img)
                    else:
                        # shouldn't happen, but just in case
                        self.doc_handler.sections[target].append({
                            "type": "image",
                            "path": img,
                            "caption": None,
                            "alt_text": "",
                            "virtual": False
                        })

                QMessageBox.information(
                    self,
                    "PDF Import Notice",
                    "PDF imported with smart extraction. Some image/table placement may be approximate."
                )

            self.text_input.setText(raw_text)
            self.refresh_structured_preview()
            QMessageBox.information(self, "PDF Imported", "PDF text imported and sections detected.")

        except Exception as e:
            QMessageBox.critical(self, "PDF Import Error", str(e))

    def detect_sections_from_text(self):
        raw = self.text_input.toPlainText().strip()
        if not raw:
            QMessageBox.warning(self, "Empty", "Paste or import text first.")
            return
        self.doc_handler = DocumentHandler()
        self.doc_handler.parse_text(raw)
        self.refresh_structured_preview()
        QMessageBox.information(self, "Sections Detected", "Sections detected from text.")

    # ====================================================
    # TAB 2: Structured Preview (raw model)
    # ====================================================
    def _init_structured_preview_tab(self):
        tab = QWidget()
        v = QVBoxLayout()

        self.section_list = QListWidget()
        self.section_preview = QTextEdit()
        self.section_preview.setReadOnly(True)

        v.addWidget(QLabel("Detected Sections (in order):"))
        v.addWidget(self.section_list)
        v.addWidget(QLabel("Section Content (raw structured view):"))
        v.addWidget(self.section_preview)

        btn_show = QPushButton("Show Selected Section")
        btn_show.clicked.connect(self.show_selected_section)
        v.addWidget(btn_show)

        tab.setLayout(v)
        self.tabs.addTab(tab, "Structured Preview")

    def refresh_structured_preview(self):
        self.section_list.clear()
        for sec in self.doc_handler.detected_order:
            self.section_list.addItem(sec)

    def show_selected_section(self):
        idx = self.section_list.currentRow()
        if idx < 0:
            return

        sec_name = self.doc_handler.detected_order[idx]
        elems = self.doc_handler.sections.get(sec_name, [])

        lines = []
        for el in elems:
            t = el.get("type")
            if t == "text":
                lines.append(el.get("content", ""))
            elif t == "table":
                lines.append("[TABLE]")
                for row in el.get("data", []):
                    lines.append(" | ".join(row))
                lines.append("[/TABLE]")
            elif t == "image":
                lines.append(
                    f"[IMAGE] path={el.get('path')}  VIRTUAL={el.get('virtual')}  "
                    f"CAPTION={el.get('caption')}  ALT={el.get('alt_text')}"
                )

        self.section_preview.setPlainText("\n".join(lines))

    # ====================================================
    # TAB 3: Formatting & Style
    # ====================================================
    def _init_style_tab(self):
        tab = QWidget()
        v = QVBoxLayout()

        self.style_combo = QComboBox()
        self.style_combo.addItems(["APA", "IEEE", "MLA", "Custom"])
        self.style_combo.currentTextChanged.connect(self.on_style_change)

        v.addWidget(QLabel("Select Formatting Style:"))
        v.addWidget(self.style_combo)

        self.custom_group = QGroupBox("Custom Style Settings")
        cv = QVBoxLayout()

        self.font_combo = QComboBox()
        self.font_combo.addItems(["Times New Roman", "Arial", "Calibri", "Georgia"])

        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(8, 24)

        self.line_spacing_spin = QDoubleSpinBox()
        self.line_spacing_spin.setRange(1.0, 3.0)
        self.line_spacing_spin.setSingleStep(0.25)

        self.margin_spin = QDoubleSpinBox()
        self.margin_spin.setRange(0.5, 2.0)
        self.margin_spin.setSingleStep(0.25)

        self.align_combo = QComboBox()
        self.align_combo.addItems(["Left", "Justify"])

        for label, widget in [
            ("Font:", self.font_combo),
            ("Font Size:", self.font_size_spin),
            ("Line Spacing:", self.line_spacing_spin),
            ("Margins (inches):", self.margin_spin),
            ("Paragraph Alignment:", self.align_combo),
        ]:
            cv.addWidget(QLabel(label))
            cv.addWidget(widget)

        self.custom_group.setLayout(cv)
        v.addWidget(self.custom_group)

        # Author / Institution
        v.addWidget(QLabel("Author Name:"))
        self.author_field = QLineEdit()
        v.addWidget(self.author_field)

        v.addWidget(QLabel("Institution Name:"))
        self.inst_field = QLineEdit()
        v.addWidget(self.inst_field)

        btn_apply = QPushButton("Apply Style")
        btn_apply.clicked.connect(self.apply_style)
        v.addWidget(btn_apply)

        tab.setLayout(v)
        self.tabs.addTab(tab, "Formatting & Style")

        # Initialize with APA defaults
        self.on_style_change("APA")

    def on_style_change(self, style_name):
        self.selected_style = style_name
        preset = STYLE_PRESETS.get(style_name, STYLE_PRESETS["APA"])

        self.custom_group.setVisible(style_name == "Custom")

        self.font_combo.setCurrentText(preset.get("font", "Times New Roman"))
        self.font_size_spin.setValue(preset.get("font_size", 12))
        self.line_spacing_spin.setValue(preset.get("line_spacing", 1.5))
        self.margin_spin.setValue(preset.get("margin_in", 1.0))
        align = preset.get("alignment", "left").lower()
        self.align_combo.setCurrentText("Justify" if align == "justify" else "Left")

    def apply_style(self):
        self.author = self.author_field.text().strip() or "Anonymous"
        self.institution = self.inst_field.text().strip() or "Unknown Institution"

        if self.selected_style == "Custom":
            self.custom_style = {
                "font": self.font_combo.currentText(),
                "font_size": self.font_size_spin.value(),
                "line_spacing": self.line_spacing_spin.value(),
                "margin_in": self.margin_spin.value(),
                "heading_size": 14,
                "alignment": self.align_combo.currentText().lower(),
                "title_page": True
            }
            QMessageBox.information(self, "Style", "Custom style applied.")
        else:
            self.custom_style = {}
            QMessageBox.information(self, "Style", f"{self.selected_style} style selected.")

    # ====================================================
    # TAB 4: Formatted Preview (DOCX -> text)
    # ====================================================
    def _init_formatted_preview_tab(self):
        tab = QWidget()
        v = QVBoxLayout()

        self.preview_box = QTextEdit()
        self.preview_box.setReadOnly(True)
        v.addWidget(self.preview_box)

        btn_gen = QPushButton("Generate Formatted Preview")
        btn_gen.clicked.connect(self.generate_preview)
        v.addWidget(btn_gen)

        tab.setLayout(v)
        self.tabs.addTab(tab, "Formatted Preview")

    def generate_preview(self):
        if not self.doc_handler.sections:
            QMessageBox.warning(self, "Error", "No structured content to preview.")
            return

        tmp_path = os.path.join(tempfile.gettempdir(), "docmaster_preview.docx")
        exporter = ExportService(
            style_choice=self.selected_style,
            custom_style=self.custom_style if self.selected_style == "Custom" else None,
            author=self.author,
            institution=self.institution
        )

        try:
            exporter.export_to_docx(self.doc_handler.sections, "Preview Document", tmp_path)
        except Exception as e:
            QMessageBox.critical(self, "Preview Error", f"Failed to generate DOCX preview:\n{e}")
            return

        try:
            txt = docx2txt.process(tmp_path)
            if len(txt) > 200000:
                txt = txt[:200000] + "\n\n[Preview truncated]"
            self.preview_box.setPlainText(txt)
            QMessageBox.information(self, "Preview", "Formatted preview generated from DOCX.")
        except Exception as e:
            QMessageBox.critical(self, "Preview Error", f"Failed to read generated DOCX:\n{e}")

    # ====================================================
    # TAB 5: Export
    # ====================================================
    def _init_export_tab(self):
        tab = QWidget()
        v = QVBoxLayout()

        v.addWidget(QLabel("Export your fully formatted document:"))

        btn_docx = QPushButton("Export to DOCX")
        btn_docx.clicked.connect(self.export_docx)
        v.addWidget(btn_docx)

        btn_pdf = QPushButton("Export to PDF")
        btn_pdf.clicked.connect(self.export_pdf)
        v.addWidget(btn_pdf)

        tab.setLayout(v)
        self.tabs.addTab(tab, "Export")

    def export_docx(self):
        if not self.doc_handler.sections:
            QMessageBox.warning(self, "Error", "No content to export.")
            return

        out, _ = QFileDialog.getSaveFileName(self, "Save DOCX", "", "Word Document (*.docx)")
        if not out:
            return

        exporter = ExportService(
            style_choice=self.selected_style,
            custom_style=self.custom_style if self.selected_style == "Custom" else None,
            author=self.author,
            institution=self.institution
        )
        try:
            exporter.export_to_docx(self.doc_handler.sections, "Final Document", out)
            QMessageBox.information(self, "Export", f"DOCX saved to: {out}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export DOCX:\n{e}")

    def export_pdf(self):
        if not self.doc_handler.sections:
            QMessageBox.warning(self, "Error", "No content to export.")
            return

        out, _ = QFileDialog.getSaveFileName(self, "Save PDF", "", "PDF File (*.pdf)")
        if not out:
            return

        exporter = ExportService(
            style_choice=self.selected_style,
            custom_style=self.custom_style if self.selected_style == "Custom" else None,
            author=self.author,
            institution=self.institution
        )
        try:
            exporter.export_to_pdf(self.doc_handler.sections, "Final Document", out)
            QMessageBox.information(self, "Export", f"PDF saved to: {out}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export PDF:\n{e}")


# ==============================
# MAIN
# ==============================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = DocMasterApp()
    win.show()
    sys.exit(app.exec_())
