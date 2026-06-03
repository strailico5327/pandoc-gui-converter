#!/usr/bin/env python3
"""
Pandoc GUI Converter
A small English GUI for converting Markdown <-> Word DOCX with Pandoc.

Requirements:
  - Python 3.9+
  - PySide6
  - Pandoc installed and available in PATH, or choose pandoc.exe in the GUI.
"""

from __future__ import annotations

import os
import re
import ctypes
import shlex
import shutil
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import QObject, QThread, Qt, Signal
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

APP_TITLE = "Pandoc Markdown / DOCX Converter"
ABOUT_TEXT = """Pandoc Markdown / DOCX Converter
© 2026 strailico5327

Convert Markdown-like files to DOCX, and DOCX files back to Markdown, through Pandoc.

Licensed under GNU GPLv3."""

MARKDOWN_EXTS = {".md", ".markdown", ".mdown", ".mkd", ".txt"}
DOCX_EXTS = {".docx"}


def enable_high_dpi_awareness() -> None:
    """
    Make the window crisp and correctly sized on high-DPI Windows displays.

    This must run before creating QApplication. It is safe to ignore failures
    because macOS/Linux and older Windows builds handle DPI differently.
    """
    if not sys.platform.startswith("win"):
        return

    try:
        awareness_context_per_monitor_v2 = ctypes.c_void_p(-4)
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(awareness_context_per_monitor_v2):
            return
    except Exception:
        pass

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass

    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def quote_cmd(args: list[str]) -> str:
    """Return a readable command line for the log box."""
    if os.name == "nt":
        return subprocess.list2cmdline(args)
    return " ".join(shlex.quote(a) for a in args)


def open_folder(path: Path) -> None:
    folder = path if path.is_dir() else path.parent
    if not folder.exists():
        raise FileNotFoundError(str(folder))

    if sys.platform.startswith("win"):
        os.startfile(str(folder))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(folder)])
    else:
        subprocess.Popen(["xdg-open", str(folder)])


def safe_output_path(input_path: Path, mode: str) -> Path:
    """Generate a default output path."""
    if mode == "md_to_docx":
        return input_path.with_suffix(".docx")
    if mode == "docx_to_md":
        return input_path.with_suffix(".md")

    ext = input_path.suffix.lower()
    if ext in MARKDOWN_EXTS:
        return input_path.with_suffix(".docx")
    if ext in DOCX_EXTS:
        return input_path.with_suffix(".md")
    return input_path.with_name(input_path.stem + "_converted")


def infer_mode(input_path: Path, selected_mode: str) -> str:
    """Return md_to_docx or docx_to_md."""
    if selected_mode != "auto":
        return selected_mode

    ext = input_path.suffix.lower()
    if ext in MARKDOWN_EXTS:
        return "md_to_docx"
    if ext in DOCX_EXTS:
        return "docx_to_md"
    raise ValueError("Cannot infer conversion mode from this file extension. Choose a mode manually.")


def remove_single_line_breaks(markdown: str) -> str:
    """
    Conservative post-processor for Markdown output.

    It joins single line breaks inside ordinary paragraphs, but keeps structural
    Markdown lines such as headings, lists, block quotes, tables, code fences,
    horizontal rules, HTML-ish lines, and image-only lines.
    """
    lines = markdown.splitlines()
    out: list[str] = []
    paragraph: list[str] = []
    in_fence = False
    fence_marker = ""

    def flush_para() -> None:
        nonlocal paragraph
        if paragraph:
            out.append(" ".join(x.strip() for x in paragraph if x.strip()))
            paragraph = []

    def is_structural(line: str) -> bool:
        s = line.strip()
        if not s:
            return True
        if re.match(r"^#{1,6}\s+", s):
            return True
        if re.match(r"^([-*_])(?:\s*\1){2,}\s*$", s):
            return True
        if re.match(r"^\s*([-+*]|\d+[.)])\s+", line):
            return True
        if re.match(r"^\s{0,3}>", line):
            return True
        if s.startswith("|") and s.endswith("|"):
            return True
        if re.match(r"^\s*[:\-| ]{3,}\s*$", line) and "|" in line:
            return True
        if s.startswith("!") and "](" in s:
            return True
        if re.match(r"^\s*</?\w+", s):
            return True
        return False

    for line in lines:
        stripped = line.strip()

        fence_match = re.match(r"^\s*(```+|~~~+)", line)
        if fence_match:
            marker = fence_match.group(1)[:3]
            if not in_fence:
                flush_para()
                in_fence = True
                fence_marker = marker
                out.append(line)
            elif stripped.startswith(fence_marker):
                out.append(line)
                in_fence = False
                fence_marker = ""
            else:
                out.append(line)
            continue

        if in_fence:
            out.append(line)
            continue

        if not stripped:
            flush_para()
            out.append("")
            continue

        if is_structural(line):
            flush_para()
            out.append(line)
            continue

        paragraph.append(line)

    flush_para()
    return "\n".join(out).rstrip() + "\n"


class ConversionWorker(QObject):
    log = Signal(str)
    finished = Signal(bool)

    def __init__(self, args: list[str], mode: str, input_path: Path, output_path: Path, remove_soft_breaks: bool) -> None:
        super().__init__()
        self.args = args
        self.mode = mode
        self.input_path = input_path
        self.output_path = output_path
        self.remove_soft_breaks = remove_soft_breaks

    def run(self) -> None:
        try:
            result = subprocess.run(
                self.args,
                cwd=str(self.input_path.parent),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            if result.stdout:
                self.log.emit(result.stdout.rstrip())
            if result.stderr:
                self.log.emit(result.stderr.rstrip())

            if result.returncode != 0:
                self.log.emit(f"Conversion failed with exit code {result.returncode}.")
                self.finished.emit(False)
                return

            if self.mode == "docx_to_md" and self.remove_soft_breaks:
                self.log.emit("Post-processing Markdown: removing single line breaks inside ordinary paragraphs...")
                text = self.output_path.read_text(encoding="utf-8", errors="replace")
                self.output_path.write_text(remove_single_line_breaks(text), encoding="utf-8", newline="\n")

            self.log.emit(f"Done: {self.output_path}")
            self.finished.emit(True)
        except Exception as exc:
            self.log.emit(f"Error: {exc}")
            self.finished.emit(False)


class PandocGui(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(980, 880)
        self.setMinimumSize(880, 760)

        self._running = False
        self.thread: QThread | None = None
        self.worker: ConversionWorker | None = None

        self._build_ui()
        self._apply_styles()
        self._refresh_command_preview()
        self._check_pandoc_on_startup()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        title = QLabel("Pandoc Markdown / DOCX Converter")
        title.setObjectName("titleLabel")
        root.addWidget(title)

        self._file_section(root)
        self._mode_section(root)
        self._options_section(root)
        self._command_section(root)
        self._buttons_section(root)
        self._log_section(root)
        self._footer_section(root)

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QWidget {
                background: palette(window);
                color: palette(window-text);
                font-family: "Segoe UI";
                font-size: 10pt;
            }
            QLabel#titleLabel {
                font-size: 16pt;
                font-weight: 700;
            }
            QLineEdit, QTextEdit, QComboBox {
                background: palette(base);
                color: palette(text);
                border: 1px solid palette(mid);
            }
            QLineEdit:disabled, QTextEdit:disabled, QComboBox:disabled {
                background: palette(alternate-base);
                color: palette(mid);
            }
            QTextEdit#monoBox {
                font-family: Consolas, "Cascadia Mono", monospace;
            }
            QPushButton {
                padding: 6px 12px;
            }
            QPushButton#runButton {
                font-weight: 700;
            }
            QPushButton#aboutButton {
                min-width: 28px;
                max-width: 28px;
                min-height: 28px;
                max-height: 28px;
                padding: 0;
                border-radius: 14px;
                font-size: 13pt;
            }
            QGroupBox {
                font-weight: 600;
                margin-top: 8px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 4px;
            }
            QRadioButton, QCheckBox {
                spacing: 8px;
            }
            """
        )

    def _file_section(self, root: QVBoxLayout) -> None:
        frame = QGroupBox("Files")
        layout = QGridLayout(frame)

        self.pandoc_path = QLineEdit(shutil.which("pandoc") or "pandoc")
        self.input_file = QLineEdit()
        self.output_file = QLineEdit()

        layout.addWidget(QLabel("Pandoc executable:"), 0, 0)
        layout.addWidget(self.pandoc_path, 0, 1)
        browse_pandoc = QPushButton("Browse...")
        browse_pandoc.clicked.connect(self._browse_pandoc)
        layout.addWidget(browse_pandoc, 0, 2)

        layout.addWidget(QLabel("Input file:"), 1, 0)
        layout.addWidget(self.input_file, 1, 1)
        browse_input = QPushButton("Browse...")
        browse_input.clicked.connect(self._browse_input)
        layout.addWidget(browse_input, 1, 2)

        layout.addWidget(QLabel("Output file:"), 2, 0)
        layout.addWidget(self.output_file, 2, 1)
        browse_output = QPushButton("Save as...")
        browse_output.clicked.connect(self._browse_output)
        layout.addWidget(browse_output, 2, 2)

        layout.setColumnStretch(1, 1)
        root.addWidget(frame)

        for edit in (self.pandoc_path, self.input_file, self.output_file):
            edit.textChanged.connect(self._refresh_command_preview)

    def _mode_section(self, root: QVBoxLayout) -> None:
        frame = QGroupBox("Conversion mode")
        layout = QHBoxLayout(frame)

        self.mode_auto = QRadioButton("Auto detect")
        self.mode_md_to_docx = QRadioButton("Markdown -> Word DOCX")
        self.mode_docx_to_md = QRadioButton("Word DOCX -> Markdown")
        self.mode_auto.setChecked(True)

        for button in (self.mode_auto, self.mode_md_to_docx, self.mode_docx_to_md):
            layout.addWidget(button)
            button.toggled.connect(self._on_mode_change)
        layout.addStretch()
        root.addWidget(frame)

    def _options_section(self, root: QVBoxLayout) -> None:
        frame = QGroupBox("Common options")
        layout = QGridLayout(frame)

        self.markdown_writer = QComboBox()
        self.markdown_writer.addItems(["gfm", "markdown", "markdown_strict", "commonmark_x", "commonmark"])
        self.markdown_reader = QLineEdit("markdown+pipe_tables+raw_html+tex_math_dollars")
        self.wrap_mode = QComboBox()
        self.wrap_mode.addItems(["none", "auto", "preserve"])
        self.eol_mode = QComboBox()
        self.eol_mode.addItems(["native", "lf", "crlf"])
        self.reference_doc = QLineEdit()
        self.media_dir = QLineEdit()
        self.extra_args = QLineEdit()

        layout.addWidget(QLabel("Markdown output format:"), 0, 0)
        layout.addWidget(self.markdown_writer, 0, 1)
        layout.addWidget(QLabel("Markdown input format:"), 0, 2)
        layout.addWidget(self.markdown_reader, 0, 3)

        layout.addWidget(QLabel("Wrap Markdown source:"), 1, 0)
        layout.addWidget(self.wrap_mode, 1, 1)
        layout.addWidget(QLabel("Line endings:"), 1, 2)
        layout.addWidget(self.eol_mode, 1, 3)

        layout.addWidget(QLabel("Reference DOCX:"), 2, 0)
        layout.addWidget(self.reference_doc, 2, 1, 1, 2)
        browse_reference = QPushButton("Browse...")
        browse_reference.clicked.connect(self._browse_reference_doc)
        layout.addWidget(browse_reference, 2, 3)

        layout.addWidget(QLabel("Media folder:"), 3, 0)
        layout.addWidget(self.media_dir, 3, 1, 1, 2)
        browse_media = QPushButton("Browse...")
        browse_media.clicked.connect(self._browse_media_dir)
        layout.addWidget(browse_media, 3, 3)

        checks = QWidget()
        checks_layout = QGridLayout(checks)
        checks_layout.setContentsMargins(0, 0, 0, 0)
        self.standalone = QCheckBox("Standalone document")
        self.standalone.setChecked(True)
        self.toc = QCheckBox("Table of contents")
        self.number_sections = QCheckBox("Number sections")
        self.overwrite = QCheckBox("Overwrite output")
        self.extract_media = QCheckBox("Extract images from DOCX")
        self.extract_media.setChecked(True)
        self.remove_soft_breaks = QCheckBox("Remove single line breaks in Markdown")
        self.remove_soft_breaks.setChecked(True)
        self.preserve_docx_styles = QCheckBox("Preserve DOCX styles as custom-style")

        checkboxes = [
            self.standalone,
            self.toc,
            self.number_sections,
            self.overwrite,
            self.extract_media,
            self.remove_soft_breaks,
            self.preserve_docx_styles,
        ]
        for index, checkbox in enumerate(checkboxes):
            checks_layout.addWidget(checkbox, index // 2, index % 2)
            checkbox.toggled.connect(self._refresh_command_preview)

        layout.addWidget(checks, 4, 0, 1, 4)
        layout.addWidget(QLabel("Extra Pandoc arguments:"), 5, 0)
        layout.addWidget(self.extra_args, 5, 1, 1, 3)

        layout.setColumnStretch(1, 1)
        layout.setColumnStretch(3, 1)
        root.addWidget(frame)

        for edit in (self.markdown_reader, self.reference_doc, self.media_dir, self.extra_args):
            edit.textChanged.connect(self._refresh_command_preview)
        for combo in (self.markdown_writer, self.wrap_mode, self.eol_mode):
            combo.currentTextChanged.connect(self._refresh_command_preview)

    def _command_section(self, root: QVBoxLayout) -> None:
        frame = QGroupBox("Command preview")
        layout = QVBoxLayout(frame)
        self.command_box = QTextEdit()
        self.command_box.setReadOnly(True)
        self.command_box.setFixedHeight(82)
        self.command_box.setObjectName("monoBox")
        layout.addWidget(self.command_box)
        root.addWidget(frame)

    def _buttons_section(self, root: QVBoxLayout) -> None:
        layout = QHBoxLayout()
        self.run_button = QPushButton("Convert")
        self.run_button.setObjectName("runButton")
        self.run_button.clicked.connect(self._run_conversion)
        check_button = QPushButton("Check Pandoc")
        check_button.clicked.connect(self._show_pandoc_version)
        open_button = QPushButton("Open output folder")
        open_button.clicked.connect(self._open_output_folder)
        clear_button = QPushButton("Clear log")
        clear_button.clicked.connect(self._clear_log)

        layout.addWidget(self.run_button)
        layout.addWidget(check_button)
        layout.addWidget(open_button)
        layout.addStretch()
        layout.addWidget(clear_button)
        root.addLayout(layout)

    def _log_section(self, root: QVBoxLayout) -> None:
        frame = QGroupBox("Log")
        layout = QVBoxLayout(frame)
        self.log_box = QTextEdit()
        self.log_box.setObjectName("monoBox")
        layout.addWidget(self.log_box)
        root.addWidget(frame, stretch=1)

    def _footer_section(self, root: QVBoxLayout) -> None:
        layout = QHBoxLayout()
        self.about_button = QPushButton("ⓘ")
        self.about_button.setObjectName("aboutButton")
        self.about_button.setToolTip("About")
        self.about_button.clicked.connect(self._show_about)
        layout.addStretch()
        layout.addWidget(self.about_button)
        root.addLayout(layout)

    def mode_value(self) -> str:
        if self.mode_md_to_docx.isChecked():
            return "md_to_docx"
        if self.mode_docx_to_md.isChecked():
            return "docx_to_md"
        return "auto"

    def _browse_pandoc(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Choose pandoc executable", "", "Pandoc executable (pandoc.exe pandoc);;All files (*.*)")
        if path:
            self.pandoc_path.setText(path)

    def _browse_input(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose input file",
            "",
            "Markdown or Word (*.md *.markdown *.mdown *.mkd *.txt *.docx);;Markdown (*.md *.markdown *.mdown *.mkd *.txt);;Word DOCX (*.docx);;All files (*.*)",
        )
        if path:
            self.input_file.setText(path)
            try:
                mode = infer_mode(Path(path), self.mode_value())
                self.output_file.setText(str(safe_output_path(Path(path), mode)))
                if mode == "docx_to_md" and not self.media_dir.text().strip():
                    self.media_dir.setText(str(Path(path).with_name(Path(path).stem + "_media")))
            except Exception:
                self.output_file.setText(str(safe_output_path(Path(path), "auto")))

    def _browse_output(self) -> None:
        input_path = Path(self.input_file.text()) if self.input_file.text() else Path("output")
        try:
            mode = infer_mode(input_path, self.mode_value())
        except Exception:
            mode = "md_to_docx"

        default_ext = ".docx" if mode == "md_to_docx" else ".md"
        filetypes = "Word DOCX (*.docx)" if mode == "md_to_docx" else "Markdown (*.md);;All files (*.*)"
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Choose output file",
            str(input_path.with_suffix(default_ext)),
            filetypes,
        )
        if path:
            self.output_file.setText(path)

    def _browse_reference_doc(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Choose reference DOCX", "", "Word DOCX (*.docx);;All files (*.*)")
        if path:
            self.reference_doc.setText(path)

    def _browse_media_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Choose media extraction folder")
        if path:
            self.media_dir.setText(path)

    def _on_mode_change(self) -> None:
        if self.input_file.text().strip():
            try:
                input_path = Path(self.input_file.text())
                mode = infer_mode(input_path, self.mode_value())
                self.output_file.setText(str(safe_output_path(input_path, mode)))
                if mode == "docx_to_md" and not self.media_dir.text().strip():
                    self.media_dir.setText(str(input_path.with_name(input_path.stem + "_media")))
            except Exception:
                pass
        self._refresh_command_preview()

    def _build_command(self) -> tuple[list[str], str, Path, Path]:
        pandoc = self.pandoc_path.text().strip() or "pandoc"
        input_path = Path(self.input_file.text().strip()).expanduser()
        output_path = Path(self.output_file.text().strip()).expanduser()

        if not input_path.exists():
            raise FileNotFoundError(f"Input file does not exist: {input_path}")
        if not output_path.parent.exists():
            raise FileNotFoundError(f"Output folder does not exist: {output_path.parent}")
        if output_path.exists() and not self.overwrite.isChecked():
            raise FileExistsError("Output file already exists. Enable 'Overwrite output' or choose another output path.")

        mode = infer_mode(input_path, self.mode_value())
        args: list[str] = [pandoc, str(input_path), "-o", str(output_path)]

        if self.standalone.isChecked():
            args.append("--standalone")

        if self.eol_mode.currentText() != "native":
            args.append(f"--eol={self.eol_mode.currentText()}")

        if mode == "md_to_docx":
            args += ["-f", self.markdown_reader.text().strip() or "markdown", "-t", "docx"]
            args.append(f"--resource-path={input_path.parent}")
            reference = self.reference_doc.text().strip()
            if reference:
                ref_path = Path(reference).expanduser()
                if not ref_path.exists():
                    raise FileNotFoundError(f"Reference DOCX does not exist: {ref_path}")
                args.append(f"--reference-doc={ref_path}")
            if self.toc.isChecked():
                args.append("--toc")
            if self.number_sections.isChecked():
                args.append("--number-sections")

        elif mode == "docx_to_md":
            docx_reader = "docx+styles" if self.preserve_docx_styles.isChecked() else "docx"
            args += ["-f", docx_reader, "-t", self.markdown_writer.currentText().strip() or "gfm"]
            args.append(f"--wrap={self.wrap_mode.currentText()}")
            if self.extract_media.isChecked():
                media = self.media_dir.text().strip()
                if not media:
                    media = str(input_path.with_name(input_path.stem + "_media"))
                    self.media_dir.setText(media)
                args.append(f"--extract-media={Path(media).expanduser()}")
        else:
            raise ValueError(f"Unsupported conversion mode: {mode}")

        extra = self.extra_args.text().strip()
        if extra:
            args.extend(shlex.split(extra, posix=(os.name != "nt")))

        return args, mode, input_path, output_path

    def _refresh_command_preview(self) -> None:
        try:
            args, _, _, _ = self._build_command_for_preview()
            text = quote_cmd(args)
        except Exception as exc:
            text = f"Command preview unavailable: {exc}"

        self.command_box.setPlainText(text)

    def _build_command_for_preview(self) -> tuple[list[str], str, Path, Path]:
        pandoc = self.pandoc_path.text().strip() or "pandoc"
        input_text = self.input_file.text().strip() or "input.md"
        input_path = Path(input_text)
        try:
            mode = infer_mode(input_path, self.mode_value())
        except Exception:
            mode = "md_to_docx" if self.mode_value() == "auto" else self.mode_value()

        output_text = self.output_file.text().strip() or str(safe_output_path(input_path, mode))
        output_path = Path(output_text)
        args: list[str] = [pandoc, str(input_path), "-o", str(output_path)]

        if self.standalone.isChecked():
            args.append("--standalone")
        if self.eol_mode.currentText() != "native":
            args.append(f"--eol={self.eol_mode.currentText()}")

        if mode == "md_to_docx":
            args += ["-f", self.markdown_reader.text().strip() or "markdown", "-t", "docx"]
            args.append(f"--resource-path={input_path.parent}")
            if self.reference_doc.text().strip():
                args.append(f"--reference-doc={self.reference_doc.text().strip()}")
            if self.toc.isChecked():
                args.append("--toc")
            if self.number_sections.isChecked():
                args.append("--number-sections")
        else:
            docx_reader = "docx+styles" if self.preserve_docx_styles.isChecked() else "docx"
            args += ["-f", docx_reader, "-t", self.markdown_writer.currentText().strip() or "gfm"]
            args.append(f"--wrap={self.wrap_mode.currentText()}")
            if self.extract_media.isChecked():
                media = self.media_dir.text().strip() or str(input_path.with_name(input_path.stem + "_media"))
                args.append(f"--extract-media={media}")

        if self.extra_args.text().strip():
            try:
                args.extend(shlex.split(self.extra_args.text().strip(), posix=(os.name != "nt")))
            except ValueError:
                args.append("<invalid extra args>")

        return args, mode, input_path, output_path

    def _run_conversion(self) -> None:
        if self._running:
            return

        try:
            args, mode, input_path, output_path = self._build_command()
        except Exception as exc:
            QMessageBox.critical(self, "Cannot start conversion", str(exc))
            return

        self._running = True
        self.run_button.setEnabled(False)
        self._log("\n=== Conversion started ===")
        self._log(quote_cmd(args))

        self.thread = QThread()
        self.worker = ConversionWorker(args, mode, input_path, output_path, self.remove_soft_breaks.isChecked())
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.log.connect(self._log)
        self.worker.finished.connect(self._finish_conversion)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self._cleanup_thread)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()

    def _finish_conversion(self, success: bool) -> None:
        self._running = False
        self.run_button.setEnabled(True)
        if success:
            QMessageBox.information(self, "Conversion complete", "The conversion finished successfully.")
        else:
            QMessageBox.critical(self, "Conversion failed", "The conversion failed. Check the log for details.")

    def _cleanup_thread(self) -> None:
        self.thread = None
        self.worker = None

    def _show_pandoc_version(self) -> None:
        pandoc = self.pandoc_path.text().strip() or "pandoc"
        try:
            result = subprocess.run([pandoc, "--version"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            text = result.stdout or result.stderr
            self._log("\n=== Pandoc version ===")
            self._log(text.strip())
            if result.returncode != 0:
                QMessageBox.critical(self, "Pandoc check failed", text.strip() or "Pandoc returned an error.")
            else:
                QMessageBox.information(self, "Pandoc found", text.splitlines()[0] if text else "Pandoc found.")
        except Exception as exc:
            QMessageBox.critical(self, "Pandoc not found", str(exc))

    def _check_pandoc_on_startup(self) -> None:
        found = shutil.which("pandoc")
        if found:
            self._log(f"Pandoc detected: {found}")
        else:
            self._log("Pandoc was not found in PATH. Install Pandoc or choose pandoc.exe manually.")

    def _open_output_folder(self) -> None:
        path_text = self.output_file.text().strip()
        if not path_text:
            QMessageBox.warning(self, "No output file", "Choose an output file first.")
            return
        try:
            open_folder(Path(path_text).expanduser())
        except Exception as exc:
            QMessageBox.critical(self, "Cannot open folder", str(exc))

    def _clear_log(self) -> None:
        self.log_box.clear()

    def _show_about(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle(APP_TITLE)
        dialog.setModal(True)
        dialog.setMinimumWidth(360)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(16)

        text = QLabel(ABOUT_TEXT)
        text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        text.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        ok_button = QPushButton("OK")
        ok_button.clicked.connect(dialog.accept)

        button_row = QHBoxLayout()
        button_row.addStretch()
        button_row.addWidget(ok_button)
        button_row.addStretch()

        layout.addWidget(text)
        layout.addLayout(button_row)

        dialog.exec()

    def _log(self, text: str) -> None:
        self.log_box.append(text)
        self.log_box.moveCursor(QTextCursor.MoveOperation.End)


def main() -> None:
    enable_high_dpi_awareness()
    app = QApplication(sys.argv)
    window = PandocGui()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
