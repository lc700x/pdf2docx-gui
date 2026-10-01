import os

from PySide6.QtCore import QObject, QThread, Signal, Slot
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import (
    BodyLabel,
    CaptionLabel,
    CardWidget,
    CheckBox,
    FluentIcon,
    FluentWindow,
    IndeterminateProgressBar,
    LineEdit,
    MessageBox,
    NavigationItemPosition,
    PrimaryPushButton,
    PushButton,
    SubtitleLabel,
    Theme,
    TitleLabel,
    setTheme,
    setThemeColor,
)


def _success_message(result):
    message = "Done."
    if result.get("images") is not None:
        replaced, missed = result["images"]
        message = f"Done. {replaced} tables and equations kept as images"
        if missed:
            message += f", {missed} left as converted text"
        message += "."
    if result.get("polish") is not None:
        counts = result["polish"]
        message += (f" Restored {counts['spaces']} spaces,"
                    f" removed {counts['page breaks dropped']} page breaks.")
    return message


class FileDropZone(CardWidget):
    """Fluent card that accepts drops for one file extension."""

    fileDropped = Signal(str)
    browseRequested = Signal()

    def __init__(self, file_type, extension, parent=None):
        super().__init__(parent)
        self.file_type = file_type
        self.extension = extension.lower().lstrip(".")
        self.setAcceptDrops(True)
        self.setObjectName(f"{self.extension}DropZone")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(8)

        top_row = QHBoxLayout()
        top_row.addWidget(SubtitleLabel(f"Input {file_type}", self))
        top_row.addStretch(1)
        self.browse_button = PushButton(f"Browse {file_type}", self)
        self.browse_button.clicked.connect(self.browseRequested)
        top_row.addWidget(self.browse_button)
        layout.addLayout(top_row)

        self.drop_label = BodyLabel(f"Drop a {file_type} file here", self)
        layout.addWidget(self.drop_label)
        layout.addWidget(CaptionLabel(
            f"or choose a file with Browse {file_type}", self))
        self.path_label = CaptionLabel(f"No {file_type} selected", self)
        self.path_label.setWordWrap(True)
        layout.addWidget(self.path_label)

    def file_paths(self, mime_data):
        return [url.toLocalFile() for url in mime_data.urls()
                if url.isLocalFile()
                and url.toLocalFile().lower().endswith("." + self.extension)]

    def dragEnterEvent(self, event: QDragEnterEvent):
        if self.file_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if self.file_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        paths = self.file_paths(event.mimeData())
        if paths:
            self.fileDropped.emit(paths[0])
            event.acceptProposedAction()
        else:
            event.ignore()

    def set_selected_file(self, path):
        self.path_label.setText(path)


class ConversionWorker(QObject):
    """Worker object moved to a QThread; all UI updates use its signals."""

    statusChanged = Signal(str)
    succeeded = Signal(object)
    failed = Signal(str)
    finished = Signal()

    def __init__(self, converter, *args, **kwargs):
        super().__init__()
        self.converter = converter
        self.args = args
        self.kwargs = kwargs

    @Slot()
    def run(self):
        try:
            kwargs = dict(self.kwargs)
            kwargs["on_status"] = self.statusChanged.emit
            result = self.converter(*self.args, **kwargs)
        except Exception as exc:
            self.failed.emit(str(exc))
        else:
            self.succeeded.emit(result)
        finally:
            self.finished.emit()


def _start_worker(page, converter, *args, **kwargs):
    page._set_busy(True)
    thread = QThread(page)
    worker = ConversionWorker(converter, *args, **kwargs)
    worker.moveToThread(thread)
    thread.started.connect(worker.run)
    worker.statusChanged.connect(page.status_label.setText)
    worker.succeeded.connect(page._conversion_succeeded)
    worker.failed.connect(page._conversion_failed)
    worker.finished.connect(thread.quit)
    worker.finished.connect(worker.deleteLater)
    thread.finished.connect(page._thread_finished)
    thread.finished.connect(thread.deleteLater)
    page._thread = thread
    page._worker = worker
    thread.start()


class PdfConverterPage(QWidget):
    def __init__(self, converter, parent=None):
        super().__init__(parent)
        self.converter = converter
        self.setObjectName("pdfConverterPage")
        self.input_path = None
        self._thread = None
        self._worker = None

        self.setAcceptDrops(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(14)

        layout.addWidget(TitleLabel("PDF to DOCX", self))
        layout.addWidget(CaptionLabel(
            "Convert locally while keeping document content editable.", self))

        self.drop_zone = FileDropZone("PDF", "pdf", self)
        self.drop_zone.fileDropped.connect(self.set_input)
        self.drop_zone.browseRequested.connect(self._browse_pdf)
        layout.addWidget(self.drop_zone)

        output_card = CardWidget(self)
        output_layout = QVBoxLayout(output_card)
        output_layout.setContentsMargins(18, 14, 18, 14)
        output_layout.setSpacing(8)
        output_layout.addWidget(SubtitleLabel("Output document", output_card))
        output_row = QHBoxLayout()
        self.output_edit = LineEdit(output_card)
        self.output_edit.setPlaceholderText("Choose where to save the DOCX")
        output_row.addWidget(self.output_edit, 1)
        self.output_browse_button = PushButton("Save as", output_card)
        self.output_browse_button.clicked.connect(self._browse_output)
        output_row.addWidget(self.output_browse_button)
        output_layout.addLayout(output_row)
        layout.addWidget(output_card)

        options_card = CardWidget(self)
        options_layout = QVBoxLayout(options_card)
        options_layout.setContentsMargins(18, 14, 18, 14)
        options_layout.setSpacing(7)
        options_layout.addWidget(SubtitleLabel("Conversion options", options_card))

        self.images_option = CheckBox(
            "Keep tables and equations as images", options_card)
        self.images_option.setChecked(True)
        options_layout.addWidget(self.images_option)
        images_help = CaptionLabel(
            "Preserves their layout while body text stays editable.", options_card)
        images_help.setWordWrap(True)
        options_layout.addWidget(images_help)

        self.tidy_option = CheckBox(
            "Tidy the document for reading and editing", options_card)
        self.tidy_option.setChecked(True)
        options_layout.addWidget(self.tidy_option)
        tidy_help = CaptionLabel(
            "Substitutes missing fonts, restores spaces, and lets text flow.",
            options_card)
        tidy_help.setWordWrap(True)
        options_layout.addWidget(tidy_help)
        layout.addWidget(options_card)

        footer = QHBoxLayout()
        self.status_label = BodyLabel("Ready", self)
        footer.addWidget(self.status_label, 1)
        self.convert_button = PrimaryPushButton("Convert", self)
        self.convert_button.setEnabled(False)
        self.convert_button.clicked.connect(self.start_conversion)
        footer.addWidget(self.convert_button)
        layout.addLayout(footer)

        self.progress = IndeterminateProgressBar(self, start=False)
        self.progress.hide()
        layout.addWidget(self.progress)

    def set_input(self, path):
        if not path or not path.lower().endswith(".pdf"):
            return False
        self.input_path = path
        self.drop_zone.set_selected_file(path)
        self.output_edit.setText(os.path.splitext(path)[0] + ".docx")
        self.convert_button.setEnabled(True)
        self.status_label.setText("Ready to convert")
        return True

    def set_output(self, path):
        if path:
            self.output_edit.setText(path)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if self.drop_zone.file_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if self.drop_zone.file_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        paths = self.drop_zone.file_paths(event.mimeData())
        if paths:
            self.set_input(paths[0])
            event.acceptProposedAction()
        else:
            event.ignore()

    def _browse_pdf(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Choose a PDF file", "", "PDF files (*.pdf)")
        if path:
            self.set_input(path)

    def _browse_output(self):
        current = self.output_edit.text().strip()
        initial = current or (
            os.path.splitext(self.input_path)[0] + ".docx"
            if self.input_path else "output.docx")
        path, _ = QFileDialog.getSaveFileName(
            self, "Save DOCX as", initial, "Word documents (*.docx)")
        if path:
            if not path.lower().endswith(".docx"):
                path += ".docx"
            self.set_output(path)

    def start_conversion(self):
        if not self.input_path:
            return
        destination = self.output_edit.text().strip()
        if not destination:
            destination = os.path.splitext(self.input_path)[0] + ".docx"
        if not destination.lower().endswith(".docx"):
            destination += ".docx"
        self.output_edit.setText(destination)

        self.status_label.setText("Converting PDF…")
        _start_worker(
            self, self.converter,
            self.input_path, destination,
            snip=self.images_option.isChecked(),
            tidy=self.tidy_option.isChecked())

    def _set_busy(self, busy):
        self.setAcceptDrops(not busy)
        self.drop_zone.setEnabled(not busy)
        self.output_edit.setEnabled(not busy)
        self.output_browse_button.setEnabled(not busy)
        self.images_option.setEnabled(not busy)
        self.tidy_option.setEnabled(not busy)
        self.convert_button.setEnabled(not busy and self.input_path is not None)
        if busy:
            self.progress.start()
        else:
            self.progress.stop()
        self.progress.setVisible(busy)

    def _conversion_succeeded(self, result):
        self.status_label.setText(_success_message(result))

    def _conversion_failed(self, message):
        self.status_label.setText("Conversion failed")
        MessageBox("Conversion failed", message, self).exec()

    @Slot()
    def _thread_finished(self):
        self._worker = None
        self._thread = None
        self._set_busy(False)

    def closeEvent(self, event):
        if self._thread and self._thread.isRunning():
            self.status_label.setText("Please wait for the conversion to finish.")
            event.ignore()
            return
        event.accept()


def _default_tex_output(source_path):
    return os.path.splitext(source_path)[0] + "_from_tex.docx"


class TexConverterPage(QWidget):
    def __init__(self, converter, parent=None):
        super().__init__(parent)
        self.converter = converter
        self.input_path = None
        self._thread = None
        self._worker = None
        self.setObjectName("texConverterPage")
        self.setAcceptDrops(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(14)
        layout.addWidget(TitleLabel("LaTeX to Word", self))
        description = CaptionLabel(
            "Convert a TeX source with its figures and bibliography.", self)
        description.setWordWrap(True)
        layout.addWidget(description)

        self.drop_zone = FileDropZone("TeX source", "tex", self)
        self.drop_zone.fileDropped.connect(self.set_input)
        self.drop_zone.browseRequested.connect(self._browse_tex)
        layout.addWidget(self.drop_zone)

        output_card = CardWidget(self)
        output_layout = QVBoxLayout(output_card)
        output_layout.setContentsMargins(18, 14, 18, 14)
        output_layout.setSpacing(8)
        output_layout.addWidget(SubtitleLabel("Output document", output_card))
        output_row = QHBoxLayout()
        self.output_edit = LineEdit(output_card)
        self.output_edit.setPlaceholderText("Choose where to save the DOCX")
        output_row.addWidget(self.output_edit, 1)
        self.output_browse_button = PushButton("Save as", output_card)
        self.output_browse_button.clicked.connect(self._browse_output)
        output_row.addWidget(self.output_browse_button)
        output_layout.addLayout(output_row)
        layout.addWidget(output_card)

        footer = QHBoxLayout()
        self.status_label = BodyLabel("Ready", self)
        footer.addWidget(self.status_label, 1)
        self.convert_button = PrimaryPushButton("Convert", self)
        self.convert_button.setEnabled(False)
        self.convert_button.clicked.connect(self.start_conversion)
        footer.addWidget(self.convert_button)
        layout.addLayout(footer)

        self.progress = IndeterminateProgressBar(self, start=False)
        self.progress.hide()
        layout.addWidget(self.progress)

    def set_input(self, path):
        if not path or not path.lower().endswith(".tex"):
            return False
        self.input_path = path
        self.drop_zone.set_selected_file(path)
        self.output_edit.setText(_default_tex_output(path))
        self.convert_button.setEnabled(True)
        self.status_label.setText("Ready to convert")
        return True

    def set_output(self, path):
        if path:
            self.output_edit.setText(path)

    def file_paths(self, mime_data):
        return self.drop_zone.file_paths(mime_data)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if self.file_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if self.file_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        paths = self.file_paths(event.mimeData())
        if paths:
            self.set_input(paths[0])
            event.acceptProposedAction()
        else:
            event.ignore()

    def _browse_tex(self):
        initial = os.path.dirname(self.input_path) if self.input_path else ""
        path, _ = QFileDialog.getOpenFileName(
            self, "Choose a LaTeX source", initial, "LaTeX files (*.tex)")
        if path:
            self.set_input(path)

    def _browse_output(self):
        initial = self.output_edit.text().strip() or (
            _default_tex_output(self.input_path) if self.input_path
            else "output_from_tex.docx")
        path, _ = QFileDialog.getSaveFileName(
            self, "Save DOCX as", initial, "Word documents (*.docx)")
        if path:
            if not path.lower().endswith(".docx"):
                path += ".docx"
            self.set_output(path)

    def start_conversion(self):
        if not self.input_path:
            return
        destination = self.output_edit.text().strip()
        if not destination:
            destination = _default_tex_output(self.input_path)
        if not destination.lower().endswith(".docx"):
            destination += ".docx"
        self.output_edit.setText(destination)
        self.status_label.setText("Converting LaTeX…")
        _start_worker(self, self.converter, self.input_path, destination)

    def _set_busy(self, busy):
        self.setAcceptDrops(not busy)
        self.drop_zone.setEnabled(not busy)
        self.output_edit.setEnabled(not busy)
        self.output_browse_button.setEnabled(not busy)
        self.convert_button.setEnabled(not busy and self.input_path is not None)
        if busy:
            self.progress.start()
        else:
            self.progress.stop()
        self.progress.setVisible(busy)

    def _conversion_succeeded(self, destination):
        self.status_label.setText(f"Done. Saved to {destination}")

    def _conversion_failed(self, message):
        self.status_label.setText("Conversion failed")
        MessageBox("Conversion failed", message, self).exec()

    @Slot()
    def _thread_finished(self):
        self._worker = None
        self._thread = None
        self._set_busy(False)


class ConverterWindow(FluentWindow):
    def __init__(self, converter, latex_converter=None, parent=None):
        super().__init__(parent)
        if latex_converter is None:
            from latex_to_docx import convert_latex_to_docx

            latex_converter = convert_latex_to_docx

        setTheme(Theme.AUTO)
        setThemeColor("#0078D4")
        self.setWindowTitle("PDF and LaTeX to Word Converter")
        self.setMinimumWidth(820)
        self.resize(940, 720)
        self.setAcceptDrops(True)

        self.pdf_page = PdfConverterPage(converter, self)
        self.tex_page = TexConverterPage(latex_converter, self)
        self.addSubInterface(
            self.pdf_page, FluentIcon.DOCUMENT, "PDF to Word",
            position=NavigationItemPosition.TOP)
        self.addSubInterface(
            self.tex_page, FluentIcon.CODE, "LaTeX to Word",
            position=NavigationItemPosition.TOP)

    def set_input(self, path):
        """Keep the existing PDF preload entry point for CLI GUI startup."""
        return self.pdf_page.set_input(path)

    def dragEnterEvent(self, event: QDragEnterEvent):
        page = self.stackedWidget.currentWidget()
        if page.file_paths(event.mimeData()) and not page._thread:
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        self.dragEnterEvent(event)

    def dropEvent(self, event: QDropEvent):
        page = self.stackedWidget.currentWidget()
        paths = page.file_paths(event.mimeData())
        if paths and not page._thread:
            page.set_input(paths[0])
            event.acceptProposedAction()
        else:
            event.ignore()

    def closeEvent(self, event):
        pages = (self.pdf_page, self.tex_page)
        active = [page for page in pages
                  if page._thread and page._thread.isRunning()]
        if active:
            active[0].status_label.setText(
                "Please wait for the conversion to finish.")
            event.ignore()
            return
        event.accept()
