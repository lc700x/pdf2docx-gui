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
    LineEdit,
    MessageBox,
    PrimaryPushButton,
    ProgressBar,
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


class PdfDropZone(CardWidget):
    """Fluent card that accepts PDF drops and offers a browse action."""

    pdfDropped = Signal(str)
    browseRequested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setObjectName("pdfDropZone")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(8)

        top_row = QHBoxLayout()
        top_row.addWidget(SubtitleLabel("Input PDF", self))
        top_row.addStretch(1)
        self.browse_button = PushButton("Browse PDF", self)
        self.browse_button.clicked.connect(self.browseRequested)
        top_row.addWidget(self.browse_button)
        layout.addLayout(top_row)

        self.drop_label = BodyLabel("Drop a PDF here", self)
        layout.addWidget(self.drop_label)
        layout.addWidget(CaptionLabel("or choose a file with Browse PDF", self))
        self.path_label = CaptionLabel("No PDF selected", self)
        self.path_label.setWordWrap(True)
        layout.addWidget(self.path_label)

    @staticmethod
    def _pdf_paths(mime_data):
        return [url.toLocalFile() for url in mime_data.urls()
                if url.isLocalFile() and url.toLocalFile().lower().endswith(".pdf")]

    def dragEnterEvent(self, event: QDragEnterEvent):
        if self._pdf_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if self._pdf_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        paths = self._pdf_paths(event.mimeData())
        if paths:
            self.pdfDropped.emit(paths[0])
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

    def __init__(self, converter, src, dst, snip, tidy):
        super().__init__()
        self.converter = converter
        self.src = src
        self.dst = dst
        self.snip = snip
        self.tidy = tidy

    @Slot()
    def run(self):
        try:
            result = self.converter(
                self.src, self.dst, snip=self.snip, tidy=self.tidy,
                on_status=self.statusChanged.emit)
        except Exception as exc:
            self.failed.emit(str(exc))
        else:
            self.succeeded.emit(result)
        finally:
            self.finished.emit()


class ConverterWindow(QWidget):
    def __init__(self, converter, parent=None):
        super().__init__(parent)
        self.converter = converter
        self.input_path = None
        self._thread = None
        self._worker = None

        setTheme(Theme.AUTO)
        setThemeColor("#0078D4")

        self.setWindowTitle("PDF to DOCX Converter")
        self.setMinimumWidth(560)
        self.resize(640, 600)
        self.setAcceptDrops(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(14)

        layout.addWidget(TitleLabel("PDF to DOCX", self))
        layout.addWidget(CaptionLabel(
            "Convert locally while keeping document content editable.", self))

        self.drop_zone = PdfDropZone(self)
        self.drop_zone.pdfDropped.connect(self.set_input)
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

        self.progress = ProgressBar(self)
        self.progress.setRange(0, 0)
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
        if PdfDropZone._pdf_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if PdfDropZone._pdf_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        paths = PdfDropZone._pdf_paths(event.mimeData())
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

        self._set_busy(True)
        self.status_label.setText("Converting PDF…")
        thread = QThread(self)
        worker = ConversionWorker(
            self.converter, self.input_path, destination,
            self.images_option.isChecked(), self.tidy_option.isChecked())
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.statusChanged.connect(self.status_label.setText)
        worker.succeeded.connect(self._conversion_succeeded)
        worker.failed.connect(self._conversion_failed)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(self._thread_finished)
        thread.finished.connect(thread.deleteLater)
        self._thread = thread
        self._worker = worker
        thread.start()

    def _set_busy(self, busy):
        self.setAcceptDrops(not busy)
        self.drop_zone.setEnabled(not busy)
        self.output_edit.setEnabled(not busy)
        self.output_browse_button.setEnabled(not busy)
        self.images_option.setEnabled(not busy)
        self.tidy_option.setEnabled(not busy)
        self.convert_button.setEnabled(not busy and self.input_path is not None)
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
