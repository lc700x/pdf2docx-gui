import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QEventLoop, QMimeData, QPoint, QPointF, Qt, QTimer, QUrl
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import QApplication

from fluent_ui import (
    ConversionWorker,
    ConverterWindow,
    PdfDropZone,
    _success_message,
)


class FluentUiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_conversion_options_default_to_enabled_and_wait_for_pdf(self):
        window = ConverterWindow(lambda *args, **kwargs: {})
        self.addCleanup(window.close)

        self.assertTrue(window.images_option.isChecked())
        self.assertTrue(window.tidy_option.isChecked())
        self.assertFalse(window.convert_button.isEnabled())

    def test_success_message_reports_the_correct_page_break_count(self):
        self.assertEqual(
            _success_message({
                "images": (12, 2),
                "polish": {"spaces": 8, "page breaks dropped": 5},
            }),
            "Done. 12 tables and equations kept as images, 2 left as converted "
            "text. Restored 8 spaces, removed 5 page breaks.")

    def test_pdf_selection_sets_default_output_and_rejects_other_files(self):
        window = ConverterWindow(lambda *args, **kwargs: {})
        self.addCleanup(window.close)

        self.assertFalse(window.set_input("paper.docx"))
        self.assertTrue(window.set_input("/tmp/paper.pdf"))
        self.assertEqual(window.output_edit.text(), "/tmp/paper.docx")
        self.assertTrue(window.convert_button.isEnabled())

        window.output_edit.setText("/tmp/custom.docx")
        self.assertEqual(window.output_edit.text(), "/tmp/custom.docx")

    def test_drop_accepts_pdf_and_uses_the_first_pdf_url(self):
        zone = PdfDropZone()
        self.addCleanup(zone.close)
        mime_data = QMimeData()
        mime_data.setUrls([
            QUrl.fromLocalFile("/tmp/readme.txt"),
            QUrl.fromLocalFile("/tmp/paper.PDF"),
        ])
        dropped = []
        zone.pdfDropped.connect(dropped.append)

        drag = QDragEnterEvent(
            QPoint(5, 5), Qt.CopyAction, mime_data,
            Qt.LeftButton, Qt.NoModifier)
        zone.dragEnterEvent(drag)
        self.assertTrue(drag.isAccepted())

        drop = QDropEvent(
            QPointF(5, 5), Qt.CopyAction, mime_data,
            Qt.LeftButton, Qt.NoModifier)
        zone.dropEvent(drop)
        self.assertTrue(drop.isAccepted())
        self.assertEqual(dropped, ["/tmp/paper.PDF"])

    def test_worker_emits_status_and_success(self):
        results = []
        statuses = []
        finished = []

        def convert(src, dst, snip, tidy, on_status):
            on_status("Tidying")
            return {"images": None, "polish": None}

        worker = ConversionWorker(convert, "input.pdf", "output.docx", True, True)
        worker.succeeded.connect(results.append)
        worker.statusChanged.connect(statuses.append)
        worker.finished.connect(lambda: finished.append(True))
        worker.run()

        self.assertEqual(results, [{"images": None, "polish": None}])
        self.assertEqual(statuses, ["Tidying"])
        self.assertEqual(finished, [True])

    def test_worker_emits_original_error_text(self):
        errors = []
        finished = []

        def fail(*args, **kwargs):
            raise RuntimeError("page conversion failed")

        worker = ConversionWorker(fail, "input.pdf", "output.docx", True, True)
        worker.failed.connect(errors.append)
        worker.finished.connect(lambda: finished.append(True))
        worker.run()

        self.assertEqual(errors, ["page conversion failed"])
        self.assertEqual(finished, [True])

    def test_window_finishes_conversion_from_background_thread(self):
        def convert(src, dst, snip, tidy, on_status):
            on_status("Processing")
            return {"images": None, "polish": None}

        window = ConverterWindow(convert)
        self.addCleanup(window.close)
        window.set_input("/tmp/paper.pdf")
        loop = QEventLoop()
        poll = QTimer()
        poll.timeout.connect(lambda: loop.quit() if window._thread is None else None)

        window.start_conversion()
        poll.start(10)
        QTimer.singleShot(3000, loop.quit)
        loop.exec()
        poll.stop()

        self.assertIsNone(window._thread)
        self.assertEqual(window.status_label.text(), "Done.")
        self.assertTrue(window.convert_button.isEnabled())


if __name__ == "__main__":
    unittest.main()
