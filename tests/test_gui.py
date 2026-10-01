import unittest

import pdf2docx_gui


class ErrorDialogTests(unittest.TestCase):
    def test_queued_dialog_retains_error_message(self):
        callbacks = []
        shown = []

        class Root:
            def after(self, delay, callback):
                callbacks.append(callback)

        class Messagebox:
            @staticmethod
            def showerror(title, message):
                shown.append((title, message))

        pdf2docx_gui._queue_error_dialog(
            Root(), Messagebox, RuntimeError("conversion failed"))
        callbacks[0]()

        self.assertEqual(shown, [("Error", "conversion failed")])


if __name__ == '__main__':
    unittest.main()
