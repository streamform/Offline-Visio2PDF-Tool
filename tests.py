import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import visio2pdf


class CommandTests(unittest.TestCase):
    def test_existing_pdf_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "图纸.vsdx"
            source.touch()
            (Path(folder) / "图纸.pdf").touch()
            (Path(folder) / "图纸_1.pdf").touch()
            self.assertEqual(visio2pdf.output_path(source).name, "图纸_2.pdf")

    def test_only_supported_existing_files_are_accepted(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "图纸.VSDX"
            source.touch()
            self.assertEqual(visio2pdf.collect_inputs([str(source)]), [source.resolve()])
            with self.assertRaises(ValueError):
                visio2pdf.collect_inputs([str(Path(folder) / "说明.txt")])

    def test_failed_visio_falls_back_to_builtin_engine(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "图纸.vsd"
            source.touch()
            target = Path(folder) / "图纸.pdf"
            with patch.object(visio2pdf, "visio_export", side_effect=RuntimeError("未安装 Visio")), patch.object(visio2pdf, "portable_export", side_effect=lambda s, t, r: t.write_bytes(b"pdf")) as fallback:
                visio2pdf.convert(source, target, Path(folder))
                fallback.assert_called_once()

    def test_pause_keeps_dragged_conversion_window_open(self):
        with patch.object(visio2pdf.sys.stdin, "isatty", return_value=True), patch.object(visio2pdf, "collect_inputs", return_value=[]), patch.object(visio2pdf, "pause_when_interactive") as pause:
            visio2pdf.main(["file.vsd"])
            pause.assert_called_once()

    def test_runtime_package_layout_has_soffice_at_destination_root(self):
        with tempfile.TemporaryDirectory() as folder:
            destination = Path(folder) / "installed"
            with patch.object(visio2pdf, "runtime_root", return_value=destination), patch.object(visio2pdf.sys, "_MEIPASS", str(Path(visio2pdf.__file__).parent), create=True), patch.object(visio2pdf, "runtime_root", return_value=destination):
                soffice = visio2pdf.prepare_runtime(Path(folder))
                self.assertEqual(soffice, destination / "program" / "soffice.com")
                self.assertTrue(soffice.is_file())

    def test_pdf_validation_requires_a_page(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "empty.pdf"
            target.write_bytes(b"not a pdf")
            with self.assertRaises(ValueError):
                visio2pdf.validate_pdf(target)


if __name__ == "__main__":
    unittest.main()
