import importlib.util
import io
import pathlib
import tempfile
import unittest
import zipfile

ROOT = pathlib.Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("mobile_upload", ROOT / "mobile_upload.py")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


def workbook(marker="test"):
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("xl/workbook.xml", "<workbook>" + marker + "</workbook>")
    return stream.getvalue()


class MobileUploadTests(unittest.TestCase):
    def test_save_hash_and_retry(self):
        import hashlib
        data = workbook()
        with tempfile.TemporaryDirectory() as folder:
            name, digest, created = module.save_mobile_report(data, "사진대지.xlsx", folder, 100000)
            self.assertTrue(created)
            self.assertEqual(hashlib.sha256(data).hexdigest(), digest)
            self.assertEqual((pathlib.Path(folder) / name).read_bytes(), data)
            self.assertFalse(module.save_mobile_report(data, "사진대지.xlsx", folder, 100000)[2])
            self.assertEqual(len(list(pathlib.Path(folder).iterdir())), 1)

    def test_same_filename_different_content_is_not_lost(self):
        with tempfile.TemporaryDirectory() as folder:
            a = module.save_mobile_report(workbook("a"), "사진.xlsx", folder, 100000)
            b = module.save_mobile_report(workbook("b"), "사진.xlsx", folder, 100000)
            self.assertNotEqual(a[0], b[0])
            self.assertTrue(b[2])

    def test_capacity_and_invalid_file(self):
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaises(ValueError):
                module.save_mobile_report(workbook(), "a.xlsx", folder, 1)
            with self.assertRaises(ValueError):
                module.save_mobile_report(b"not excel", "a.xlsx", folder, 100000)
            with self.assertRaises(ValueError):
                module.save_mobile_report(workbook(), "a.exe", folder, 100000)
            self.assertEqual(list(pathlib.Path(folder).iterdir()), [])

    def test_filename_cannot_escape_store(self):
        with tempfile.TemporaryDirectory() as folder:
            name, _, _ = module.save_mobile_report(workbook(), "../escape.xlsx", folder, 100000)
            self.assertNotIn("/", name)
            self.assertTrue((pathlib.Path(folder) / name).is_file())

    def test_concurrent_retries(self):
        from concurrent.futures import ThreadPoolExecutor
        with tempfile.TemporaryDirectory() as folder:
            with ThreadPoolExecutor(max_workers=6) as pool:
                results = list(pool.map(lambda _: module.save_mobile_report(workbook(), "a.xlsx", folder, 100000), range(6)))
            self.assertEqual(sum(result[2] for result in results), 1)


if __name__ == "__main__":
    unittest.main()
