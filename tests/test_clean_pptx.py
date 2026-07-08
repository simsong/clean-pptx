import unittest
from pathlib import Path

import clean_pptx


class TestCleanPptx(unittest.TestCase):
    def test_derive_output_paths(self):
        pptx_path = Path("/tmp/demo deck.pptx")
        cleaned_pptx, cleaned_pdf = clean_pptx.derive_output_paths(pptx_path)
        self.assertEqual(cleaned_pptx, Path("/tmp/demo deck-cleaned.pptx"))
        self.assertEqual(cleaned_pdf, Path("/tmp/demo deck-cleaned.pdf"))

    def test_parse_args_accepts_skip_pdf(self):
        args = clean_pptx.parse_args(["slides.pptx", "--skip-pdf"])
        self.assertEqual(args.pptx_file, Path("slides.pptx"))
        self.assertTrue(args.skip_pdf)

    def test_main_rejects_non_pptx_extension(self):
        rc = clean_pptx.main(["file.txt"])
        self.assertEqual(rc, 2)


if __name__ == "__main__":
    unittest.main()
