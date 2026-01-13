from __future__ import annotations

import unittest

import tools.corpus.triage as triage_mod


class TriageDiffIgnoreArgsTests(unittest.TestCase):
    def test_default_diff_ignore_is_enabled_by_default(self) -> None:
        parser = triage_mod._build_arg_parser()
        args = parser.parse_args(["--corpus-dir", "corpus", "--out-dir", "out"])

        self.assertFalse(args.no_default_diff_ignore)

        diff_ignore = triage_mod._compute_diff_ignore(
            diff_ignore=args.diff_ignore, use_default=not args.no_default_diff_ignore
        )
        self.assertEqual(diff_ignore, set(triage_mod.DEFAULT_DIFF_IGNORE))

    def test_no_default_diff_ignore_disables_builtins(self) -> None:
        parser = triage_mod._build_arg_parser()
        args = parser.parse_args(
            ["--corpus-dir", "corpus", "--out-dir", "out", "--no-default-diff-ignore"]
        )

        self.assertTrue(args.no_default_diff_ignore)

        diff_ignore = triage_mod._compute_diff_ignore(
            diff_ignore=args.diff_ignore, use_default=not args.no_default_diff_ignore
        )
        self.assertEqual(diff_ignore, set())

    def test_no_default_diff_ignore_still_allows_user_ignores(self) -> None:
        parser = triage_mod._build_arg_parser()
        args = parser.parse_args(
            [
                "--corpus-dir",
                "corpus",
                "--out-dir",
                "out",
                "--no-default-diff-ignore",
                "--diff-ignore",
                "/xl/workbook.xml",
            ]
        )

        diff_ignore = triage_mod._compute_diff_ignore(
            diff_ignore=args.diff_ignore, use_default=not args.no_default_diff_ignore
        )
        self.assertEqual(diff_ignore, {"xl/workbook.xml"})


if __name__ == "__main__":
    unittest.main()

