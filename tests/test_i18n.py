from __future__ import annotations

import string
import unittest

from i18n import STRINGS, normalize_language, tr


class I18nTests(unittest.TestCase):
    def test_every_translation_has_matching_placeholders(self) -> None:
        formatter = string.Formatter()
        for key, translations in STRINGS.items():
            self.assertEqual(set(translations), {"zh", "en"}, key)
            placeholders = []
            for language in ("zh", "en"):
                fields = {
                    field_name
                    for _literal, field_name, _format_spec, _conversion in formatter.parse(
                        translations[language]
                    )
                    if field_name is not None
                }
                placeholders.append(fields)
            self.assertEqual(placeholders[0], placeholders[1], key)

    def test_language_normalization_and_lookup(self) -> None:
        self.assertEqual(normalize_language("zh-CN"), "zh")
        self.assertEqual(normalize_language("en_US"), "en")
        self.assertEqual(tr("en", "browse"), "Browse…")


if __name__ == "__main__":
    unittest.main()
