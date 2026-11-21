from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import main


class SignatureExtractionTests(unittest.TestCase):
    def test_labeled_name_on_company_line(self) -> None:
        self.assertEqual(
            main.extract_signature_identity("株式会社JNT 営業担当：五十嵐"),
            ("株式会社JNT", "五十嵐"),
        )

    def test_unlabeled_name_on_company_line(self) -> None:
        self.assertEqual(
            main.extract_signature_identity("株式会社恵比寿キャピタル 山崎"),
            ("株式会社恵比寿キャピタル", "山崎"),
        )

    def test_sender_display_name_has_priority(self) -> None:
        self.assertEqual(
            main.extract_contact_identity(
                "<taro@example.jp>",
                "山崎 太郎 <taro@example.jp>",
                "taro@example.jp",
                "株式会社JNT\n営業担当：五十嵐",
                "",
            ),
            ("株式会社JNT", "山崎 太郎"),
        )

    def test_missing_display_and_signature_name_stays_empty(self) -> None:
        self.assertEqual(
            main.extract_contact_identity(
                "<taro@example.jp>",
                "<taro@example.jp>",
                "taro@example.jp",
                "お世話になっております。",
                "",
            ),
            ("", ""),
        )

    def test_short_body_sentence_is_not_treated_as_name(self) -> None:
        self.assertEqual(
            main.extract_signature_identity("株式会社JNT\n確認ください"),
            ("株式会社JNT", ""),
        )

    def test_current_sender_signature_wins_over_reply_history(self) -> None:
        body = """お世話になっております。

株式会社JNT 営業担当：五十嵐

-----Original Message-----
From: old@example.jp
Sent: Monday, November 16, 2026
To: taro@example.jp
Subject: 過去メール

株式会社過去商事 過去 花子
"""
        self.assertEqual(
            main.extract_signature_identity(body),
            ("株式会社JNT", "五十嵐"),
        )

    def test_reply_history_is_not_used_as_sender_signature(self) -> None:
        body = """承知しました。

差出人: old@example.jp
送信日時: 2026年11月16日
宛先: taro@example.jp
件名: 過去メール

株式会社過去商事 過去 花子
"""
        self.assertEqual(main.extract_signature_identity(body), ("", ""))

    def test_branch_name_remains_part_of_company(self) -> None:
        self.assertEqual(
            main.extract_signature_identity("株式会社さくら都市 千葉支店"),
            ("株式会社さくら都市 千葉支店", ""),
        )

    def test_department_and_role_are_not_in_person_name(self) -> None:
        body = "株式会社JNT 営業部\n部長 山崎 太郎"
        self.assertEqual(
            main.extract_signature_identity(body),
            ("株式会社JNT 営業部", "山崎 太郎"),
        )

    def test_html_table_keeps_lines_and_excludes_outlook_quote(self) -> None:
        html_body = b"""<html><body>
<table><tr><td>\xe6\xa0\xaa\xe5\xbc\x8f\xe4\xbc\x9a\xe7\xa4\xbeJNT</td></tr>
<tr><td>\xe9\x83\xa8\xe9\x95\xb7</td><td>\xe5\xb1\xb1\xe5\xb4\x8e \xe5\xa4\xaa\xe9\x83\x8e</td></tr></table>
<div id=\"divRplyFwdMsg\">From: old@example.jp<br>
\xe6\xa0\xaa\xe5\xbc\x8f\xe4\xbc\x9a\xe7\xa4\xbe\xe9\x81\x8e\xe5\x8e\xbb\xe5\x95\x86\xe4\xba\x8b \xe9\x81\x8e\xe5\x8e\xbb \xe8\x8a\xb1\xe5\xad\x90</div>
</body></html>"""
        text = main.body_to_text(html_body)
        self.assertNotIn("株式会社過去商事", text)
        self.assertEqual(
            main.extract_signature_identity(text),
            ("株式会社JNT", "山崎 太郎"),
        )


class CsvContractTests(unittest.TestCase):
    def test_csv_header_and_bom_remain_fixed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "mail_list.csv"
            main.write_csv(
                output_path,
                [main.ContactRow("株式会社JNT", "五十嵐", "info@example.jp")],
            )
            data = output_path.read_bytes()
        self.assertTrue(data.startswith(b"\xef\xbb\xbf"))
        self.assertEqual(
            data.decode("utf-8-sig").splitlines()[0],
            "会社名,お名前,メールアドレス",
        )


if __name__ == "__main__":
    unittest.main()
