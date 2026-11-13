from tools.generate_version_info import build_version_info


def test_build_version_info_uses_shared_windows_identity() -> None:
    text = build_version_info(
        {
            "display_name": "DakePDF結合",
            "site_title": "PDFを結合するだけ",
            "exe_name": "DakePDF_Merge.exe",
            "version": "1.2.3",
        }
    )

    assert "StringStruct('ProductName', 'DakePDF結合')" in text
    assert "StringStruct('FileDescription', 'PDFを結合するだけ')" in text
    assert "StringStruct('CompanyName', 'PEAKHEADZ')" in text
    assert "StringStruct('LegalCopyright', '© 2026 Yukihiko Kikuta')" in text
    assert "StringStruct('Authors', 'Yukihiko Kikuta')" in text
    assert "StringStruct('FileVersion', '1.2.3')" in text
    assert "StringStruct('ProductVersion', '1.2.3')" in text
    assert "StringStruct('OriginalFilename', 'DakePDF_Merge.exe')" in text
