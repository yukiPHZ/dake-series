from pathlib import Path
import json
import hashlib
import shutil

ROOT=Path(__file__).resolve().parent
APP=ROOT/'01_apps/DAKE_PDF_Compress'
GATE=APP/'human_gate'
for sub in ('fixtures','outputs','visual'): (GATE/sub).mkdir(parents=True,exist_ok=True)
rows=json.loads((ROOT/'evidence/benchmark.json').read_text(encoding='utf-8'))
for row in rows:
    name=row['file']
    shutil.copy2(ROOT/'fixtures'/name,GATE/'fixtures'/name)
    if row['selected']:
        shutil.copy2(row['selected']['path'],GATE/'outputs'/(Path(name).stem+'_compressed.pdf'))
for p in (ROOT/'evidence/visual').glob('*.png'):
    if 'legacy' not in p.name:shutil.copy2(p,GATE/'visual'/p.name)
shutil.copy2(ROOT/'evidence/C_fax_candidate.pdf',GATE/'outputs/C_fax_candidate.pdf')
shutil.copy2(APP/'dist/DakePDF_Compress.exe',GATE/'DakePDF_Compress.exe')
shutil.copy2(APP/'dev/ADAPTIVE_BENCHMARK.md',GATE/'BENCHMARK.md')
manifest={str(p.relative_to(GATE)):{'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in GATE.rglob('*') if p.is_file()}
(GATE/'manifest.json').write_text(json.dumps(manifest,indent=2),encoding='utf-8')
(GATE/'README.txt').write_text('''DakePDF圧縮 / Adaptive Compression 人手確認用（未出荷）

DakePDF_Compress.exeを起動してください。設定はありません。
fixtures/ に個人情報を含まない入力PDF、outputs/ に自動選択した結果、visual/ に拡大比較PNGを入れています。

確認してほしいこと
1. 再起動後の起動が快適か（真のcold 3回とwarm 5回は人手測定待ち）。
2. fixturesのPDFをExplorerから直接ドロップして、すぐ反応するか。
3. PDFを確認中→圧縮できます、処理フェーズ→結果が自然に分かるか。
4. 完了ダイアログのOK後に保存先フォルダが開くか。
5. 日本語・5pt文字・細線・印影・頁順を元PDFと比較し、読めるか。
6. ご自身の実務PDFでも、画質を保って十分に軽くなるか。
7. 標準PDFビューア（Edge/Acrobat等）で正しく開けるか。
8. iLovePDFを開かなくてよいと感じるか。

注意
・今回の高い削減率は、未圧縮画像・繰り返しを含む試験データでの結果です。
・F_precompressed.pdfは圧縮済みで、出力を作らず案内します。
・グレー画像は品質保護のため非可逆再圧縮を行いません。
・iLovePDF比較結果は未提供です。アプリから外部サービスへ送信しません。
・正式Release/BOOTH/Web/Storeの更新は、この確認が終わった後です。

詳細: BENCHMARK.md。バイナリ照合: manifest.json。
''',encoding='utf-8')
print('Human gate package:',GATE)
print('exe sha256:',manifest['DakePDF_Compress.exe']['sha256'])
