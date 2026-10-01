Word（その1・その2）の組み立て一式（2026/10/1）

作り直し方（このフォルダで実行）:
  python build_docx.py "..\..\【更新版】作成中【参考様式】事業計画書（その１その２）第8回フォーマット.docx" --template template.docx
- 章ごとの本文は chapters\row0.json〜row11.json（row0＝1．事業者の概要 … row11＝6．補足事項）
- （1）〜（4）の前段は front.json
- 文言の置換は fix_chapters.py（置換リストのJSONを渡す）、中身の一覧表示は dump_ch.py
- ページ数の目安は pagecheck.py（LibreOffice）。最終のページ数はWord本体で数えること
- 書き方の約束は SPEC.md
- 保存の前に、出力先のファイルをWordで開いていないことを確認すること
