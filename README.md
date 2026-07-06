# EasyExcel

EasyExcel は、デスクトップ版 Visual Studio Code で Excel 系ファイルをプレビュー・編集するための拡張機能です。

## 対応形式

- `.xlsx`
- `.xlsm`
- `.xls`
- `.csv`
- `.tsv`
- `.ods`

## 主な機能

- ワークブックとシートの表示
- セル編集
- 検索と置換
- 保存と名前を付けて保存
- 基本的なスタイル、結合セル、数式、ハイパーリンク、入力規則、保護情報、画像の読み書き

`.csv` と `.tsv` は単一シートのテキスト形式です。スタイル、結合セル、画像などのリッチな情報は保持できないため、必要に応じて `.xlsx` として保存してください。

## 開発

```bash
npm install
npm run build
```

開発時は次のコマンドで Vite と拡張ホストのビルドを起動します。

```bash
npm run dev
```

## Credits

本プロジェクトは [cweijan/vscode-office](https://github.com/cweijan/vscode-office) をベースに、Excel 系ファイルのプレビュー・編集機能に絞って再構成しています。元プロジェクトのライセンス表記は `LICENSE` に保持しています。
