# 作業報告書管理システム

Google Apps Script（GAS）を使用した作業報告書管理システムです。

## 開発環境の構築

### 前提条件

- Node.js（v14以上推奨）
- npm
- Googleアカウント

### セットアップ手順

1. Google Apps Scriptの設定
   - [Google Apps Script 設定画面](https://script.google.com/home/usersettings)にアクセス
   - 「Google Apps Script API」を有効にする
   - 「clasp」の使用を許可する

2. claspのインストール
```bash
npm install -g @google/clasp
```

3. Googleアカウントでのログイン
```bash
npx clasp login
```
- ブラウザが開くので、Googleアカウントでログインし、必要な権限を付与してください。

4. プロジェクトの初期化
```bash
mkdir -p src
npm init -y
```

5. 依存関係のインストール
```bash
npm install @google/clasp --save-dev
```

6. プロジェクトの設定
- `.clasp.json`ファイルを作成し、以下の内容を設定：
```json
{
  "scriptId": "YOUR_SCRIPT_ID",
  "rootDir": "src"
}
```

7. プロジェクトのプル
```bash
npx clasp pull
```

### 開発コマンド

- ローカルの変更をGASにアップロード
```bash
npm run push
```

- GASの変更をローカルにダウンロード
```bash
npm run pull
```

- GASエディタを開く
```bash
npm run open
```

- 新しいデプロイメントを作成
```bash
npm run deploy
```

## プロジェクト構造

```
.
├── src/                    # ソースコードディレクトリ
│   ├── appsscript.json     # GAS設定ファイル
│   ├── Code.js            # メインスクリプト
│   ├── index.html         # フロントエンドHTML
│   ├── styles.css         # スタイルシート
│   └── script.js          # フロントエンドJavaScript
├── .clasp.json            # clasp設定ファイル
├── .gitignore             # Git除外ファイル
├── package.json           # npm設定ファイル
└── README.md              # このファイル
```

## トラブルシューティング

### アクセストークンのエラー

`Error retrieving access token: TypeError: Cannot read properties of undefined (reading 'access_token')`というエラーが発生した場合：

1. Google Apps Scriptの設定を確認
   - [Google Apps Script 設定画面](https://script.google.com/home/usersettings)で設定を確認
   - 「Google Apps Script API」が有効になっているか確認
   - 「clasp」の使用が許可されているか確認

2. claspのログイン状態を確認
```bash
npx clasp login --status
```

3. ログイン状態が異常な場合は、以下の手順で再ログイン
```bash
# 既存の認証情報をクリア
rm -rf ~/.clasprc.json

# 再度ログイン
npx clasp login
```

4. それでも解決しない場合は、以下の手順を試してください：
```bash
# claspを再インストール
npm uninstall -g @google/clasp
npm install -g @google/clasp

# プロジェクトディレクトリ内のnode_modulesを削除
rm -rf node_modules

# 依存関係を再インストール
npm install
```

### その他の一般的な問題

1. 権限エラーが発生する場合：
   - Googleアカウントに適切な権限が付与されているか確認
   - プロジェクトの共有設定を確認

2. スクリプトIDが無効な場合：
   - `.clasp.json`の`scriptId`が正しいか確認
   - GASエディタのURLから正しいIDを取得

## 注意事項

- `.clasp.json`ファイルには機密情報が含まれるため、Gitにコミットしないでください。
- 開発時は`src`ディレクトリ内のファイルを編集してください。
- デプロイ前に必ずテストを行ってください。 