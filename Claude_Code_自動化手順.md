# HMEP パイプライン — Claude Code による運用自動化 手順書

ハワイ大学沖縄事務所向け。`要件定義書.md`・`実装手順書.md`・`hmep_pipeline/README.md` と併用する。**Claude Code（Anthropic のターミナル上のエージェント CLI）** を使い、抽出・リネーム・アップロードの手順を繰り返し可能にし、人的ミスと説明コストを減らすことを目的とする。

---

## 1. この手順書でできること・できないこと

### できること

- **毎回同じ順序・同じコマンド**でパイプラインを実行させる（対話でも非対話でも）。
- **リポジトリ内のルール**（`config.yaml`、`main.py`、ログの見方）を Claude Code に常に読ませる（`CLAUDE.md` など）。
- **実行結果の要約・ログ確認・異常時の切り分け**をエージェントに任せる。
- **スクリプトの `--dry-run` や `--limit` を付けた安全運転**を指示に含めやすくする。

### できないこと（または人の介入が必須）

- **Outlook デスクトップが起動済み**でないとメール抽出（F2）は実行できない（要件定義の制約）。Claude Code が Outlook を代わりに起動・操作する保証はない。
- **Google OAuth の初回ブラウザ認証**やトークン無効時の再認証は、利用者のブラウザ操作が必要。
- **YouTube Data API の日次クォータ**超過時は、その日のうちに追加アップロードはできない。自動化しても API 上限は変わらない。

完全無人の「深夜だけで全部完了」は、上記の制約があるため現実的ではない。**「人が PC と Outlook を用意したうえで、Claude Code に一括オペを任せる」** 形が安全である。

---

## 2. 前提（環境）

| 項目 | 内容 |
|------|------|
| OS | Windows 10/11（本プロジェクト前提） |
| Python | 3.10 以上、`hmep_pipeline/README.md` のとおり venv と `requirements.txt` 済み |
| Outlook | Microsoft 365 **デスクトップ**版（抽出を行う日のみ起動） |
| 認証ファイル | `hmep_pipeline/credentials/client_secrets.json`・必要に応じて `token.json`（Git に含めない） |
| 公式ドキュメント | [Claude Code 概要](https://docs.anthropic.com/en/docs/claude-code)、[セットアップ](https://docs.anthropic.com/en/docs/claude-code/setup)、[CLI 利用](https://docs.anthropic.com/en/docs/claude-code/cli-usage) |

### 2.1 Claude Code のインストール（例）

PowerShell（管理者権限は公式手順に従う）:

```powershell
irm https://claude.ai/install.ps1 | iex
```

または次のいずれか（環境に合わせて選択）: `winget install Anthropic.ClaudeCode` など。詳細は公式の [Advanced setup](https://docs.anthropic.com/en/docs/claude-code/setup) を参照する。

### 2.2 認証（Claude Code 本体）

初回は `claude` のログインや API キー設定が必要になる。組織ポリシーに合わせ、社内手順がある場合はそちらを優先する。CI や長期トークンが必要な場合は公式の **`claude setup-token`** 等の説明を参照する。

---

## 3. 自動化のパターン（おすすめ順）

### パターン A：対話型（推奨・週次運用向け）

運用担当が PC を用意し、**作業ディレクトリを `D:\code4biz\HMEP`（実際のパスに置き換え）にした状態**で Claude Code を起動し、毎回ほぼ同じ **プロンプト雛形**を渡す。

1. Outlook を起動する（抽出する場合）。
2. ターミナルでリポジトリルートへ移動し、venv を有効化する。
3. `claude` を起動し、**「次のチェックリストとコマンドでパイプラインを実行し、ログを要約して」** と指示する（下記セクション 5 の雛形をコピー）。

**メリット:** OAuth やエラー時の追加質問に柔軟に対応できる。  
**デメリット:** 人手で Claude Code を起動する必要がある。

### パターン B：非対話モード `-p`（定型レポート・軽作業向け）

公式 CLI の **print / 非対話モード**（`claude -p "…"`）で、短い指示をスクリプトやタスクから渡す。長時間のパイプラインや Outlook 前提の作業より、**「ログの末尾を読んで要約して」「`--dry-run` の結果だけ報告」** などに向く。

```text
詳細は https://docs.anthropic.com/en/docs/claude-code/cli-usage の「Non-interactive mode」等を参照。
```

### パターン C：スケジューラは Python を直接叩く（アップロードのみ等）

** Windows タスク スケジューラ**で、Claude Code ではなく **`python hmep_pipeline\main.py upload --limit N`** を定期実行する。抽出（Outlook）を含むジョブは前述の通り Outlook 前提のため、無人夜間実行は非推奨。

クラウド CI での完全自動は、Outlook・OAuth・クォータの制約から本プロジェクトでは推奨しない。

---

## 4. 作業ディレクトリとコンテキストファイル（重要）

Claude Code に毎回同じ説明をさせないため、**リポジトリ直下に `CLAUDE.md`** を置くと、多くの場合エージェントがプロジェクトルールとして読む（公式のプロジェクトメモリの推奨事項に沿う）。**本リポジトリには `CLAUDE.md` を同梱済み**（必要に応じて追記する）。

### 4.1 `CLAUDE.md` に書いておくとよい内容（例）

- このリポジトリは **HMEP 動画パイプライン**であること。
- **作業ディレクトリは `hmep_pipeline` 基準**で `config.yaml` を読むこと。
- 実行の基本形:
  - `cd hmep_pipeline` のうえ `python main.py …`、または README の各ディレクトリからの個別スクリプト。
- **抽出前:** Outlook デスクトップ起動済み。
- **アップロード前:** `credentials` 配置・クォータ意識（`--limit`）。
- **ログ:** `logs/pipeline_YYYY-MM-DD.log` と各 `extract_` / `rename_` / `upload_` ログ。
- **編集禁止:** `credentials/*` のコミット、個人情報のログ貼り付け。

実体ファイルの追加は任意だが、自動化の再現性が上がる。

---

## 5. Claude Code 用・実行プロンプト雛形（コピー用）

以下をそのまま（パスだけ環境に合わせて）貼り付けて使う。

### 5.1 フルチェック（乾燥実行）

```text
リポジトリは HMEP の hmep_pipeline です。次を実行し、各ログの要点と注意点だけ日本語で要約してください。

1. 抽出は Outlook が必要なため、まず「Outlook を起動してから続行」と短く表示し、venv を有効化したうえで hmep_pipeline に cd。
2. python main.py extract --dry-run
3. python main.py rename --dry-run
4. python main.py upload --dry-run --limit 3

config.yaml と logs/ 内の当日ログを参照してよい。秘密情報は出力しない。
```

### 5.2 本番（段階的・例）

運用ポリシーに合わせて `--limit` を変える。

```text
Outlook は起動済み前提。venv 有効化後、hmep_pipeline で:

1. python main.py extract
2. python main.py rename
3. python main.py upload --limit 6

各ステップの終了コードと logs/pipeline_*.log の記録を確認し、失敗時はどのファイルを見るべきか指示してください。
```

### 5.3 抽出だけ再実行（台帳繰り上げ後）

```text
hmep_pipeline で python main.py extract --dry-run を実行し、台帳に取り込まれそうな新規行の件数を報告。問題なければ本番 extract のコマンド案内のみ（実行は指示後に確認）。
```

---

## 6. セキュリティ・運用上の禁止事項

- `credentials/` や `token.json` を **チャット・ログ・スクリーンショットに含めない**。
- Claude Code の会話に **個人情報・患者情報**を貼らない。ログに含まれる場合はマスクした抜粋のみ。
- **公開リポジトリ**に API キーや OAuth JSON をコミットしない（`.gitignore` を維持する）。

---

## 7. トラブルシュート（Claude Code 連携時）

| 現象 | 確認すること |
|------|----------------|
| 抽出が異常に遅い | `scan_from_last_lecture_date` と台帳の最終開催日、受信日の下限。初回は `false` や広い `received_date_from` も検討（README F2）。 |
| アップロード 403/クォータ | 当日の残クォータ。`--limit` を下げる、翌日再実行。 |
| 「config がない」 | カレントディレクトリが `hmep_pipeline` か、`--config` で `config.yaml` を明示。 |
| Claude Code が古い手順を参照 | `CLAUDE.md` または本手順書のパスをプロンプトに明示。 |

---

## 8. 次の文書との位置づけ

- **開発・実装フェーズ:** `実装手順書.md`（F0〜F7）。
- **本書:** Claude Code を使った**運用上の定型化・自動化のしかた**。
- **利用者向けの画面付きマニュアル:** 実装手順書 F7（引き渡し）で別紙化する想定。

---

*本手順書は運用フィードバックに応じて更新する。*
