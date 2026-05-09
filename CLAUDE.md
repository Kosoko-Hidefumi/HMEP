# HMEP — Claude Code 向けプロジェクトメモ

ハワイ大学沖縄事務所の **HMEP レクチャー動画管理パイプライン**。詳細は `要件定義書.md`・`実装手順書.md`・`hmep_pipeline/README.md`・`Claude_Code_自動化手順.md`。

## 何をするリポジトリか

1. **抽出（F2）** — Outlook → `hmep_pipeline/data/lectures.xlsx`
2. **リネーム（F4）** — 元動画をコピーし `hmep_pipeline/videos_renamed/` に命名規則どおり保存
3. **アップロード（F5）** — YouTube Data API v3・説明テンプレート・履歴 `data/upload_log.csv`

統合エントリは **`hmep_pipeline/main.py`**（`extract` / `rename` / `upload` / `all`）。

## 作業ディレクトリと設定

- **設定ファイル:** `hmep_pipeline/config.yaml`。**相対パスは `hmep_pipeline/` が基準**。
- コマンド例は **`hmep_pipeline` にカレントを合わせてから** 実行する。

```powershell
# リポジトリルート HMEP で venv 有効化後
Set-Location hmep_pipeline
python main.py --help
python main.py extract --dry-run
python main.py rename --dry-run
python main.py upload --dry-run --limit 3
python main.py all --dry-run
```

単体スクリプトでも同じ `--config` 既定（`hmep_pipeline/config.yaml`）。

## 人の前提（エージェントが代替できない）

- **抽出を実行するとき:** Microsoft Outlook **デスクトップ**を事前に起動しておく（要件定義・README）。
- **YouTube 初回・トークン失効時:** ブラウザ OAuth。`credentials/client_secrets.json` が必要。
- **クォータ:** アップロードは 1 日本数に上限。`python main.py upload --limit N` や `config.yaml` の `youtube.max_uploads_per_run` で抑える。

## ログ（トラブル時）

| 種別 | パス例 |
|------|--------|
| パイプライン境界 | `hmep_pipeline/logs/pipeline_YYYY-MM-DD.log` |
| 抽出 | `hmep_pipeline/logs/extract_YYYY-MM-DD.log` |
| リネーム | `hmep_pipeline/logs/rename_YYYY-MM-DD.log` |
| アップロード | `hmep_pipeline/logs/upload_YYYY-MM-DD.log` |

## 安全に実行するとき

- 初回や設定変更後は **`--dry-run`** で対象件数を確認。
- アップロードは **`--limit`** で小さく試す。再アップロード防止は `upload_log.csv` と台帳の `youtube_url` を参照（`--force` は重複リスク）。

## 触らない・出さないもの

- **`hmep_pipeline/credentials/`** の内容を Git にコミットしない。チャットやログに貼らない。
- 台帳・メール由来の **個人を特定しうる情報** は要約時も不必要に複製しない。

## 関連ファイル

- `hmep_pipeline/01_extract/extract_lectures.py`
- `hmep_pipeline/02_rename/rename_videos.py`
- `hmep_pipeline/03_upload/upload_youtube.py`
- `hmep_pipeline/03_upload/description_template.txt`
