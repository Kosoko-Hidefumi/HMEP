# HMEP レクチャー動画管理パイプライン（開発用）

`要件定義書.md` / `実装手順書.md` に基づく実装置き場です。**運用マニュアルは F7 で別途整備**します。

## 前提

- Windows 10/11、Python 3.10 以上
- ① メール抽出を行う場合: **Microsoft Outlook デスクトップ**が利用でき、`win32com` 実行時は **Outlook が起動済み**であること（要件定義 6）
- ③ YouTube アップロードを行う場合: Google Cloud で **YouTube Data API v3** を有効化し、`credentials/client_secrets.json` を配置（本リポジトリにはコミットしない）

## セットアップ（リポジトリルート `HMEP/` で実行）

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -c "import yaml, openpyxl; print('ok')"
```

`requirements.txt` は **ASCII のみ**にしてある。日本語コメントのみの `requirements.txt` は、環境によって `pip install` の文字コードエラーになることがある。

## 構成

- `config.yaml` — パス・Outlook フォルダ名など（**相対パスは `hmep_pipeline/` を基準**とする想定で後続スクリプトで読み込む）
- `01_extract/extract_lectures.py` — ① Outlook → `data/lectures.xlsx`（F2）
- `02_rename/rename_videos.py` — ② 動画のコピー＋リネーム → `videos_renamed/`（F4）
- `03_upload/upload_youtube.py` — ③ YouTube アップロード（F5）
- `data/` — `lectures.xlsx` 等の生成先
- `credentials/` — OAuth 秘密情報（`.gitignore` 対象）

## F2 — メール抽出（extract_lectures.py）

1. Outlook で対象メールを **`outlook.folder_name`** に置く。空なら **受信トレイ**を走査。**Outlook 起動済み**で実行。
2. **`outlook.sender_email`** / **`outlook.subject_must_contain`** が設定されている場合、その送信元・件名キーワードで絞り込み（いずれも空ならフィルタなし）。
3. **`outlook.received_date_from` / `received_date_to`** … メールを **受信日** で走査する範囲。2025年1月開催でも案内が2024年に届く場合があるため、`lecture_date_*` より広く取る。
4. **`outlook.lecture_date_from` / `lecture_date_to`** … 台帳に載せる **開催日** の範囲。
5. 重複行は **同一開催日 + 正規化した講師名 + タイトル** で1行にまとめ、**最新の受信日時**のメールを採用（リマインド複数通に対応）。
6. 本文は `【講義紹介文】` / `【講義紹介】`、`【講師略歴】` / `【講師紹介】` などの表記ゆれに対応。略歴が無い場合は `parse.allow_empty_biography` でプレースホルダ可。1通に複数講義がある場合は `parse.split_multiple_sessions` で分割。

```powershell
Set-Location hmep_pipeline\01_extract
python extract_lectures.py --dry-run
python extract_lectures.py
```

Outlook なしでパースだけ試す（UTF-8 コンソール推奨: `chcp 65001` または Windows Terminal）:

```powershell
python extract_lectures.py --fixture fixtures\sample_ok.json
```

- 出力: `data/lectures.xlsx`（シート `lectures` / `failed`）。保存のたびに **講義キーで重複除去**。
- ログ: `logs/extract_YYYY-MM-DD.log`

## F4 — 動画リネーム（rename_videos.py）

1. **`paths.videos_dir`** に元動画のルートを設定する（既定: `../2025` = リポジトリ直下の `2025/` 再帰走査）。容量が大きい場合は `rename.recursive: false` にして直下のみにすること。
2. **`rename.date_from` / `date_to`** の期間内のレクチャー行だけを処理（既定は `outlook` と同じ 2025-01～2026-05）。
3. マッチ: 動画の更新時刻がレクチャー日時の **± `time_window_hours`**（既定 24）以内かつ、ファイル名と **講師名＋タイトル** の類似度が **`similarity_threshold`** 以上。
4. 出力ファイル名（タイトルは「」で囲む。`"` は Windows 禁止のため使わない）: `YYYY M D 氏名 先生 「タイトル」.mp4`
5. `lectures.xlsx` に `source_video_file` / `renamed_video_file` / `rename_match_score` / `rename_status` を追記。

```powershell
Set-Location hmep_pipeline\02_rename
python rename_videos.py --dry-run
python rename_videos.py
```

- ログ: `logs/rename_YYYY-MM-DD.log`

## F5 — YouTube アップロード（upload_youtube.py）

1. Google Cloud で **YouTube Data API v3** を有効化し、OAuth クライアント（デスクトップアプリ）の JSON を **`credentials/client_secrets.json`** に保存する。
2. `config.yaml` の **`youtube.playlist_id`** に、登録先プレイリスト ID を入れる（空のときはアップロードのみでプレイリスト登録はスキップ）。
3. **1 日のクォータ**（アップロードはユニット消費が大きい）を避けるため、`youtube.max_uploads_per_run`（既定 6）または **`python upload_youtube.py --limit 3`** で件数を切る。

```powershell
Set-Location hmep_pipeline\03_upload
python upload_youtube.py --dry-run
python upload_youtube.py
```

- 入力: `paths.videos_renamed_dir`（既定 `videos_renamed/`）のファイル名が、台帳の `renamed_video_file` と一致する行。開催日は `rename.date_from`～`date_to`（なければ `outlook.lecture_date_*`）。
- 出力: `data/upload_log.csv`（成功・失敗の履歴）、`lectures.xlsx` に `youtube_video_id` / `youtube_url` / `upload_status` / `uploaded_at`。
- `upload_log.csv` に **`status=ok`** の `local_file` がある場合、**`--force` がない限り再アップロードしません**（重複防止）。
- ログ: `logs/upload_YYYY-MM-DD.log`

## 次のステップ

実装手順書 **F6**（`main.py` で extract / rename / upload を統合）。
