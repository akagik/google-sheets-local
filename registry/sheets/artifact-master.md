# ArtifactMaster

## 概要

アーティファクトのマスターデータ。

## 最初に確認すること

新規にアーティファクトを追加する場合、区分列を確認して、今回のアーティファクトをどこに挿入するかを決定する. 特に rewardId (index) は他と被ってはいけないのと、自然に同じ区分の箇所に連続して配置したい（遠征機能関連・禁断の番人関連・・・など）

## 注意事項

- 区分ヘッダー行（例:「記念AF（6000-6999）」）がデータ行の間に存在する
  - 遠征アーティファクトは index 5000-5999 番台、rewardId は 301005xxx
- **名前** 列は実際のゲームでは取り込まれない。実際にはローカライズマスターのデータが使われる
- **target** と **targetIndex** は併用できない（両方値が入っていると動作しない）
- **isMultDisplay** / **isPerMille** は Deprecated。今後は **displayFormat** を使用する
- **iconParentFolder**: 特定バージョンで追加する場合は `x.x.x-yyyy` のように親フォルダを指定するのが推奨

## よく使う操作例

```bash
# バージョン 3.10.0 のアーティファクトを一覧取得
python .claude/skills/google-sheets-local/scripts/sheets_tool.py filter "https://docs.google.com/spreadsheets/d/1VPEQiB4vmU37wJaXncWgF9t0xS7sbtQHH8VxvwtQuow/edit?gid=1153062716#gid=1153062716" "ArtifactMaster" "バージョン" "3.10.0"

# ヘッダー一覧取得
python .claude/skills/google-sheets-local/scripts/sheets_tool.py headers "https://docs.google.com/spreadsheets/d/1VPEQiB4vmU37wJaXncWgF9t0xS7sbtQHH8VxvwtQuow/edit?gid=1153062716#gid=1153062716" "ArtifactMaster"
```