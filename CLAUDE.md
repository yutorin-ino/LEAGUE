# HP0524 プロジェクト

## プロジェクト概要

住宅リフォーム会社「Sumai＋」のウェブサイト制作と、海外クラウドファンディング商品リスト（1683件）管理を行うプロジェクト。

---

## GitHub Pages（重要・毎回確認不要）

> **GitHubリポジトリ名は `LEAGUE`（旧名 `HP0524`）。URLは必ず `/LEAGUE/` を使う。**

| 用途 | URL |
|---|---|
| トップ | `https://yutorin-ino.github.io/LEAGUE/` |
| ダッシュボード | `https://yutorin-ino.github.io/LEAGUE/mail_dashboard.html` |
| Sumai＋ | `https://yutorin-ino.github.io/LEAGUE/sumai-plus/` |

- `/HP0524/` は **404になる**ので絶対に使わない
- リポジトリURL: `https://github.com/yutorin-ino/LEAGUE`

---

## カスタムスラッシュコマンド

### `/summarize <ファイルパス>`

指定したファイルを読み込み、**タイトル・トピック3つ・結論**の形式で日本語要約を出力する。

```
/summarize README.md
/summarize sumai_report.html
/summarize generate_pdf.js
```

定義: `.claude/skills/summarize/SKILL.md`

---

## ディレクトリ構成

```
HP0524/
├── CLAUDE.md                        # Claude への指示（このファイル）
├── .claude/
│   ├── settings.json                # プロジェクト共有設定
│   ├── settings.local.json          # 個人設定（gitignore 推奨）
│   └── skills/
│       └── summarize/
│           └── SKILL.md             # /summarize コマンド定義
│
├── sumai-plus/                      # Sumai＋ ホームページ
│   ├── index.html                   # トップページ
│   ├── promo.html                   # プロモーション LP
│   └── images/
│
├── generate_pdf.js                  # PDF 生成スクリプト
├── generate_sumai_pdf.js
├── sumai_report.html
├── housing_whitepaper.html
├── housing_whitepaper.docx
├── mail_dashboard.html              # メール一括送信ダッシュボード
└── README.md
```

---

## 主要ファイルの役割

| ファイル | 役割 |
|---|---|
| `sumai-plus/index.html` | Sumai＋ メインサイト |
| `sumai-plus/promo.html` | プロモーション LP |
| `generate_pdf.js` | PDF 生成（Node.js） |
| `sumai_report.html` | Sumai＋ レポート |
| `housing_whitepaper.*` | 住宅ホワイトペーパー |
| `mail_dashboard.html` | メール一括送信 UI |

---

## 技術スタック

- **フロントエンド**: HTML / CSS / JavaScript（フレームワークなし）
- **スクリプト**: Node.js（PDF 生成）、Python（データ処理）
- **デザイン**: Noto Sans JP / Noto Serif JP、レスポンシブ（スマホファースト）
- **カラー**: ホワイト / ネイビー `#1a2d5a` / ベージュゴールド `#c8a96e`
