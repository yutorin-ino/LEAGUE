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
- **スクリプト**: Node.js のみ（`xlsx` パッケージ使用）。**Python は使用不可**（PATH未設定）
- **デザイン**: Noto Sans JP / Noto Serif JP、レスポンシブ（スマホファースト）
- **カラー**: ホワイト / ネイビー `#1a2d5a` / ベージュゴールド `#c8a96e`

---

## 海外ガジェットリサーチ方法論

### 目的
日本未上陸の海外便利グッズを発掘し、メーカーへパートナーシップ交渉メールを送る。  
**製品情報とメールアドレスを必ずセットで収集する**（メールなしではダッシュボードから送信できない）。

### 情報源（優先順位）
1. **CES**（Consumer Electronics Show・毎年1月）— Innovation Award受賞製品・新発表製品
2. **Kickstarter / Indiegogo** — クラウドファンディング中・達成済みの革新製品
3. **TIME誌 Best Inventions** — 年間ベスト発明品リスト
4. **Amazon US** — 海外で売れているが日本未展開の製品

### 検索キーワードパターン
```
"CES 2026" innovation award [カテゴリ名]
"best inventions 2025" [カテゴリ名]
"Kickstarter" funded [製品タイプ] 2025
```

カテゴリ別キーワード例：
- キッチン: `AI cooking robot` / `precision coffee maker` / `smart kitchen appliance`
- ウェアラブル: `smart ring health` / `sleep tracker wearable` / `EEG headband`
- ペット: `GPS pet tracker` / `smart pet collar AI` / `auto dog washer`
- 美容: `LED face mask clinical` / `laser hair regrowth home device`
- テクノロジー: `desktop metal 3D printer` / `AR glasses lightweight`

### 10カテゴリ
1. キッチン・調理器具
2. スマートホーム・インテリア・照明
3. ウェアラブル・ヘルス・フィットネス
4. アウトドア・スポーツ・旅行
5. ペット用品
6. テクノロジー・ガジェット
7. 美容・スキンケア
8. 子供・教育
9. ファッション・アクセサリー
10. クリーニング・収納・整理

### 製品選定基準
| 基準 | 内容 |
|---|---|
| 日本未上陸 | Amazon.co.jp・楽天で未販売であること |
| 革新性 | 既存製品にない機能・コンセプト |
| 日本市場ニーズ | 健康・時短・清潔・安全・ペット・美容への適合度 |
| 連絡先 | メーカーのメールアドレスが取得できること |

### 日本市場需要★評価の目安
- **★★★★★**: 日本の社会課題・トレンドに直結（睡眠・健康・ペット愛護・美容）
- **★★★★☆**: 日本市場に需要あり・競合少ない
- **★★★☆☆**: ニッチだが一定需要あり

### メールアドレスの探し方
1. 公式サイトの `Contact` / `Support` / `Business` ページ
2. よく使われるパターン: `info@` / `support@` / `hello@` / `business@` / `distribution@` / `wholesale@`
3. コンタクトフォームのみの場合: 連絡先欄に `https://…/contact` と記載してExcelに登録

### 1セットの目標件数
- 合計200件を目標に各カテゴリ均等に（10カテゴリ × 約20件）
- 不足しがちなカテゴリ: **ファッション・アクセサリー**、**子供・教育**（意識的に補充）

### スクリプト作成・実行フロー
```bash
# 1. 製品追加スクリプトを作成（例: add_200_items_6.js）
#    NEW_PRODUCTS配列に [番号, カテゴリ, 製品名, メーカー, ECサイト, 製品URL, メーカーHP, メール, ★, コメント] 形式で格納

# 2. 実行してExcelに追加
node add_200_items_6.js

# 3. ダッシュボード再生成
node generate_dashboard.js

# 4. GitHubにpush（GitHub Pages自動更新）
git add -A && git commit -m "Add XX products" && git push origin master
```

### 現在の収録状況
- **1683件収録**（マスター: `海外便利グッズリスト_日本未上陸1705件_評価付.xlsx`）
- メールアドレスあり: **1251件**
- 次回追加時は No.1684 から開始
