# HP0524 プロジェクト

## 概要
住宅リフォーム会社「Sumai＋」のウェブサイト制作プロジェクト。

---

## 🌐 ブラウザで見る（GitHub Pages）

| ページ | URL |
|---|---|
| **Sumai＋ トップページ** | https://yutorin-ino.github.io/LEAGUE/sumai-plus/ |
| **リフォーム概算診断ツール** | https://yutorin-ino.github.io/LEAGUE/sumai-plus/estimator.html |
| **プロモーションLP** | https://yutorin-ino.github.io/LEAGUE/sumai-plus/promo.html |

---

## Sumai＋ ホームページ

### コンセプト
**「住まいのことで困ったら、まず相談できる存在へ。」**

- 住宅業界23年の代表が「住まいの総合相談窓口」として地域で選ばれるLPを目指す
- ターゲット：30〜60代の住宅所有者・リフォーム検討者（一宮市周辺）
- ブランドイメージ：相談しやすい・誠実・押し売りしない・地域密着

### ディレクトリ構成
```
sumai-plus/
├── index.html         # トップページ（メインHP）
├── estimator.html     # リフォーム30秒概算診断ツール
├── promo.html         # プロモーションLP
└── images/
    ├── family.png     # FVメインビジュアル（家族）
    ├── staff.png      # スタッフ・代表写真
    ├── living.png     # リビング施工事例
    ├── kitchen.png    # キッチン施工事例
    ├── bathroom.png   # 浴室・洗面施工事例
    ├── exterior.png   # 外壁・外構施工事例
    └── customer.png   # お客様の声アバター
```

### デザイン仕様
- カラー：ホワイトベース / ネイビー `#1a2d5a` / ベージュゴールド `#c8a96e`
- フォント：Noto Sans JP / Noto Serif JP（Google Fonts）
- レスポンシブ対応（スマホファースト）
- スマホ：画面下部に固定フローティングCTA（LINE / 無料相談 / 電話）

### ページ構成（index.html）
1. ファーストビュー（住宅業界23年・相談窓口コンセプト）
2. 実績数値バー（住宅業界23年・施工実績・顧客満足度・相談無料）
3. 選ばれる5つの理由
4. 代表の想い（ストーリーセクション）
5. 対応リフォーム一覧
6. 概算診断ツール誘導バナー
7. 施工事例 Before/After
8. 空気まで考える住まいづくり
9. お客様の声 ＋ Google口コミ導線
10. ご相談の流れ（6ステップ）
11. 代表メッセージ
12. お問い合わせ（LINE強化CTA）
13. フッター（一宮市周辺エリア明記）

### 対応エリア
一宮市・稲沢市・江南市・岩倉市・北名古屋市・清須市・岐阜市・羽島市・各務原市

---

## その他ファイル

| ファイル | 内容 |
|---|---|
| `mail_dashboard.html` | メール一括送信ダッシュボード |
| `generate_pdf.js` | PDF生成スクリプト（Node.js） |
| `sumai_report.html` | Sumai＋ レポート |
| `housing_whitepaper.*` | 住宅ホワイトペーパー |
