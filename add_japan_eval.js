const XLSX = require('xlsx');

// ===================== 評価ロジック =====================

const CATEGORY_BASE = {
  "キッチン・調理器具":           { base: 3, high: "日本の料理好き層・ガジェット好き層に高い訴求力。高品質・デザイン性で差別化可能。", mid: "実用的な製品だが類似品が国内に存在。独自機能での訴求が必要。", low: "国内競合が強く、海外ブランドの優位性が薄い。価格競争になりやすい。" },
  "スマートホーム・インテリア・照明": { base: 3, high: "IoT・スマートホーム市場は急成長中。Matter規格対応製品は特に有利。", mid: "スマートホーム認知は広まっているが、設置ハードルや日本語対応が課題。", low: "日本の住宅事情（賃貸・スペース制限）や規格問題で普及に障壁がある。" },
  "ウェアラブル・ヘルス・フィットネス": { base: 4, high: "健康意識の高い日本市場で非常に高い需要。予防医療トレンドと完全合致。", mid: "健康管理ニーズは高いが、競合製品（Garmin・Apple Watch等）との差別化が必要。", low: "日本での競合製品が充実しており、独自ポジションの確立が課題。" },
  "アウトドア・スポーツ・旅行":   { base: 3, high: "アウトドアブーム継続中。軽量・コンパクト・機能性を重視する日本市場に最適。", mid: "アウトドア需要は堅調だが、国内外の競合ブランドが強い分野。", low: "日本のアウトドア文化・地形と合わない可能性あり。ニッチな需要にとどまる。" },
  "ペット用品":                  { base: 4, high: "ペット人口増加・高齢化ペット・ペットのプレミアム化トレンドと完全合致。", mid: "ペット市場は拡大中だが、日本独自ブランドとの競争が激しい分野。", low: "日本のペット文化（室内飼い中心・小型犬猫）と製品特性のミスマッチあり。" },
  "テクノロジー・ガジェット":     { base: 4, high: "ガジェット感度の高い日本市場で高い訴求力。AIやクリエイター向け製品は特需あり。", mid: "技術的差異は明確だが、日本語対応・サポート体制の整備が普及の鍵。", low: "日本メーカーの競合製品が強く、海外ブランドの優位性を出しにくい分野。" },
  "美容・スキンケア":            { base: 4, high: "世界最高水準の美容意識を持つ日本市場。美容機器・高機能スキンケアへの需要は旺盛。", mid: "美容市場は大きいが、日本ブランドへの信頼が高く、海外ブランドの浸透に時間を要する。", low: "成分規制・薬事法対応が必要な場合あり。参入ハードルが高い分野。" },
  "子供・教育":                  { base: 3, high: "STEM教育・プログラミング教育への注目が高まり需要拡大中。教育投資意欲が高い市場。", mid: "教育系製品は親の購買意欲は高いが、カリキュラムとの連動や日本語化が課題。", low: "国内の教育玩具ブランドが充実しており、差別化ポイントが限られる。" },
  "ファッション・アクセサリー":   { base: 3, high: "デザイン性・機能性を兼ね備えたアイテムは日本のファッション感度層に高い訴求力。", mid: "ファッション市場は大きいが、トレンド感度が高く参入タイミングが重要。", low: "国内外の強力ブランドが多く、認知獲得に多大なマーケティング投資が必要。" },
  "クリーニング・収納・整理":     { base: 4, high: "清潔・整理整頓を重視する日本文化と完全合致。スマート家電・省スペース設計は特に有利。", mid: "掃除・収納用品へのニーズは高いが、価格帯と機能のバランスが重要。", low: "国内家電・収納ブランド（パナソニック・山崎実業等）との競合が激しい。" },
};

const PRODUCT_OVERRIDES = {
  "oura ring": 5, "whoop": 5, "ultrahuman ring": 5, "ringconn": 5,
  "plaud note": 5, "hoverair": 5, "eight sleep": 5, "lumen": 5,
  "dexcom": 5, "apollo neuro": 5, "higherdose": 5, "omnilux": 5,
  "lyma laser": 5, "currentbody": 5, "jovs": 5, "therabody": 5,
  "remarkable": 5, "flipper zero": 5, "bambu lab": 5, "prusa": 5,
  "framework": 5, "hatch restore": 5, "manta sleep": 5, "sensate": 5,
  "elemind": 5, "kokoon": 5, "litter-robot": 5, "enabot": 5,
  "furbo": 5, "petcube": 5, "fluentpet": 5, "embark": 5,
  "muse s": 5, "flowtime": 5, "dreem": 5,
  "bartesian": 5, "spinn coffee": 5, "chef iq": 5,
  "typhur": 5, "combustion inc": 5,
  "solawave": 5, "nuface": 5, "foreo": 5, "nira": 5,
  "narwal": 5, "dreame": 5, "roborock": 5,
  "brilliant": 5, "halo rise": 5,
  "kiwico": 5, "osmo": 5, "sphero": 5,
  "bedjet": 4, "airthings": 4, "flume": 4, "graywind": 4,
  "brisk it": 4, "solo stove": 4, "brava": 4, "tovala": 4,
  "fellow": 4, "caraway": 4, "our place": 4,
  "nanoleaf": 4, "level lock": 4, "flair smart": 4, "emporia": 4,
  "withings": 4, "polar": 4, "biostrap": 4,
  "tractive": 4, "wagz": 4, "halo collar": 4, "spoton": 4,
  "helinox": 4, "biolite": 4, "rumpl": 4, "mpowerd": 4,
  "xreal": 4, "rabbit r1": 4,
  "toniebox": 4, "wonder workshop": 4, "kano": 4,
  "peak design": 4, "bellroy": 4, "nomatic": 4, "tropicfeel": 4,
  "hedgehog dryer": 4, "townew": 4, "simplehuman": 4,
  "lomi": 4, "fryaway": 4,
  "elegoo": 4, "anycubic": 4, "xtool": 4,
  "august wi-fi": 4, "schlage": 4, "wyze": 4,
  "coway": 4, "iqair": 4, "blueair": 4, "molekule": 4,
  "eight sleep": 5, "levels health": 5,
};

const SCORE_UP_WORDS = ["ai", "smart", "wireless", "gps", "health", "sleep", "ring", "infrared", "laser",
  "led", "therapy", "monitor", "sensor", "precision", "premium", "pro", "ultra",
  "robot", "automatic", "electric", "biometric", "app-enabled", "app enabled"];

const COMMENT_EXTRA = {
  "ring": "スマートリング市場は日本でも急成長中。ファッション感覚で使える健康管理デバイスとして高い支持を期待。",
  "sleep": "睡眠の質への関心が急上昇する日本市場でタイムリーな訴求が可能。",
  "ai": "AI搭載製品への注目度が高く、先進技術訴求で高付加価値ポジションを確立できる。",
  "laser": "美容レーザー機器は美容大国・日本での需要が特に高い。サロン代替として家庭用需要も拡大。",
  "led": "光療法・LED美容機器はSNS拡散で急速に認知獲得している分野。",
  "robot": "ロボット掃除機・ペットロボット等、日本市場でのロボット受容度は非常に高い。",
  "pizza": "ピザ文化が定着する日本で、本格ピザ焼き機は食のこだわり層への訴求に有効。",
  "water": "安全な水へのこだわりが強い日本市場。高機能浄水製品は信頼獲得が早い。",
  "gps": "安心・安全への意識が高い日本でGPS追跡デバイスへの需要は堅調。",
  "coffee": "コーヒー文化が成熟した日本市場。高品質抽出機器への投資意欲が高い層が存在。",
  "camping": "キャンプブーム継続中の日本で高い需要。軽量・コンパクト設計が特に評価される。",
  "pet": "ペットを家族として扱う文化が定着。プレミアム製品への支出を惜しまない層が拡大。",
  "coding": "小学校プログラミング教育必修化を背景に、STEM教育製品への需要が急増。",
  "backpack": "アウトドア・通勤・旅行すべてに対応できる多機能バッグへの日本市場ニーズは高い。",
  "purifier": "花粉症・PM2.5対策として空気清浄機は日本で定番化。高性能モデルへの需要が根強い。",
  "vacuum": "清潔志向の日本市場でロボット掃除機・コードレス掃除機の普及率は世界最高水準。",
  "thermometer": "料理精度へのこだわりが強い日本市場。ワイヤレス温度計は料理愛好家層へ高い訴求力。",
  "solar": "再エネ関心の高まりとともにポータブル太陽光製品の認知が向上中。",
  "lock": "セキュリティ意識向上でスマートロックへの関心は高まっているが、ドア規格への対応が課題。",
  "blender": "スムージー・健康飲料ブームが継続。高性能ブレンダーへの需要は底堅い。",
  "drone": "空撮ニーズは高いが規制対応（航空法・許可申請）が参入の鍵となる。",
};

const STAR_LABELS = { 5: "★★★★★", 4: "★★★★☆", 3: "★★★☆☆", 2: "★★☆☆☆", 1: "★☆☆☆☆" };

function calcScore(productName, makerName, category) {
  const text = (productName + " " + makerName).toLowerCase();
  const cat = CATEGORY_BASE[category] || { base: 3, high: "日本市場で一定の需要が期待できる。", mid: "市場調査が必要。", low: "参入ハードルが高い。" };
  let score = cat.base;

  // 個別オーバーライド
  for (const [keyword, forced] of Object.entries(PRODUCT_OVERRIDES)) {
    if (text.includes(keyword)) {
      score = forced;
      break;
    }
  }

  // キーワード補正（オーバーライドなかった場合のみ）
  if (!Object.keys(PRODUCT_OVERRIDES).some(k => text.includes(k))) {
    if (SCORE_UP_WORDS.some(w => text.includes(w))) {
      score = Math.min(5, score + 1);
    }
  }

  score = Math.max(1, Math.min(5, score));

  // コメント
  let base;
  if (score >= 4) base = cat.high;
  else if (score === 3) base = cat.mid;
  else base = cat.low;

  // 追加コメント
  let extra = "";
  for (const [keyword, detail] of Object.entries(COMMENT_EXTRA)) {
    if (text.includes(keyword)) {
      extra = " " + detail;
      break;
    }
  }

  return { score, comment: base + extra };
}

// ===================== Excel処理 =====================
const INPUT  = "C:\\Users\\hiro\\HP0524\\海外便利グッズリスト_日本未上陸1505件.xlsx";
const OUTPUT = "C:\\Users\\hiro\\HP0524\\海外便利グッズリスト_日本未上陸1505件_評価付.xlsx";

console.log("Excelファイル読み込み中...");
const wb = XLSX.readFile(INPUT);
const ws = wb.Sheets[wb.SheetNames[0]];

// ヘッダー追加
const range = XLSX.utils.decode_range(ws['!ref']);

// I1, J1 ヘッダー
ws['I1'] = { v: '日本市場需要評価', t: 's' };
ws['J1'] = { v: '需要マッチコメント', t: 's' };

const scoreDist = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
let processed = 0;

for (let row = 1; row <= range.e.r; row++) {  // row 0 = header
  const noCell   = ws[XLSX.utils.encode_cell({ r: row, c: 0 })];
  if (!noCell || !noCell.v) continue;

  const productName = (ws[XLSX.utils.encode_cell({ r: row, c: 2 })] || {}).v || "";
  const makerName   = (ws[XLSX.utils.encode_cell({ r: row, c: 3 })] || {}).v || "";
  const category    = (ws[XLSX.utils.encode_cell({ r: row, c: 1 })] || {}).v || "";

  const { score, comment } = calcScore(productName, makerName, category);
  scoreDist[score]++;

  ws[XLSX.utils.encode_cell({ r: row, c: 8 })] = { v: STAR_LABELS[score], t: 's' };
  ws[XLSX.utils.encode_cell({ r: row, c: 9 })] = { v: comment, t: 's' };

  processed++;
  if (processed % 200 === 0) console.log(`  処理中: ${processed}件`);
}

// 範囲更新
ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: range.e.r, c: 9 } });

// 列幅設定
ws['!cols'] = ws['!cols'] || [];
ws['!cols'][8] = { wch: 14 };
ws['!cols'][9] = { wch: 60 };

XLSX.writeFile(wb, OUTPUT);

console.log(`\n完了: ${OUTPUT}`);
console.log(`処理件数: ${processed} 件`);
console.log('\n評価分布:');
for (let s = 5; s >= 1; s--) {
  const bar = '■'.repeat(Math.round(scoreDist[s] / 5));
  console.log(`  ${STAR_LABELS[s]}: ${String(scoreDist[s]).padStart(4)}件  ${bar}`);
}
