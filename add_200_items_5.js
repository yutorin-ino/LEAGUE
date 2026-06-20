/**
 * add_200_items_5.js — No.1506〜1705 (200件追加)
 * node add_200_items_5.js
 */

const XLSX = require('xlsx');
const path = require('path');

const INPUT_FILE  = path.join(__dirname, '海外便利グッズリスト_日本未上陸1505件_評価付.xlsx');
const OUTPUT_FILE = path.join(__dirname, '海外便利グッズリスト_日本未上陸1705件_評価付.xlsx');
const SENDER      = 'yutorin.ino@gmail.com';

function buildGmailUrl(email, maker, product) {
  const subject = 'Potential Distribution Partnership for Japan';
  const body = `Dear ${maker} Team,

My name is Hiroyuki Inoguchi from
Sumai pluS Co., Ltd. in Japan

I hope this message finds you all well.

Our company focuses on promoting products that enrich people's daily lives. We are currently seeking unique international brands to introduce to the Japanese market and support their growth.

I recently had the opportunity to review your product (${product}) and was very impressed by its potential in the Japanese market.

I am confident that your product will strongly appeal to Japanese customers, and I would like to explore the possibility of building a partnership with your company to help ensure its success.

Furthermore, our company has a proven track record of promoting overseas products through our proprietary sales network and Japanese distribution channels.

Would it be possible to schedule a brief online meeting sometime next week to discuss this matter?

I look forward to hearing from you.


Sincerely,

Hiroyuki Inoguchi
Sumai pluS Co., Ltd.
 (LEAGUE Co., Ltd. Agent )

================================================

Sumai pluS Co., Ltd.
 (LEAGUE Co., Ltd. Agent )
Hiroyuki Inoguchi

E-mail:yutorin.ino@gmail.com

Address: 49-5 Kitazakuno, Higashi-Itsushiro, Ichinomiya City, Aichi Prefecture, Japan

=================================================`;
  const enc = (s) => encodeURIComponent(s);
  return `https://mail.google.com/mail/?authuser=${enc(SENDER)}&view=cm&fs=1&to=${enc(email)}&su=${enc(subject)}&body=${enc(body)}`;
}

// 200件の新製品データ
const NEW_PRODUCTS = [
  // --- キッチン・調理器具 ---
  [1506,"キッチン・調理器具","Liffo AI Cooking Robot","Liffo (Robomagister)","Kickstarter","https://www.liffo.com/en/product/liffo-robot/","https://www.liffo.com/","support@liffo.com","★★★★★","AI搭載で完全自律調理を実現したロボット。イタリア発でUS展開中。日本の料理好き・ガジェット好き層への訴求力は非常に高い。"],
  [1507,"キッチン・調理器具","Foodease All-in-One Smart Cooker","Foodease","Kickstarter/Indiegogo","https://www.foodease.com/","https://www.foodease.com/","support@foodease.ai","★★★★☆","全自動スマート調理器。多機能・レシピ連動型で料理自動化を実現。日本の時短ニーズに合致。"],
  [1508,"キッチン・調理器具","Solidteknics Polished Wrought Iron Pan","Solidteknics","Amazon AU","https://www.solidteknics.com/cookware","https://www.solidteknics.com/","info@solidteknics.com","★★★★☆","オーストラリア製鍛造鉄フライパン。無塗装・長寿命・高品質でこだわり料理人層に強い訴求力。"],
  [1509,"キッチン・調理器具","CookingPal Multo Smart Kitchen Hub","CookingPal","CookingPal公式","https://www.cookingpal.com/pages/meet-multo-v2","https://www.cookingpal.com/","support@cookingpal.com","★★★★☆","15以上の調理機能を搭載したスマートキッチンシステム。ガイドレシピで初心者も本格料理が可能。"],
  [1510,"キッチン・調理器具","Cometeer Flash-Frozen Coffee Capsules","Cometeer","Cometeer公式","https://cometeer.com/","https://cometeer.com/","support@cometeer.com","★★★★☆","スペシャルティコーヒーを瞬間冷凍カプセル化。お湯を注ぐだけで本格コーヒーが楽しめる革新的製品。"],
  [1511,"キッチン・調理器具","Ember Mug 2 Temperature Control Mug","Ember","Amazon US","https://ember.com/products/ember-mug-2","https://ember.com/","contact@ember.com","★★★★☆","温度管理スマートマグ。スマホアプリで好みの温度に保温。コーヒー文化が発展した日本市場に好適。"],
  [1512,"キッチン・調理器具","Ember Travel Mug 2+ Smart Mug","Ember","Ember公式","https://ember.com/products/ember-travel-mug-2-plus","https://ember.com/","contact@ember.com","★★★★☆","外出先でもアプリ管理で温度を維持できるスマート保温タンブラー。通勤・アウトドア需要に合致。"],
  [1513,"キッチン・調理器具","Ember Tumbler 16oz Smart Cup","Ember","Ember公式","https://ember.com/products/ember-tumbler","https://ember.com/","contact@ember.com","★★★★☆","コールドドリンク対応の温度管理スマートタンブラー。ホット・コールド両対応で汎用性が高い。"],
  [1597,"キッチン・調理器具","Typhur InstaProbe Wireless Thermometer","Typhur","Amazon US","https://www.typhur.com/","https://www.typhur.com/","support@typhur.com","★★★★★","0.5秒で検温できる超高速精度ワイヤレス肉温度計。料理精度にこだわる日本市場でロングセラーが見込まれる。"],
  [1598,"キッチン・調理器具","Typhur Dome Plus Air Fryer","Typhur","Amazon US","https://www.typhur.com/","https://www.typhur.com/","support@typhur.com","★★★★★","透明ドーム型の高性能スマートエアフライヤー。リアルタイム映像確認・高精度温度管理で料理の失敗ゼロ。"],
  [1599,"キッチン・調理器具","Combustion Inc Predictive Thermometer","Combustion Inc","Combustion公式","https://combustion.inc/","https://combustion.inc/","support@combustion.inc","★★★★★","8点計測の予測型スマートミート温度計。完成時刻を事前に通知するBluetooth対応の革新的調理ツール。"],
  [1600,"キッチン・調理器具","Fellow Aiden Precision Coffee Maker","Fellow","Amazon US","https://fellowproducts.com/","https://fellowproducts.com/","hello@fellowproducts.com","★★★★★","スマフォアプリ制御の精密コーヒーメーカー。温度・タイミング・プリウェット設定でスペシャルティコーヒーを自宅で再現。"],
  [1601,"キッチン・調理器具","Fellow Opus Conical Burr Coffee Grinder","Fellow","Amazon US","https://fellowproducts.com/","https://fellowproducts.com/","hello@fellowproducts.com","★★★★★","コニカルバーグラインダー。均一な粒度でエスプレッソ〜コールドブリューまで対応。コーヒー文化が深い日本向け。"],
  [1602,"キッチン・調理器具","Bartesian Premium Cocktail Maker","Bartesian","Bartesian公式","https://www.bartesian.com/","https://www.bartesian.com/","support@bartesian.com","★★★★★","カプセル式のプレミアムカクテルメーカー。自宅でバーテンダー品質のカクテルが楽しめる。おうち時間需要に合致。"],
  [1670,"キッチン・調理器具","Caraway Home Cookware Set (Ceramic)","Caraway Home","Amazon US","https://www.carawayhome.com/","https://www.carawayhome.com/","hello@carawayhome.com","★★★★★","PTFE・PFOA不使用のセラミックコーティングクックウェアセット。健康志向で安全な調理器具を求める日本市場に最適。"],
  [1671,"キッチン・調理器具","Our Place Always Pan 2.0","Our Place","Amazon US","https://fromourplace.com/","https://fromourplace.com/","hello@fromourplace.com","★★★★★","フライパン・鍋・蒸し器など8機能を一台に統合したオールインワンパン。少ないキッチン道具でシンプル調理を求める日本人向け。"],
  [1672,"キッチン・調理器具","Our Place Perfect Pot Cast Iron","Our Place","Amazon US","https://fromourplace.com/","https://fromourplace.com/","hello@fromourplace.com","★★★★★","セラミックコーティングの鋳鉄製多機能鍋。炒め・煮込み・オーブン調理まで対応する一生モノの調理器具。"],
  [1673,"キッチン・調理器具","Brava Home Smart Oven","Brava","Amazon US","https://www.brava.com/","https://www.brava.com/","support@brava.com","★★★★★","カメラ内蔵・光加熱で食材を自動認識して最適調理するスマートオーブン。AIシェフ機能で料理の失敗ゼロを目指す製品。"],
  [1674,"キッチン・調理器具","Tovala Smart Steam Oven","Tovala","Amazon US","https://tovala.com/","https://tovala.com/","support@tovala.com","★★★★★","QRコードスキャンでレシピ設定を自動化するスマートスチームオーブン。食事サブスクとの連携で調理の手間を大幅削減。"],
  [1675,"キッチン・調理器具","Brisk It Origin-580 Smart Pellet Grill","Brisk It","Amazon US","https://www.briskitgrill.com/","https://www.briskitgrill.com/","support@briskitgrill.com","★★★★☆","AI搭載のスマートペレットグリル。Alexaと連携して温度管理・調理時間を自動最適化。アウトドアBBQ需要に合致。"],
  [1676,"キッチン・調理器具","Solo Stove Pi Prime Pizza Oven","Solo Stove","Amazon US","https://www.solostove.com/","https://www.solostove.com/","support@solostove.com","★★★★★","家庭用・屋外用の本格ピザオーブン。本格ナポリピザ文化が定着した日本で食のこだわり層への訴求に有効。"],
  [1692,"キッチン・調理器具","Chef iQ Smart Pressure Cooker","Chef iQ","Amazon US","https://chefiq.com/","https://chefiq.com/","support@chefiq.com","★★★★★","Bluetooth接続でレシピ自動設定・調理ガイド機能付きのスマート圧力鍋。料理初心者でもプロ品質の料理が可能。"],
  [1693,"キッチン・調理器具","Chef iQ World's Smartest Scale","Chef iQ","Amazon US","https://chefiq.com/","https://chefiq.com/","support@chefiq.com","★★★★☆","レシピ連動の栄養成分・重量計測スマートキッチンスケール。食材計量と栄養管理を同時にこなせる便利な製品。"],
  [1694,"キッチン・調理器具","Spinn Coffee Machine Pro","Spinn","Amazon US","https://spinn.com/","https://spinn.com/","support@spinn.com","★★★★★","コーヒー豆を直接挽いてエスプレッソ〜コールドブリューまで全自動で作るスマートコーヒーマシン。"],

  // --- スマートホーム・インテリア・照明 ---
  [1514,"スマートホーム・インテリア・照明","Aqara Smart Lock U400 UWB","Aqara","Amazon US","https://us.aqara.com/products/smart-lock-u400","https://www.aqara.com/","support@aqara.com","★★★★☆","超広帯域（UWB）技術で近づくだけでロック解除。Apple HomeKey対応。セキュリティ関心が高い日本市場に適合。"],
  [1515,"スマートホーム・インテリア・照明","Ultraloq Bolt Sense Palm Vein Lock","U-tec","Amazon US","https://ultraloq.com/","https://ultraloq.com/","support@ultraloq.com","★★★★☆","手のひら静脈認証スマートロック。非接触でセキュリティ性能が高い。スマートホーム需要が伸びる日本向け。"],
  [1516,"スマートホーム・インテリア・照明","Govee AI Lighting System","Govee","Amazon US","https://us.govee.com/","https://us.govee.com/","distribution@govee.com","★★★★☆","AI搭載スマート照明システム。Samsung SmartThings連携対応。インテリア志向が強い日本市場でも需要拡大中。"],
  [1517,"スマートホーム・インテリア・照明","Govee Neon Rope Light 2","Govee","Amazon US","https://us.govee.com/","https://us.govee.com/","distribution@govee.com","★★★★☆","RGBIC対応の柔軟なネオンRGBテープライト。部屋の雰囲気づくりに最適でSNS映えも抜群。"],
  [1518,"スマートホーム・インテリア・照明","Govee RGBIC Floor Lamp Pro","Govee","Amazon US","https://us.govee.com/","https://us.govee.com/","distribution@govee.com","★★★★☆","音楽連動・ゲーム連動のスマートフロアランプ。ゲーミング・ストリーミング需要が高まる日本に好適。"],
  [1519,"スマートホーム・インテリア・照明","LeafyPod AI Smart Self-Watering Planter","LeafyPod","LeafyPod公式","https://www.theleafypod.com/","https://www.leafypod.one/","hello@leafypod.one","★★★★☆","AIが植物の状態を検知し自動で水やりを行うスマートプランター。テスラ・Meta出身者が創業したスタートアップ。"],
  [1520,"スマートホーム・インテリア・照明","NuraLogix Longevity Mirror","NuraLogix","NuraLogix公式","https://longevitymirror.com/","https://longevitymirror.com/","info@longevitymirror.com","★★★★★","30秒の顔スキャンで将来の健康リスクを予測するAIスマートミラー。日本の健康意識の高い層に刺さる革新製品。"],
  [1521,"スマートホーム・インテリア・照明","Reolink TrackFlex Floodlight WiFi","Reolink","Amazon US","https://reolink.com/","https://reolink.com/","brand@reolink.com","★★★★☆","4K二眼カメラ搭載の360度フラッドライト付きセキュリティカメラ。防犯意識向上の日本市場に需要がある。"],
  [1522,"スマートホーム・インテリア・照明","Baseus Security X1 Pro AI Dual Camera","Baseus","Amazon US","https://www.baseus.com/products/security-x1-pro-outdoor-dual-camera","https://www.baseus.com/","pr@baseus.com","★★★★☆","世界初AI二眼追跡型ソーラーセキュリティカメラ。サブスク不要でローカル録画可能。セキュリティ重視層に訴求。"],
  [1523,"スマートホーム・インテリア・照明","Gardyn Studio 2 Indoor Hydroponic Garden","Gardyn","Gardyn公式","https://mygardyn.com/product/gardyn-studio-gen2/","https://mygardyn.com/","support@mygardyn.com","★★★★☆","CES 2026 Innovation Award受賞の屋内水耕栽培システム。自産自消トレンドが高まる日本でも需要が期待できる。"],
  [1608,"スマートホーム・インテリア・照明","Hatch Restore 2 Sleep Alarm Clock","Hatch","Amazon US","https://www.hatch.co/","https://www.hatch.co/","support@hatch.co","★★★★★","日の出を再現して自然に目覚めるスマートサンライズ目覚まし時計。睡眠の質改善ニーズが高い日本市場に好適。"],
  [1609,"スマートホーム・インテリア・照明","Govee Outdoor String Lights H70C2","Govee","Amazon US","https://us.govee.com/","https://us.govee.com/","distribution@govee.com","★★★★☆","防水IP65対応のRGBIC屋外スマート電球ストリングライト。テラス・庭・バルコニーの雰囲気づくりに最適。"],
  [1610,"スマートホーム・インテリア・照明","Nanoleaf Lines 3D Smart Lighting","Nanoleaf","Amazon US","https://nanoleaf.me/","https://nanoleaf.me/","help@nanoleaf.me","★★★★★","3D立体形状のスマートパネルライト。Thread対応でスマートホームエコシステムに完全統合。インテリア感度の高い日本向け。"],
  [1641,"スマートホーム・インテリア・照明","Flair Smart Home Vent System","Flair","Flair公式","https://flair.co/","https://flair.co/","support@flair.co","★★★★☆","部屋ごとに温度を個別管理できるスマートベントシステム。エネルギー効率向上と快適性向上を同時実現。"],
  [1642,"スマートホーム・インテリア・照明","Emporia Smart Home Energy Monitor","Emporia Energy","Amazon US","https://www.emporiaenergy.com/","https://www.emporiaenergy.com/","support@emporiaenergy.com","★★★★☆","家庭の電力消費をリアルタイム監視するスマートエネルギーモニター。電気代節約意識が高い日本家庭向け。"],
  [1643,"スマートホーム・インテリア・照明","Level Lock+ Smart Door Lock","Level Home","Amazon US","https://www.levellock.com/","https://www.levellock.com/","support@levellock.com","★★★★☆","ドアの内部に完全に隠れる業界最薄スマートロック。外観変更なしに設置でき賃貸にも対応しやすい画期的製品。"],
  [1644,"スマートホーム・インテリア・照明","August Wi-Fi Smart Lock Pro","August","Amazon US","https://august.com/","https://august.com/","support@august.com","★★★★☆","既存デッドボルトに装着するだけで設置できるスマートロック。オートロック・遠隔操作・来客管理に対応。"],
  [1645,"スマートホーム・インテリア・照明","Brilliant Smart Home Control Panel","Brilliant Home Technology","Amazon US","https://www.brilliant.tech/","https://www.brilliant.tech/","support@brilliant.tech","★★★★★","Amazon Alexa内蔵のタッチスクリーン型スマートホームコントロールパネル。壁スイッチと交換するだけで全スマートデバイスを一元管理。"],
  [1697,"スマートホーム・インテリア・照明","Halo Rise Sleep-Wake Light","Halo (Amazon)","Amazon US","https://www.amazon.com/halo-rise","https://www.amazon.com/","sleep@amazon.com","★★★★☆","睡眠追跡機能と朝の日の出シミュレーションを組み合わせたスマートベッドサイドライト。睡眠改善ニーズに合致。"],
  [1698,"スマートホーム・インテリア・照明","Manta Sleep Mask Pro","Manta Sleep","Amazon US","https://mantasleep.com/","https://mantasleep.com/","support@mantasleep.com","★★★★★","目を圧迫しない3Dカップ設計の究極の遮光スリープマスク。完全遮光・快適素材で睡眠の質向上に寄与。"],
  [1699,"スマートホーム・インテリア・照明","Nanoleaf Skylight Smart Ceiling Panel","Nanoleaf","Amazon US","https://nanoleaf.me/","https://nanoleaf.me/","help@nanoleaf.me","★★★★★","天井設置型の大型スマートLEDパネルライト。RGBWで自然光・夕暮れ・ナイトライトを演出。インテリア性が高い。"],

  // --- ウェアラブル・ヘルス・フィットネス ---
  [1524,"ウェアラブル・ヘルス・フィットネス","SOND Dreambuds Sleep Platform","SOND","SOND公式","https://shop.sond.com/products/dreambuds-pro","https://sond.com/","orders@sond.com","★★★★★","耳で12種の生体信号を計測しリアルタイムで睡眠を最適化する革新的イヤバッド。睡眠課題が深刻な日本に最適。"],
  [1525,"ウェアラブル・ヘルス・フィットネス","Circular Ring 2 Smart Health Ring","Circular","Circular公式","https://www.circular.xyz/","https://www.circular.xyz/","contact@circular.xyz","★★★★★","ECG対応スマートリング。睡眠・活動・回復を24時間追跡。Ouraに続くスマートリング市場急成長の中で台頭。"],
  [1526,"ウェアラブル・ヘルス・フィットネス","sun-a-wear UV Smart Tracker","sun-a-wear","sun-a-wear公式","https://sun-a-wear.us/products/uv-tracker-and-app","https://sun-a-wear.us/","support@sun-a-wear.com","★★★★☆","充電不要の防水型ウェアラブルUVセンサー。日焼け予防意識が高い日本市場向けに高い訴求力。"],
  [1527,"ウェアラブル・ヘルス・フィットネス","HidrateSpark PRO 2 Smart Water Bottle","HidrateSpark","HidrateSpark公式","https://hidratespark.com/","https://hidratespark.com/","support@hidratespark.com","★★★★☆","センサーで水分補給量を計測しLEDで知らせるスマートボトル。健康管理意識の高い日本市場に需要あり。"],
  [1528,"ウェアラブル・ヘルス・フィットネス","LARQ PureVis 2 Self-Cleaning Bottle","LARQ","LARQ公式","https://www.livelarq.com/","https://www.livelarq.com/","hello@livelarq.com","★★★★☆","UV-C LEDで自動除菌する世界初の自己洗浄ウォーターボトル。衛生意識の高い日本に最適。BRITA傘下。"],
  [1529,"ウェアラブル・ヘルス・フィットネス","WaterH Vita Smart Water Bottle","WaterH","Amazon US","https://www.waterh.com/products/filtered-waterh-vita-smart-water-bottle","https://www.waterh.com/","hello@waterh.com","★★★★☆","フィルター付き・Bluetooth連動で水分摂取を可視化するスマートボトル。健康志向の日本市場に訴求。"],
  [1530,"ウェアラブル・ヘルス・フィットネス","Rally Orbital Massager Recovery Tool","Rally","Rally公式","https://getrally.co/products/orbital-massager","https://getrally.co/","support@getrally.co","★★★★☆","TIME誌Best Inventions 2025選出の独自オービタル振動マッサージガン。スポーツ回復・筋膜リリースに最適。"],
  [1531,"ウェアラブル・ヘルス・フィットネス","Hydrow Arc Smart Rowing Machine","Hydrow","Hydrow公式","https://hydrow.com/products/rowers/hydrow-arc-rower/","https://hydrow.com/","commercialsales@hydrow.com","★★★★★","ライブ映像連動の高臨場感スマートローイングマシン。フィットネス投資意欲の高い日本向けプレミアム製品。"],
  [1532,"ウェアラブル・ヘルス・フィットネス","Hydrow Wave Compact Rower","Hydrow","Hydrow公式","https://hydrow.com/products/rowers/hydrow-wave/","https://hydrow.com/","commercialsales@hydrow.com","★★★★☆","コンパクト設計の高性能スマートローイングマシン。狭い日本の住宅事情に合わせたサイズ感が魅力。"],
  [1533,"ウェアラブル・ヘルス・フィットネス","Speediance GYM One AI Home Gym","Speediance","Speediance公式","https://www.speediance.com/","https://www.speediance.com/","after-sales@speediance.com","★★★★★","AI搭載の一体型スマートホームジム。160種以上のクラスと自動負荷調整でパーソナルジム体験を自宅で実現。"],
  [1534,"ウェアラブル・ヘルス・フィットネス","Speediance Gym Monster 2 Smart Rack","Speediance","Speediance公式","https://www.speediance.com/","https://www.speediance.com/","after-sales@speediance.com","★★★★★","AIパワーラックシステム。リアルタイムフォーム解析・プログラム管理で本格フィットネスが自宅で可能。"],
  [1535,"ウェアラブル・ヘルス・フィットネス","Hyperice Hypervolt Go 2 Massage Gun","Hyperice","Amazon US","https://hyperice.com/","https://hyperice.com/","support@hyperice.com","★★★★☆","コンパクトで高性能なパーカッションマッサージガン。NBA・NHLなど世界トップ選手が愛用する定番ブランド。"],
  [1536,"ウェアラブル・ヘルス・フィットネス","Hyperice Normatec 3 Compression Boots","Hyperice","Amazon US","https://hyperice.com/","https://hyperice.com/","support@hyperice.com","★★★★☆","空気圧式リカバリーブーツ。脚の疲労回復に効果的。プロスポーツ・フィットネス層の需要が高い製品。"],
  [1537,"ウェアラブル・ヘルス・フィットネス","MesoLyft Infrared Beauty Wearable","MesoLyft Technologies","MesoLyft公式","https://mesolyft.com/","https://mesolyft.com/","info@mesolyft.com","★★★★☆","赤外線照射による家庭用美容ウェアラブルデバイス。マイクロニードリングセラムとの組み合わせで相乗効果。"],
  [1538,"ウェアラブル・ヘルス・フィットネス","VITURE Pro XR Smart Glasses","VITURE","Amazon US","https://www.viture.com/","https://www.viture.com/","bd@viture.com","★★★★★","135インチ仮想スクリーン・Harmanオーディオ搭載のARグラス。軽量コンパクトで日本市場の先進技術需要に合致。"],
  [1539,"ウェアラブル・ヘルス・フィットネス","VITURE One XR Glasses","VITURE","Amazon US","https://www.viture.com/","https://www.viture.com/","bd@viture.com","★★★★★","エントリーモデルのXRグラス。軽量設計でスマートフォン・PCとの接続が容易。ゲーム・動画視聴に最適。"],
  [1603,"ウェアラブル・ヘルス・フィットネス","Lumen Metabolism Tracker","Lumen","Amazon US","https://www.lumen.me/","https://www.lumen.me/","support@lumen.me","★★★★★","呼気CO2で代謝をリアルタイム計測するデバイス。食事とダイエットを科学的に最適化。日本のダイエット需要に刺さる。"],
  [1604,"ウェアラブル・ヘルス・フィットネス","Withings ScanWatch Nova 2 Hybrid Smartwatch","Withings","Amazon US","https://www.withings.com/","https://www.withings.com/","support@withings.com","★★★★★","心電図・睡眠・血中酸素対応のアナログ風フランス発ハイブリッドスマートウォッチ。デザイン性と機能性を両立。"],
  [1605,"ウェアラブル・ヘルス・フィットネス","Withings Body Comp Smart Scale","Withings","Amazon US","https://www.withings.com/","https://www.withings.com/","support@withings.com","★★★★★","内臓脂肪・骨格筋・体内水分を計測できるWi-Fi接続スマートスケール。健康管理に積極的な日本市場向け。"],
  [1606,"ウェアラブル・ヘルス・フィットネス","Apollo Neuro Stress Relief Wearable","Apollo Neuroscience","Amazon US","https://apolloneuro.com/","https://apolloneuro.com/","support@apolloneuro.com","★★★★★","振動波形でストレス反応を抑制するウェアラブルデバイス。科学的根拠に基づくストレス管理ツール。"],
  [1607,"ウェアラブル・ヘルス・フィットネス","Sensate 2 Vagus Nerve Stimulator","Sensate","Sensate公式","https://www.getsensate.com/","https://www.getsensate.com/","support@getsensate.com","★★★★★","迷走神経刺激で5分間でリラックスを実現するウェアラブルデバイス。瞑想・ストレス軽減需要の高い日本市場に最適。"],
  [1631,"ウェアラブル・ヘルス・フィットネス","Eight Sleep Pod 4 Smart Mattress Cover","Eight Sleep","Eight Sleep公式","https://www.eightsleep.com/","https://www.eightsleep.com/","support@eightsleep.com","★★★★★","水冷温度制御で睡眠環境を自動最適化するスマートマットレスカバー。睡眠の質改善ニーズが高い日本市場に高い訴求力。"],
  [1632,"ウェアラブル・ヘルス・フィットネス","Flowtime Biosensing Meditation Headband","Flowtime","Amazon US","https://www.flowtime.cn/","https://www.flowtime.cn/","support@meetflowtime.com","★★★★★","脳波・心拍・呼吸をリアルタイム計測してガイドする瞑想用バイオセンシングヘッドバンド。マインドフルネス市場向け。"],
  [1633,"ウェアラブル・ヘルス・フィットネス","Muse S Gen 2 Brain Sensing Headband","Muse (InteraXon)","Amazon US","https://choosemuse.com/","https://choosemuse.com/","support@choosemuse.com","★★★★★","EEG脳波計測で瞑想の深さをリアルタイム追跡するスマートヘッドバンド。睡眠トラッキングも対応の多機能製品。"],
  [1634,"ウェアラブル・ヘルス・フィットネス","Levels Health Continuous Glucose Monitor","Levels Health","Levels公式","https://www.levelshealth.com/","https://www.levelshealth.com/","support@levelshealth.com","★★★★★","CGMとAIによる食事・運動に対する血糖反応分析サービス。健康意識が高く予防医療投資に積極的な日本市場向け。"],
  [1635,"ウェアラブル・ヘルス・フィットネス","Airthings View Plus Air Quality Monitor","Airthings","Amazon US","https://www.airthings.com/","https://www.airthings.com/","support@airthings.com","★★★★★","ラドン・CO2・VOC・湿度・気圧を計測するスマート空気質モニター。室内空気質意識が高い日本市場に需要大。"],
  [1636,"ウェアラブル・ヘルス・フィットネス","Elemind Neural Interface Headband","Elemind","Elemind公式","https://elemind.com/","https://elemind.com/","info@elemind.com","★★★★★","神経インターフェース技術で入眠を促進するEEGヘッドバンド。睡眠障害に悩む日本人向けに高い訴求力。"],
  [1637,"ウェアラブル・ヘルス・フィットネス","Kokoon Nightbuds Sleep Headphones","Kokoon","Amazon US","https://kokoon.io/","https://kokoon.io/","support@kokoon.io","★★★★★","睡眠時に装着できる超薄型スリープヘッドフォン。EEGで入眠を検知し自動フェードアウト機能搭載。不眠対策製品。"],
  [1677,"ウェアラブル・ヘルス・フィットネス","Polar H10 Heart Rate Monitor","Polar","Amazon US","https://www.polar.com/","https://www.polar.com/","support@polar.com","★★★★★","医療グレードの精度を誇る胸部装着型心拍センサー。Bluetooth・ANT+対応でほぼ全スポーツアプリと連携可能。"],
  [1678,"ウェアラブル・ヘルス・フィットネス","Polar Vantage V3 Multi-Sport Watch","Polar","Amazon US","https://www.polar.com/","https://www.polar.com/","support@polar.com","★★★★★","トップアスリート向けGPS・心電図・皮膚温度計測対応マルチスポーツウォッチ。フィンランド発の老舗スポーツブランド。"],
  [1679,"ウェアラブル・ヘルス・フィットネス","Biostrap EVO Health Wearable","Biostrap","Biostrap公式","https://biostrap.com/","https://biostrap.com/","support@biostrap.com","★★★★☆","臨床グレードの光学センサーで心拍・血中酸素・呼吸数・睡眠を計測するスマートリストバンド。"],
  [1695,"ウェアラブル・ヘルス・フィットネス","Ultrahuman Ring AIR Smart Ring","Ultrahuman","Amazon US","https://www.ultrahuman.com/","https://www.ultrahuman.com/","support@ultrahuman.com","★★★★★","チタン製の軽量スマートリング。睡眠・活動・回復・代謝スコアをAIで解析。サブスク料金不要で日本市場にも好適。"],
  [1696,"ウェアラブル・ヘルス・フィットネス","RingConn Gen 2 Health Smart Ring","RingConn","Amazon US","https://www.ringconn.com/","https://www.ringconn.com/","support@ringconn.com","★★★★★","最長10日バッテリーと豊富なヘルス指標を持つスマートリング。Oura特許訴訟解決後に米国で再販売開始。"],

  // --- アウトドア・スポーツ・旅行 ---
  [1540,"アウトドア・スポーツ・旅行","Nomadic.cool Backpack Compressor Cooler","Nomadic.cool","Indiegogo","https://www.nomadic.cool/","https://www.nomadic.cool/","heinz@szolarz.at","★★★★☆","従来品の1/4の重量で実現した超軽量バックパック型コンプレッサークーラー。アウトドアブーム継続の日本市場向け。"],
  [1541,"アウトドア・スポーツ・旅行","BLUETTI AC300B Portable Power Station","BLUETTI","Amazon US","https://bluetti.com/","https://bluetti.com/","business@bluetti.com","★★★★☆","大容量ポータブル電源。ソーラーパネル充電対応でキャンプ・防災・車中泊に最適。日本の防災需要にも合致。"],
  [1542,"アウトドア・スポーツ・旅行","BLUETTI EB3A Mini Power Station","BLUETTI","Amazon US","https://bluetti.com/","https://bluetti.com/","business@bluetti.com","★★★★☆","コンパクトで軽量な小型ポータブル電源。携帯・アウトドア・防災用として日本の需要が高い分野。"],
  [1543,"アウトドア・スポーツ・旅行","Jackery Explorer 2000 Pro Solar Generator","Jackery","Amazon US","https://www.jackery.com/","https://www.jackery.com/","info@jackerysolutions.com","★★★★☆","ソーラー充電対応の大容量ポータブル電源。防災・キャンプ・電力不安に対応する人気製品。"],
  [1544,"アウトドア・スポーツ・旅行","Jackery SolarSaga 200W Solar Panel","Jackery","Amazon US","https://www.jackery.com/","https://www.jackery.com/","info@jackerysolutions.com","★★★★☆","高効率ポータブルソーラーパネル。折りたたみ式で携帯性が高く、再エネ需要が高まる日本に好適。"],
  [1545,"アウトドア・スポーツ・旅行","Segway MUXI Utility E-Bike","Segway","Segway公式","https://store.segway.com/muxi-electric-bike","https://store.segway.com/","support@segway.com","★★★★☆","実用性重視のカーゴe-バイク。通勤・配達・レジャー用途で需要が伸びる日本の都市部向け。"],
  [1546,"アウトドア・スポーツ・旅行","Segway Xaber 300 Electric Dirt Bike","Segway","Segway公式","https://store.segway.com/","https://store.segway.com/","support@segway.com","★★★★☆","CES 2026発表の電動ダートバイク。ダカールラリー参戦モデルからの派生で本格性能を実現。"],
  [1548,"アウトドア・スポーツ・旅行","Db Essential Bag 26L Travel Backpack","Db","Db公式","https://dbjourney.com/","https://dbjourney.com/","support@dbjourney.com","★★★★☆","ノルウェー発・LVMH出資のデザイン性と機能性を兼ね備えたトラベルバックパック。品質志向の日本市場に好適。"],
  [1549,"アウトドア・スポーツ・旅行","WANDRD PRVKE V4 Photography Backpack","WANDRD","WANDRD公式","https://www.wandrd.com/collections/backpacks","https://www.wandrd.com/","wholesale@wandrd.com","★★★★☆","フォトグラファー向け多機能バックパック。全面リデザインのV4モデル。撮影旅行・アウトドア兼用で日本需要高。"],
  [1550,"アウトドア・スポーツ・旅行","WANDRD HEXAD Carryall Duffel Bag","WANDRD","WANDRD公式","https://www.wandrd.com/collections/backpacks","https://www.wandrd.com/","wholesale@wandrd.com","★★★★☆","バックパック・ダッフル・トートの3WAY仕様。旅行・日常・アウトドア全対応の多機能バッグ。"],
  [1611,"アウトドア・スポーツ・旅行","BioLite CampStove 2+ Camping Stove","BioLite","Amazon US","https://www.bioliteenergy.com/","https://www.bioliteenergy.com/","info@bioliteenergy.com","★★★★★","発電機能付きウッドバーニングキャンプストーブ。薪で調理しながらUSB充電が可能な革新アウトドアギア。"],
  [1612,"アウトドア・スポーツ・旅行","BioLite SolarPanel 10+ Solar Charger","BioLite","Amazon US","https://www.bioliteenergy.com/","https://www.bioliteenergy.com/","info@bioliteenergy.com","★★★★☆","最適角度自動調整機能付きの高効率ポータブルソーラーパネル。太陽光充電でキャンプ・登山での電力確保に最適。"],
  [1613,"アウトドア・スポーツ・旅行","Helinox Chair One Ultralight Camp Chair","Helinox","Amazon US","https://helinox.com/","https://helinox.com/","support@helinox.com","★★★★★","860gの超軽量コンパクト折りたたみアウトドアチェア。韓国発・世界的人気ブランドで日本キャンプ市場でも需要大。"],
  [1614,"アウトドア・スポーツ・旅行","Helinox Table One Ultralight Camp Table","Helinox","Amazon US","https://helinox.com/","https://helinox.com/","support@helinox.com","★★★★★","900gの超軽量コンパクトキャンプテーブル。アルミポールとDACポールの組み合わせで高剛性を実現。"],
  [1615,"アウトドア・スポーツ・旅行","Rumpl Original Puffy Blanket","Rumpl","Amazon US","https://www.rumpl.com/","https://www.rumpl.com/","hello@rumpl.com","★★★★☆","軽量・コンパクトで洗濯可能なアウトドア用プリントダウンブランケット。キャンプ・スポーツ観戦・旅行に最適。"],
  [1616,"アウトドア・スポーツ・旅行","MPOWERD Luci Pro Solar Inflatable Lantern","MPOWERD","Amazon US","https://www.mpowerd.com/","https://www.mpowerd.com/","hello@mpowerd.com","★★★★★","ソーラー充電式の折りたたみ型防水LEDランタン。USB充電機能付きで非常時・キャンプ双方に活躍。"],

  // --- ペット用品 ---
  [1551,"ペット用品","SATELLAI Collar Go Smart GPS Dog Collar","SATELLAI","SATELLAI公式","https://satellai.com/products/collar-go","https://satellai.com/","support@satellai.com","★★★★★","衛星GPSと専用AI（PetSense™）を搭載した次世代スマートペットカラー。CES 2025デビューのHuami出身チーム創業。"],
  [1552,"ペット用品","Woofwoof Lux AI Dog Washing Machine","Allviews Technology","Woofwoof Lux公式","https://www.woofwooflux.com/","https://www.woofwooflux.com/","support@woofwooflux.com","★★★★★","世界初AIパワードの全自動犬用洗浄・乾燥マシン。CES 2025・2026に出展。ペットを家族として扱う日本文化に刺さる。"],
  [1553,"ペット用品","PetPace Health 3.0 AI Smart Collar","PetPace","PetPace公式","https://petpace.com/product/collar-v3/","https://petpace.com/","commercial@petpace.com","★★★★★","24時間バイタル監視と内蔵獣医遠隔医療サービスを組み合わせた世界初のAIスマートカラー。FDA認証取得済み。"],
  [1554,"ペット用品","Fi Smart Dog Collar Series 3","Fi","tryfi.com公式","https://tryfi.com/","https://tryfi.com/","support@tryfi.com","★★★★★","GPSリアルタイム追跡と活動量・睡眠モニタリングを搭載した高精度ドッグカラー。電池持続2ヶ月以上。"],
  [1555,"ペット用品","Pawfit 3S GPS Pet Tracker","Pawfit","Pawfit公式","https://www.pawfit.com/en-us/product/pawfit-3s-pet-tracker.html","https://www.pawfit.com/","support@pawfit.com","★★★★☆","猫・犬対応の防水GPS追跡デバイス。全世界4G対応。活動量・睡眠・カロリーを追跡。"],
  [1557,"ペット用品","Tractive GPS DOG XL Health & Fitness","Tractive","Tractive公式","https://tractive.com/en/pd/gps-tracker-dog","https://tractive.com/","support@tractive.com","★★★★★","次世代GPS犬用カラーに心拍・呼吸数・吠え検知を搭載。130万人以上が利用するNo.1ペットGPSブランド。"],
  [1558,"ペット用品","Tractive GPS CAT 4 Tracker","Tractive","Tractive公式","https://tractive.com/","https://tractive.com/","support@tractive.com","★★★★★","猫専用GPS追跡デバイス。超軽量設計で外出猫の追跡に最適。日本では室内外両用猫の見守りに需要がある。"],
  [1646,"ペット用品","Furbo 360 AI Dog Camera","Furbo","Amazon US","https://furbo.com/","https://furbo.com/","support@furbo.com","★★★★★","AI犬吠え検知・360度パン・自動追尾・おやつ投げ機能付きペットモニカメラ。ペット見守り市場でトップブランド。"],
  [1647,"ペット用品","Enabot ROLA AI Pet Companion Robot","Enabot","Amazon US","https://www.enabot.com/","https://www.enabot.com/","contact@enabot.com","★★★★★","自律移動・ペット追跡・おやつ投げ・音声通話機能付きペットコンパニオンロボット。ペットの孤独対策に最適。"],
  [1648,"ペット用品","FluentPet Smart Button Communication Kit","FluentPet","FluentPet公式","https://fluent.pet/","https://fluent.pet/","support@fluent.pet","★★★★★","犬・猫が押して意思疎通を図れる音声録音ボタンキット。「ペット語学」という新しいコンセプトで注目を集める。"],
  [1649,"ペット用品","Embark Breed & Health DNA Test Kit","Embark Veterinary","Amazon US","https://embarkvet.com/","https://embarkvet.com/","support@embarkvet.com","★★★★★","230以上の健康状態リスクを検査できる犬専用DNA検査キット。ペット医療費節約につながる予防的ケアツール。"],
  [1650,"ペット用品","Townew T4 Smart Self-Sealing Trash Can","Townew","Amazon US","https://www.townew.co/","https://www.townew.co/","support@townew.co","★★★★★","自動袋交換・自動シーリング・ハンズフリー開閉のスマートゴミ箱。衛生意識が高く清潔好きな日本市場に完全合致。"],
  [1686,"ペット用品","Litter-Robot 4 Automatic Self-Cleaning Litter Box","Litter-Robot","Amazon US","https://www.litter-robot.com/","https://www.litter-robot.com/","support@litter-robot.com","★★★★★","完全自動洗浄・健康モニタリング機能付きスマート自動猫トイレ。ペットの健康管理と衛生維持を同時実現。"],
  [1688,"ペット用品","Halo Collar 3 AI GPS Dog Fence","Halo Collar","Halo公式","https://www.halocollar.com/","https://www.halocollar.com/","support@halocollar.com","★★★★★","AI搭載の見えない電子フェンスシステム。Cesar Milanと共同開発で犬のトレーニングにも活用できる製品。"],
  [1703,"ペット用品","Petcube Play 2 Interactive Pet Camera","Petcube","Amazon US","https://petcube.com/","https://petcube.com/","info@petcube.com","★★★★☆","2Kカメラ・レーザーポインター・おやつ投げ機能付きのインタラクティブペットカメラ。遠隔でペットと遊べる機能が充実。"],
  [1704,"ペット用品","PetSafe Smart Feed 2.0 Automatic Feeder","PetSafe","Amazon US","https://store.petsafe.net/","https://store.petsafe.net/","support@petsafe.net","★★★★☆","スマートフォンで給餌時間・量を管理できる自動ペットフィーダー。遠隔操作・スケジュール設定で留守中も安心。"],
  [1705,"ペット用品","Whistle Go Explore Health + Location Pet Tracker","Whistle Labs","Amazon US","https://www.whistle.com/","https://www.whistle.com/","support@whistle.com","★★★★★","GPS追跡・健康モニタリング・ライフステージ管理を統合したオールインワンペットトラッカー。異常アラートも搭載。"],

  // --- テクノロジー・ガジェット ---
  [1561,"テクノロジー・ガジェット","XPANCEO Smart Contact Lens","XPANCEO","XPANCEO公式","https://www.xpanceo.com/","https://www.xpanceo.com/","hello@xpanceo.com","★★★★★","AR表示・健康モニタリング・視力矯正を一体化したスマートコンタクトレンズ。2026年に消費者向け発売予定の革新製品。"],
  [1563,"テクノロジー・ガジェット","Scrap Labs Scrap 1 Desktop Metal 3D Printer","Scrap Labs","Scrap Labs公式","https://www.scraplabs3d.com/","https://www.scraplabs3d.com/","info@scraplabs3d.com","★★★★★","9,600ドルから購入可能な卓上型レーザー粉末床溶融金属3Dプリンター。メイカー・研究者・中小製造業向けに革命的製品。"],
  [1564,"テクノロジー・ガジェット","Bambu Lab H2D All-in-One Maker Station","Bambu Lab","Amazon US","https://bambulab.com/en-us/h2d","https://bambulab.com/","support@bambulab.com","★★★★★","3Dプリント・レーザー彫刻・カッティングを統合したメイカーステーション。クリエイター・DIY愛好家に最適。"],
  [1565,"テクノロジー・ガジェット","xTool F1 Ultra Laser Cutter","xTool","xTool公式","https://www.xtool.com/","https://www.xtool.com/","support@xtool.com","★★★★★","ファイバーレーザーとダイオードレーザーを搭載したデュアルレーザーカッター。金属・非金属両対応で高性能。"],
  [1566,"テクノロジー・ガジェット","Creality Falcon A1 Smart Laser Engraver","Creality","Creality公式","https://www.crealityfalcon.com/","https://www.crealityfalcon.com/","support@creality.com","★★★★☆","AIアシスト機能を搭載したスマートレーザー彫刻機。初心者でも高精度加工が可能。DIY需要が高い日本市場向け。"],
  [1567,"テクノロジー・ガジェット","Clicks Technology iPhone Keyboard Case","Clicks Technology","Clicks公式","https://www.getclicks.com/","https://www.getclicks.com/","hello@getclicks.com","★★★★☆","iPhoneにクリップオンできる物理キーボードケース。UK発の斬新なアイデア製品。スマートフォン生産性向上アイテム。"],
  [1617,"テクノロジー・ガジェット","Flipper Zero Portable Multi-Tool","Flipper Devices","Amazon US","https://flipperzero.one/","https://flipperzero.one/","info@flipperzero.one","★★★★★","RFID・NFC・赤外線・Bluetooth・Bad USBを搭載したポータブルハッカー向けマルチツール。セキュリティ研究者に人気。"],
  [1618,"テクノロジー・ガジェット","Rabbit R1 AI Hardware Device","Rabbit","Amazon US","https://www.rabbit.tech/","https://www.rabbit.tech/","support@rabbit.tech","★★★★★","LLMベースのAIエージェントを搭載した独立型ポケットAIデバイス。スマートフォンの次世代インターフェースとして注目。"],
  [1619,"テクノロジー・ガジェット","PLAUD NotePin AI Voice Recorder","PLAUD","Amazon US","https://www.plaud.ai/","https://www.plaud.ai/","support@plaud.ai","★★★★★","クレジットカードサイズのAI音声レコーダー。会議・講義・インタビューの自動文字起こし・要約が可能。"],
  [1620,"テクノロジー・ガジェット","HoverAir X1 Self-Flying Camera","HoverAir","Amazon US","https://www.hoverair.com/","https://www.hoverair.com/","support@hoverair.com","★★★★★","手から飛び立つ自撮り用自律飛行ドローン。125g以下で規制対象外（国内条件確認要）。旅行・スポーツ撮影に最適。"],
  [1621,"テクノロジー・ガジェット","reMarkable 2 Digital Paper Tablet","reMarkable","Amazon US","https://remarkable.com/","https://remarkable.com/","support@remarkable.com","★★★★★","紙のような書き心地を再現した電子ペーパータブレット。ノート・スケッチ・PDFアノテーションに最適な生産性ツール。"],
  [1622,"テクノロジー・ガジェット","Framework Laptop 13 Modular Laptop","Framework","Framework公式","https://frame.work/","https://frame.work/","support@frame.work","★★★★★","ユーザー自身でパーツ交換・アップグレード可能なモジュラーノートPC。サステナビリティ意識の高い日本市場に訴求。"],
  [1623,"テクノロジー・ガジェット","Elegoo Saturn 4 Ultra 12K 3D Printer","Elegoo","Amazon US","https://www.elegoo.com/","https://www.elegoo.com/","support@elegoo.com","★★★★★","12K解像度の超高精細LCD光造形3Dプリンター。フィギュア・歯科モデル・ジュエリー製造に最適。日本のホビー市場向け。"],
  [1638,"テクノロジー・ガジェット","Bambu Lab X1 Carbon 3D Printer","Bambu Lab","Amazon US","https://bambulab.com/","https://bambulab.com/","support@bambulab.com","★★★★★","多色印刷対応の高速高精度3Dプリンター。16色まで自動切り替え可能でフィギュア・プロトタイプ製作に最適。"],
  [1639,"テクノロジー・ガジェット","Prusa MK4S FDM 3D Printer","Prusa Research","Prusa公式","https://www.prusa3d.com/","https://www.prusa3d.com/","info@prusa3d.com","★★★★★","チェコ発の世界的ベストセラーオープンソース3Dプリンター。品質・信頼性・サポートで業界標準の地位を確立。"],
  [1640,"テクノロジー・ガジェット","Anycubic Photon Mono M7 Pro Resin Printer","Anycubic","Amazon US","https://www.anycubic.com/","https://www.anycubic.com/","support@anycubic.com","★★★★★","12K解像度の高精度LCD光造形3Dプリンター。大型プラットフォームで大型フィギュア・パーツ製作が可能。"],
  [1683,"テクノロジー・ガジェット","Airthings Wave Plus Smart Air Monitor","Airthings","Amazon US","https://www.airthings.com/","https://www.airthings.com/","support@airthings.com","★★★★★","ラドン・CO2・VOC・湿度・温度・気圧を計測するスマート壁掛け型空気質モニター。室内環境改善ニーズの高い日本向け。"],
  [1684,"テクノロジー・ガジェット","XREAL Air 2 Ultra AR Glasses","XREAL","Amazon US","https://www.xreal.com/","https://www.xreal.com/","support@xreal.com","★★★★★","Google Project Aura対応の次世代ARグラス。空間コンピューティング・ゲーム・動画視聴に最適な軽量ARデバイス。"],
  [1700,"テクノロジー・ガジェット","Elegoo Mars 4 Ultra 9K 3D Printer","Elegoo","Amazon US","https://www.elegoo.com/","https://www.elegoo.com/","support@elegoo.com","★★★★★","9K解像度の高速LCD光造形3Dプリンター。フィギュア・ジュエリー・デンタルモデル製造に最適な精度。"],
  [1701,"テクノロジー・ガジェット","Anycubic Kobra 2 Pro FDM Printer","Anycubic","Amazon US","https://www.anycubic.com/","https://www.anycubic.com/","support@anycubic.com","★★★★★","500mm/s超高速プリントが可能なFDM 3Dプリンター。自動レベリング・高速印刷で初心者から上級者まで対応。"],
  [1702,"テクノロジー・ガジェット","xTool S1 Enclosed Laser Cutter","xTool","xTool公式","https://www.xtool.com/","https://www.xtool.com/","support@xtool.com","★★★★★","密閉型エンクロージャー搭載の高精度ダイオードレーザーカッター。安全性が高く家庭・オフィスでも使用可能。"],

  // --- 美容・スキンケア ---
  [1568,"美容・スキンケア","Mito Red Light MitoGLOW LED Face Mask","Mito Red Light","Mito Red Light公式","https://mitoredlight.com/products/mitoglow-led-mask","https://mitoredlight.com/","info@mitoredlight.com","★★★★★","1064個のLED・4波長搭載の高性能赤色光療法フェイスマスク。2025年美容革新賞受賞。美容大国日本での需要大。"],
  [1569,"美容・スキンケア","Appotronics Laser Hair Regrowth Cap","Appotronics","Appotronics公式","https://www.appotronics.com/","https://www.appotronics.com/","machunli@appotronics.com","★★★★★","ALPD®レーザー技術を活用した発毛促進レーザーキャップ。CES 2026で発表。FDA承認申請中の革新製品。"],
  [1570,"美容・スキンケア","Appotronics Laser Hair Growth Band","Appotronics","Appotronics公式","https://www.appotronics.com/","https://www.appotronics.com/","machunli@appotronics.com","★★★★★","頭部前面・こめかみ対応のレーザー発毛バンド。外出中でも使用できるディスクリートな設計が特徴。"],
  [1572,"美容・スキンケア","iPolish Smart Color-Changing Nails","iPolish","iPolish公式","https://ipolish.fashion/","https://ipolish.fashion/","CustomerService@ipolish.fashion","★★★★☆","スマートフォンアプリで色変更できる次世代プレスオンスマートネイル。CES 2026出展。ファッション×テックの融合製品。"],
  [1573,"美容・スキンケア","Tech Activated Beauty SilverFusion Mask","Tech Activated Beauty","TAB公式","https://techactivatedbeauty.com/","https://techactivatedbeauty.com/","info@techactivatedbeauty.com","★★★★☆","ナノカレント技術搭載の導電性シルバー繊維シートマスク。美顔器要らずでリフトアップが可能な革新スキンケア。"],
  [1575,"美容・スキンケア","Shark CryoGlow LED + Cryo Face Mask","SharkBeauty","Amazon US","https://www.sharkbeauty.com/","https://www.sharkbeauty.com/","support@sharkbeauty.com","★★★★★","赤・青・近赤外線LEDと冷却クライオ技術を組み合わせた最先端フルフェイスマスク。米国スキンケアデバイス1位ブランド。"],
  [1576,"美容・スキンケア","Shark FlexStyle Air Styler","SharkBeauty","Amazon US","https://www.sharkbeauty.com/","https://www.sharkbeauty.com/","support@sharkbeauty.com","★★★★★","ドライヤーとスタイリングツールを統合したマルチスタイラー。Dysonに対抗する高性能かつコスパ型製品。"],
  [1624,"美容・スキンケア","HigherDOSE Infrared Sauna Blanket V3","HigherDOSE","Amazon US","https://higherdose.com/","https://higherdose.com/","support@higherdose.com","★★★★★","家庭用赤外線サウナブランケット。自宅でサウナ体験を実現。美容・デトックス・回復に効果的で日本のサウナブームに合致。"],
  [1625,"美容・スキンケア","Omnilux Contour Face LED Mask","Omnilux","Amazon US","https://omniluxled.com/","https://omniluxled.com/","info@omniluxled.com","★★★★★","クリニカルグレードのRED+NIRデュアル光療法フェイスマスク。FDA認証済み。エビデンスに基づく美容機器として日本でも需要大。"],
  [1626,"美容・スキンケア","CurrentBody Skin LED Light Therapy Mask","CurrentBody","Amazon US","https://www.currentbody.com/","https://www.currentbody.com/","info@currentbody.com","★★★★★","フレキシブルシリコン製LED光療法フェイスマスク。632nm+830nmの臨床波長でコラーゲン生成を促進。"],
  [1627,"美容・スキンケア","JOVS Venus Pro II IPL Hair Removal","JOVS","Amazon US","https://www.jovs.com/","https://www.jovs.com/","support@jovs.com","★★★★★","AI肌診断機能付き家庭用IPL脱毛器。自動スキンカラー検知でサロン品質の脱毛を自宅で実現。"],
  [1628,"美容・スキンケア","NIRA Pro Laser Anti-Aging Device","NIRA Skincare Laser","Amazon US","https://www.niraskincare.com/","https://www.niraskincare.com/","support@niraskincare.com","★★★★★","家庭用レーザーアンチエイジングデバイス。ニキビ跡・シワ・肌荒れに効果的な臨床グレードのレーザー治療器。"],
  [1629,"美容・スキンケア","SolaWave Advanced Skincare Wand","SolaWave","Amazon US","https://solawave.co/","https://solawave.co/","support@solawave.co","★★★★★","微弱電流・赤色光・温感・マッサージを組み合わせた多機能スキンケアワンド。リフトアップ・むくみ改善に効果的。"],
  [1630,"美容・スキンケア","NuFace Trinity+ Facial Toning Device","NuFace","Amazon US","https://www.mynuface.com/","https://www.mynuface.com/","support@mynuface.com","★★★★★","FDA認証の微弱電流フェイシャルトーニングデバイス。たるみ・シワ改善でリフトアップ効果が高い美顔器。"],

  // --- 子供・教育 ---
  [1577,"子供・教育","bondu AI Companion Plush Toy","bondu","bondu公式","https://bondu.com/products/bondu","https://bondu.com/","support@bondu.com","★★★★☆","AI搭載の会話型ぬいぐるみ。子供の質問・物語・ゲームに対応。27言語対応で国際家庭にも最適。SF起業家が5.3Mドル調達。"],
  [1578,"子供・教育","Nanit Pro Smart Baby Monitor","Nanit","Nanit公式","https://www.nanit.com/products/smart-baby-monitor","https://www.nanit.com/","help@nanit.com","★★★★★","AIによる睡眠追跡・呼吸モニタリング機能付き高解像度ベビーカメラ。新生児期から使える安心製品。"],
  [1579,"子供・教育","Owlet Dream Duo Monitoring System","Owlet","Owlet公式","https://owletcare.com/products/dream-duo-3","https://owletcare.com/","support@owletcare.com","★★★★★","FDA認可の血中酸素・心拍監視スマートソックス+ビデオモニターのセット。医療グレードで安心の新生児モニタリング。"],
  [1658,"子供・教育","KiwiCo Tinker Crate STEM Box","KiwiCo","KiwiCo公式","https://www.kiwico.com/","https://www.kiwico.com/","support@kiwico.com","★★★★★","毎月届くエンジニアリング・サイエンス系DIYキット。子供が実験・組み立てを通じてSTEMを楽しく学べる定番製品。"],
  [1659,"子供・教育","KiwiCo Eureka Crate Advanced Engineering","KiwiCo","KiwiCo公式","https://www.kiwico.com/","https://www.kiwico.com/","support@kiwico.com","★★★★★","14歳以上向けの高度なエンジニアリングプロジェクトキット。本格的な機械設計・電子回路の学習に最適。"],
  [1660,"子供・教育","Osmo Genius Starter Kit iPad","Osmo","Amazon US","https://www.playosmo.com/","https://www.playosmo.com/","support@playosmo.com","★★★★★","iPad用AR教育ゲームキット。カメラで実物を検知してデジタルゲームと連動する革新的学習ツール。"],
  [1661,"子供・教育","Sphero BOLT+ Robot Ball Coding Toy","Sphero","Amazon US","https://sphero.com/","https://sphero.com/","support@sphero.com","★★★★★","プログラミングで操作できるLED付きロボットボール。ブロックコーディングからJavaScriptまで学べる教育製品。"],
  [1662,"子供・教育","Wonder Workshop Dash Coding Robot","Wonder Workshop","Amazon US","https://www.makewonder.com/","https://www.makewonder.com/","hello@makewonder.com","★★★★★","5歳以上対応のビジュアルプログラミングロボット。ブロックコーディングで動かせるかわいいロボットで楽しくコーディング学習。"],
  [1663,"子供・教育","Kano PC 2-in-1 Build-A-PC Kit","Kano","Amazon US","https://www.kano.me/","https://www.kano.me/","hello@kano.me","★★★★☆","PC自作キット付きプログラミング学習コンピューター。組み立てから学べる子供向けSTEM特化PC。"],

  // --- ファッション・アクセサリー ---
  [1584,"ファッション・アクセサリー","WANDRD PRVKE 31L V4 Backpack","WANDRD","WANDRD公式","https://www.wandrd.com/collections/backpacks","https://www.wandrd.com/","wholesale@wandrd.com","★★★★☆","カメラバックとしての機能性と普段使いを両立させたプレミアムバックパック。写真家・旅行者・ビジネス層に人気。"],
  [1585,"ファッション・アクセサリー","Db Ramverk Pro 36L Travel Backpack","Db","Db公式","https://dbjourney.com/","https://dbjourney.com/","support@dbjourney.com","★★★★☆","LVMH出資のノルウェー発・機能美と耐久性を兼ね備えたプレミアムトラベルバックパック。品質重視の日本市場向け。"],
  [1587,"ファッション・アクセサリー","Circular Ring 2 Smart Ring - Rose Gold","Circular","Circular公式","https://www.circular.xyz/shop","https://www.circular.xyz/","contact@circular.xyz","★★★★★","ローズゴールドカラーのおしゃれなスマートリング。健康追跡をファッションアクセサリーとして楽しめる製品。"],
  [1588,"ファッション・アクセサリー","iPolish Starter Kit Smart Nail Set","iPolish","iPolish公式","https://ipolish.fashion/collections/starter-kits","https://ipolish.fashion/","CustomerService@ipolish.fashion","★★★★☆","スマートネイルスターターキット。本体・充電器・サンプルネイル付属でスマートネイルを始めるのに最適。"],
  [1664,"ファッション・アクセサリー","Peak Design Travel Backpack 45L","Peak Design","Peak Design公式","https://www.peakdesign.com/","https://www.peakdesign.com/","support@peakdesign.com","★★★★★","撮影・旅行・日常を完全網羅するモジュラーシステム対応プレミアムトラベルバックパック。Kickstarter史上最高額調達の実績。"],
  [1665,"ファッション・アクセサリー","Peak Design Everyday Backpack 20L V2","Peak Design","Peak Design公式","https://www.peakdesign.com/","https://www.peakdesign.com/","support@peakdesign.com","★★★★★","カメラ・PC・日常品を完全収納できる機能美型バックパック。品質とデザインにこだわる日本のユーザーに人気。"],
  [1666,"ファッション・アクセサリー","Bellroy Tokyo Totepack Backpack","Bellroy","Amazon US","https://bellroy.com/","https://bellroy.com/","hello@bellroy.com","★★★★★","オーストラリア発・精密な内部設計と上質素材のプレミアムビジネストートバックパック。日本の通勤ビジネス層に人気。"],
  [1667,"ファッション・アクセサリー","Bellroy Hide & Seek Slim Wallet","Bellroy","Amazon US","https://bellroy.com/","https://bellroy.com/","hello@bellroy.com","★★★★★","スリム設計で大容量収納を実現する本革スマートウォレット。キャッシュレス社会の日本でも根強いウォレット需要に応える。"],
  [1668,"ファッション・アクセサリー","Nomatic Navigator Travel Pack 32L","Nomatic","Nomatic公式","https://www.nomatic.com/","https://www.nomatic.com/","support@nomatic.com","★★★★☆","ビジネス・旅行・日常を完全対応した機能重視型トラベルバックパック。堅牢な設計と豊富な収納が特徴。"],
  [1690,"ファッション・アクセサリー","Peak Design Capture Clip V3 Camera Clip","Peak Design","Amazon US","https://www.peakdesign.com/","https://www.peakdesign.com/","support@peakdesign.com","★★★★★","バックパック・ベルトにカメラを素早くクリップできる革新的カメラクリップシステム。撮影効率を大幅向上。"],
  [1691,"ファッション・アクセサリー","Peak Design Everyday Sling 10L V2","Peak Design","Peak Design公式","https://www.peakdesign.com/","https://www.peakdesign.com/","support@peakdesign.com","★★★★★","カメラ・タブレット・日常品を収納できる機能的スリングバッグ。日本の街歩き・撮影に最適なサイズ感。"],

  // --- クリーニング・収納・整理 ---
  [1589,"クリーニング・収納・整理","Mammotion SPINO S1 Pro Pool Cleaner","Mammotion","Amazon US","https://us.mammotion.com/products/spino-e1-robotic-pool-cleaner","https://us.mammotion.com/","distribution@mammotion.com","★★★★★","自律デッキへの自己昇降機能付き完全ハンズフリーロボットプールクリーナー。CES 2026発表の革新製品。"],
  [1590,"クリーニング・収納・整理","Mammotion LUBA AWD 5000 Robotic Mower","Mammotion","Amazon US","https://us.mammotion.com/","https://us.mammotion.com/","distribution@mammotion.com","★★★★☆","GPSベースの自律型ロボット芝刈り機。ワイヤーレス設置・スマートフォン操作対応。庭付き住宅の日本でも需要あり。"],
  [1591,"クリーニング・収納・整理","Dreame FP10 Self-Cleaning Air Purifier","Dreame","Dreame公式","https://www.dreametech.com/","https://global.dreametech.com/","sales@dreame.tech","★★★★★","ペット毛対応のローラー自動洗浄システム搭載空気清浄機。CES 2026 DREAME NEXT発表の新製品。"],
  [1592,"クリーニング・収納・整理","Dreame MF10 Bladeless Smart Fan","Dreame","Amazon US","https://www.dreametech.com/","https://global.dreametech.com/","sales@dreame.tech","★★★★☆","Amazon新着No.1の羽根なしスマートファン。静音・清潔・スマートコントロールで日本市場に最適。"],
  [1593,"クリーニング・収納・整理","Dreame Halo 3-in-1 Fan Heater Purifier","Dreame","Amazon US","https://www.dreametech.com/","https://global.dreametech.com/","sales@dreame.tech","★★★★★","扇風機・ヒーター・空気清浄機の3機能を一台に統合した多機能家電。省スペース設計で日本の住宅に最適。"],
  [1594,"クリーニング・収納・整理","Tineco FLOOR ONE S9 Steam Wet Dry Vac","Tineco","Amazon US","https://us.tineco.com/","https://us.tineco.com/","marketing@tineco.com","★★★★★","スチームモップ・ウェット・ドライ掃除を一台で実現するスマート床クリーナー。清潔志向の日本市場に最適。"],
  [1595,"クリーニング・収納・整理","Tineco CARPET ONE Smart Carpet Cleaner","Tineco","Amazon US","https://us.tineco.com/","https://us.tineco.com/","marketing@tineco.com","★★★★☆","AIセンサーで汚れを検知し自動で洗浄液を調整するスマートカーペットクリーナー。カーペット家庭向け清掃家電。"],
  [1651,"クリーニング・収納・整理","Lomi Smart Food Waste Composter","Lomi","Amazon US","https://lomi.com/","https://lomi.com/","support@lomi.com","★★★★★","生ゴミを数時間で堆肥に変換するスマートコンポスター。生ゴミ処理問題と環境意識が高い日本市場に最適。"],
  [1653,"クリーニング・収納・整理","Simplehuman Sensor Can Smart Trash Can","Simplehuman","Amazon US","https://www.simplehuman.com/","https://www.simplehuman.com/","support@simplehuman.com","★★★★☆","声・モーションセンサー対応の高品質スマートゴミ箱。耐久性・デザイン性に優れシンプルさにこだわる日本市場向け。"],
  [1654,"クリーニング・収納・整理","Coway Airmega 400S Air Purifier","Coway","Amazon US","https://us.coway.com/","https://us.coway.com/","support@coway.com","★★★★★","大空間対応のWi-Fi接続スマート空気清浄機。花粉症・PM2.5対策が重要な日本で非常に高い需要が見込まれる。"],
  [1655,"クリーニング・収納・整理","IQAir HealthPro Plus Air Purifier","IQAir","IQAir公式","https://www.iqair.com/","https://www.iqair.com/","info@iqair.com","★★★★★","スイス発・世界最高水準のHyperHEPA技術搭載空気清浄機。アレルギー・感染症対策として世界の病院でも採用。"],
  [1656,"クリーニング・収納・整理","BlueAir Blue Pure 311i+ Max","BlueAir","Amazon US","https://www.blueair.com/","https://www.blueair.com/","support@blueair.com","★★★★★","スウェーデン発の高性能HEPASilent空気清浄機。静音・省エネ・洗えるプレフィルター付きで長期使用に経済的。"],
];

// ==================== メイン処理 ====================
console.log('Excelファイル読み込み中...');
const wb = XLSX.readFile(INPUT_FILE);
const ws = wb.Sheets[wb.SheetNames[0]];

// 現在の行数を確認
const range = XLSX.utils.decode_range(ws['!ref']);
const lastRow = range.e.r; // 0-indexed
console.log(`現在の行数: ${lastRow + 1} (ヘッダー含む)`);

// 製品を番号順でソート
NEW_PRODUCTS.sort((a, b) => a[0] - b[0]);
console.log(`追加予定件数: ${NEW_PRODUCTS.length}`);

// 各製品を追加
let added = 0;
for (const prod of NEW_PRODUCTS) {
  const [no, category, productName, makerName, ecSite, prodUrl, makerHp, contact, stars, comment] = prod;

  const rowIdx = lastRow + 1 + added; // 0-indexed

  const setCellVal = (col, val) => {
    const addr = XLSX.utils.encode_cell({ r: rowIdx, c: col });
    ws[addr] = { v: val, t: typeof val === 'number' ? 'n' : 's' };
  };

  setCellVal(0, no);
  setCellVal(1, category);
  setCellVal(2, productName);
  setCellVal(3, makerName);
  setCellVal(4, ecSite);
  setCellVal(5, prodUrl);
  setCellVal(6, makerHp);

  // メールアドレスの場合はGmail URLをハイパーリンクとして設定
  if (contact.includes('@') && !contact.startsWith('http')) {
    const gmailUrl = buildGmailUrl(contact, makerName, productName);
    const addr = XLSX.utils.encode_cell({ r: rowIdx, c: 7 });
    ws[addr] = { v: contact, t: 's', l: { Target: gmailUrl } };
  } else {
    setCellVal(7, contact);
  }

  setCellVal(8, stars);
  setCellVal(9, comment);

  added++;
  if (added % 50 === 0) console.log(`  追加済み: ${added}件`);
}

// レンジを更新
range.e.r = lastRow + added;
ws['!ref'] = XLSX.utils.encode_range(range);

// ファイル保存
console.log('保存中...');
XLSX.writeFile(wb, OUTPUT_FILE);

console.log(`\n完了！`);
console.log(`出力ファイル: ${OUTPUT_FILE}`);
console.log(`追加件数: ${added} 件`);
console.log(`合計件数: ${1505 + added} 件`);
