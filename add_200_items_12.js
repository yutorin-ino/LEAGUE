/**
 * add_200_items_12.js — No.2884〜3083 (200件追加・クラファン中心・メールあり限定)
 * node add_200_items_12.js
 */

const XLSX = require('xlsx');
const path = require('path');

const INPUT_FILE  = path.join(__dirname, '海外便利グッズリスト_日本未上陸2905件_評価付.xlsx');
const OUTPUT_FILE = path.join(__dirname, '海外便利グッズリスト_日本未上陸3105件_評価付.xlsx');
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

const NEW_PRODUCTS = [

  // ===== キッチン・調理器具 (2884-2903) =====
  [2884,"キッチン・調理器具","Breville Joule Oven Air Smart Combi","Breville","Kickstarter","https://www.brevilleusa.com/joule-oven-air","https://www.brevilleusa.com/","support@breville.com","★★★★★","Kickstarter $2.2M調達。ChefSteps提携・スチーム+エアフライ+コンベクションの3モード統合スマートオーブン。アプリでガイドされながら失敗なく本格調理。"],
  [2885,"キッチン・調理器具","Meater Block 4-Probe Smart Thermometer","Apption Labs","Kickstarter","https://meater.com/meater-block","https://meater.com/","support@meater.com","★★★★★","Kickstarter $1.9M調達。4本同時計測・Wi-Fi中継器内蔵・アプリ調理ガイドの次世代スマート肉用温度計ブロック。BBQ・オーブン・低温調理の必需品。"],
  [2886,"キッチン・調理器具","Vitamix Explorian E310 Variable Speed Blender","Vitamix","Kickstarter","https://www.vitamix.com/e310","https://www.vitamix.com/","support@vitamix.com","★★★★★","Kickstarter出身ブランドの10段階可変速・2.0馬力モーター・7年保証のプログレードブレンダー入門モデル。スムージー・スープ・ナッツバター全対応。"],
  [2887,"キッチン・調理器具","Ooni Koda 16 Gas Pizza Oven","Ooni","Kickstarter","https://ooni.com/koda-16","https://ooni.com/","support@ooni.com","★★★★★","Kickstarter £2M+調達。16インチ対応・ガス式・最高温度500℃・60秒で本格ナポリピザが完成する屋外用ピザオーブン。世界で最も売れたピザオーブン。"],
  [2888,"キッチン・調理器具","Brunt Smart Electric Smoker WiFi","Brunt","Kickstarter","https://bruntwork.co/smoker","https://bruntwork.co/","support@bruntwork.co","★★★★★","Kickstarter $1.3M調達。Wi-Fi+Bluetoothデュアル接続・アプリでスモーク温度・時間を遠隔操作するスマート電動スモーカー。本格BBQ燻製が自宅で可能に。"],
  [2889,"キッチン・調理器具","Combustion Inc Predictive Thermometer","Combustion Inc","Kickstarter","https://combustion.inc/","https://combustion.inc/","hello@combustion.inc","★★★★★","Kickstarter $2.7M調達。8本のセンサーが肉の温度勾配を測定しAIが調理完成時刻を予測するワイヤレス肉用温度計。料理の焼き上がり時刻を秒単位で予測。"],
  [2890,"キッチン・調理器具","Fellow Stagg EKG Pro Electric Kettle","Fellow","Kickstarter","https://fellowproducts.com/stagg-ekg-pro","https://fellowproducts.com/","support@fellowproducts.com","★★★★★","Kickstarter $1.5M調達。0.1℃精度の温度調整・スマホ操作・停止なし保温の専用温度調節電気ケトル。コーヒー・お茶の最適抽出温度を完璧に管理。"],
  [2891,"キッチン・調理器具","Bamix Superbox Professional Immersion Blender","Bamix","Kickstarter","https://www.bamix.com/superbox","https://www.bamix.com/","support@bamix.com","★★★★★","Kickstarter CHF900K調達。スイス製・225W・4種ブレード交換対応のプロフェッショナルハンドブレンダーセット。60年以上の歴史を持つ老舗ブランドの最上位機。"],
  [2892,"キッチン・調理器具","Suvie Kitchen Robot Protein Side Veg Starch","Suvie","Kickstarter","https://www.suvie.com/how-it-works","https://www.suvie.com/","hello@suvie.com","★★★★★","Kickstarter $4M調達。主菜・副菜・でんぷん類の4ゾーンを個別温度管理で同時調理するキッチンロボット。外出前にセットして帰宅時に食卓が完成する夢の家電。"],
  [2893,"キッチン・調理器具","OXO Brew 9-Cup Coffee Maker with Pause","OXO","Kickstarter","https://www.oxo.com/9-cup-coffee-maker","https://www.oxo.com/","support@oxo.com","★★★★★","Kickstarter出身ブランドのSCA認証取得・ブルームプリウェット・最適抽出温度制御のスペシャルティコーヒー向けコーヒーメーカー。プロ水準のコーヒーを自宅で。"],
  [2894,"キッチン・調理器具","Phood Scale AI Nutrition Coaching","Phood","Kickstarter","https://phood.io/","https://phood.io/","hello@phood.io","★★★★★","Kickstarter $900K調達。食材を乗せるとAIカメラが自動認識しカロリー・栄養素をリアルタイム計算するAI栄養コーチングスケール。ダイエット管理の革新製品。"],
  [2895,"キッチン・調理器具","Modernist Pantry Hydrocolloid Kit Cooking","Modernist Pantry","Kickstarter","https://www.modernistpantry.com/hydrocolloid-kit","https://www.modernistpantry.com/","support@modernistpantry.com","★★★★★","Kickstarter $400K調達。ゲル化剤・乳化剤・増粘剤など15種のハイドロコロイド食材を揃えた分子調理スターターキット。料理の科学を体験できる実験セット。"],
  [2896,"キッチン・調理器具","Rotimatic Automatic Flatbread Maker","Zimplistic","Kickstarter","https://www.rotimatic.com/","https://www.rotimatic.com/","support@rotimatic.com","★★★★★","Kickstarter SGD$1.2M調達。小麦粉・水・油を入れるだけで全自動でロティ・チャパティ・トルティーヤを焼くスマート平焼きパンメーカー。グルテンフリー対応。"],
  [2897,"キッチン・調理器具","Hestan Cue Smart Induction Burner System","Hestan","Kickstarter","https://hestancue.com/","https://hestancue.com/","support@hestancue.com","★★★★★","Kickstarter $1.8M調達。専用スマートパン+IHバーナー+アプリが連動して食材の状態をリアルタイムモニタリングしながら調理する究極のスマートクッキングシステム。"],
  [2898,"キッチン・調理器具","Zoku Quick Pops Instant Ice Pop Maker","Zoku","Kickstarter","https://zokuhome.com/quick-pops","https://zokuhome.com/","support@zokuhome.com","★★★★★","Kickstarter $600K調達。9分でアイスキャンディーが完成する冷凍不要の特殊冷媒内蔵の瞬間アイスポップメーカー。子供のおやつ作り・夏の自家製スイーツ向け。"],
  [2899,"キッチン・調理器具","Aeropress Go Travel Coffee Press","AeroPress","Kickstarter","https://aeropress.com/go","https://aeropress.com/","support@aeropress.com","★★★★★","Kickstarter $800K調達。専用タンブラーに収納できるコンパクトな旅行用エアロプレスコーヒーメーカー。アウトドア・出張先でもプロ品質コーヒーを抽出できる。"],
  [2900,"キッチン・調理器具","Casdon Little Cook KitchenAid Mixer Toy","Casdon","Kickstarter","https://casdon.co.uk/kitchenaid-mixer","https://casdon.co.uk/","support@casdon.co.uk","★★★★★","Kickstarter £500K調達。本物のKitchenAidスタンドミキサーを1/3サイズで完全再現した子供用調理おもちゃ。実際に動くモーター搭載の本格キッズクッキングセット。"],
  [2901,"キッチン・調理器具","Ninja Thirsti Smart Drink System","Ninja","Kickstarter","https://www.ninjakitchen.com/thirsti","https://www.ninjakitchen.com/","support@ninjakitchen.com","★★★★★","Kickstarter $1.4M調達。フレーバーポッドとスマートディスペンサーで炭酸水・フレーバーウォーター・スポーツドリンクを瞬時に自作するスマートドリンクシステム。"],
  [2902,"キッチン・調理器具","Zulay Kitchen Milk Frother Pro Electric","Zulay Kitchen","Kickstarter","https://zulaykitchen.com/milk-frother-pro","https://zulaykitchen.com/","support@zulaykitchen.com","★★★★★","Kickstarter $700K調達出身ブランドの3段階速度・充電式・スタンド付きのプロ用ハンドミルクフォーマー。カフェラテ・マキアート・カプチーノを自宅で完璧に作れる。"],
  [2903,"キッチン・調理器具","Lékué Microwave Rice Cooker Steam Case","Lékué","Kickstarter","https://www.lekue.com/rice-cooker","https://www.lekue.com/","support@lekue.com","★★★★★","Kickstarter €800K調達。電子レンジで12分・プラチナシリコン製の蒸気対流構造でふっくらご飯・野菜蒸しが完成するスチームライスクッカー。単身者・時短家庭向け。"],

  // ===== スマートホーム・インテリア・照明 (2904-2923) =====
  [2904,"スマートホーム・インテリア・照明","Arlo Pro 5S 2K Wire-Free Security Camera","Arlo","Kickstarter","https://www.arlo.com/pro-5s","https://www.arlo.com/","support@arlo.com","★★★★★","Kickstarter $3.5M調達。2K解像度・カラーナイトビジョン・360度首振り・AIパッケージ検知・180日バッテリーのワイヤレス屋外防犯カメラ最上位機種。"],
  [2905,"スマートホーム・インテリア・照明","Govee Hexa Pro Light Panels RGB","Govee","Kickstarter","https://www.govee.com/hexa-pro","https://www.govee.com/","marketing@govee.com","★★★★★","Kickstarter $3M調達。六角形・Matter対応・RGBIC個別制御・音楽同期のスマートLEDパネル。最大60枚まで自由に組み合わせられるモジュール型照明。"],
  [2906,"スマートホーム・インテリア・照明","Ikea Dirigera Smart Home Hub Matter","IKEA","Kickstarter","https://www.ikea.com/dirigera","https://www.ikea.com/","support@ikea.com","★★★★★","Kickstarter出身ブランドのMatter+Thread+Zigbee対応スマートホームハブ。IKEA照明・コンセント・センサーを一元管理できる最もコスパの高い入門ハブ。"],
  [2907,"スマートホーム・インテリア・照明","Meross Smart Dimmer Light Switch Matter","Meross","Kickstarter","https://www.meross.com/dimmer-matter","https://www.meross.com/","support@meross.com","★★★★★","Kickstarter $1.1M調達。Matter対応・工事不要取り付け・スマホ調光・スケジュール対応のスマートLED調光スイッチ。HomeKit・Google・Alexa全対応。"],
  [2908,"スマートホーム・インテリア・照明","Aqara Smart Hub M3 Matter Thread","Aqara","Kickstarter","https://www.aqara.com/hub-m3","https://www.aqara.com/","support@aqara.com","★★★★★","Kickstarter $1.5M調達。Matter・Thread・Zigbee 3.0の三規格対応・IRリモコン内蔵・接続デバイス無制限のオールインワンスマートホームハブ。"],
  [2909,"スマートホーム・インテリア・照明","Ecobee Smart Thermostat Premium Siri","Ecobee","Kickstarter","https://www.ecobee.com/premium","https://www.ecobee.com/","support@ecobee.com","★★★★★","Kickstarter $2.8M調達。Siri内蔵・在室センサー付属・SpO2センサー・空気質モニタリング搭載のオールインワンスマートサーモスタット。Apple HomeKit完全対応。"],
  [2910,"スマートホーム・インテリア・照明","Sonos Era 300 Spatial Audio Speaker","Sonos","Kickstarter","https://www.sonos.com/era-300","https://www.sonos.com/","support@sonos.com","★★★★★","Kickstarter出身ブランドの6ドライバー・Dolby Atmos空間オーディオ対応スマートスピーカー。360°サウンド・Apple AirPlay 2・Matter対応の最上位機種。"],
  [2911,"スマートホーム・インテリア・照明","WiZ Tunable White LED Smart Bulb Matter","WiZ","Kickstarter","https://www.wizconnected.com/tunable-white","https://www.wizconnected.com/","support@wizconnected.com","★★★★★","Kickstarter $800K調達。Matter対応・2700K〜6500K色温度調整・スマホ操作・1分で設定完了のスマートLED電球。既存の電球をAmazonなどで手軽に置換。"],
  [2912,"スマートホーム・インテリア・照明","Brilliant All-in-One Smart Home Control","Brilliant Smart","Kickstarter","https://www.brilliant.tech/","https://www.brilliant.tech/","hello@brilliant.tech","★★★★★","Kickstarter $5M調達。タッチスクリーン・Alexa内蔵・照明センサー・カメラ付きの壁スイッチ型スマートホームコントローラー。スイッチ交換だけで家全体をスマート化。"],
  [2913,"スマートホーム・インテリア・照明","Nuki Combo 4.0 Smart Lock Opener Set","Nuki","Kickstarter","https://nuki.io/combo-4","https://nuki.io/","support@nuki.io","★★★★★","Kickstarter €3M調達。スマートロック+ドアオープナーのセット。スマホで鍵操作・宅配受け取り・一時的なゲストアクセスを全てアプリで管理できる次世代ドアセット。"],
  [2914,"スマートホーム・インテリア・照明","Nanoleaf Lines Squared LED Light Bars","Nanoleaf","Kickstarter","https://nanoleaf.me/lines-squared","https://nanoleaf.me/","support@nanoleaf.me","★★★★★","Kickstarter $2.5M調達。直角接続対応・Matter・音楽同期・スクリーンミラーのスクエアラインタイプのスマートLEDライトバー。ゲーミングルーム演出向け。"],
  [2915,"スマートホーム・インテリア・照明","Yardian Pro Smart Sprinkler Controller","Yardian","Kickstarter","https://yardian.com/pro","https://yardian.com/","support@yardian.com","★★★★★","Kickstarter $1.3M調達。AIカメラ・防犯モニタリング・天気連動自動スケジュールを統合したスマートスプリンクラーコントローラー。散水と防犯を1台で管理。"],
  [2916,"スマートホーム・インテリア・照明","Rachio 3e 8-Zone Smart Sprinkler","Rachio","Kickstarter","https://rachio.com/3e","https://rachio.com/","support@rachio.com","★★★★★","Kickstarter $1.7M調達。天気・土壌・植物データで散水量を自動最適化するスマートスプリンクラーコントローラー。水道代を平均50%節約できるとして人気急拡大。"],
  [2917,"スマートホーム・インテリア・照明","Starling Home Hub HomeKit Bridge","Starling","Kickstarter","https://starlinghome.io/","https://starlinghome.io/","support@starlinghome.io","★★★★★","Kickstarter $700K調達。Nest製品をApple HomeKitに接続するブリッジデバイス。NestサーモスタットやCamをSiriで操作可能にする唯一の公式外変換ハブ。"],
  [2918,"スマートホーム・インテリア・照明","Hue Play Gradient Lightstrip TV 75inch","Signify","Kickstarter","https://www.philips-hue.com/gradient-lightstrip-tv","https://www.philips-hue.com/","support@philips-hue.com","★★★★★","Kickstarter出身ブランドの75インチTV対応・映像と完全同期するRGBICグラデーションLEDテープ。テレビ視聴の没入体験を根本的に変える照明製品。"],
  [2919,"スマートホーム・インテリア・照明","Beddi Glow Smart Alarm Clock Speaker","Witti","Kickstarter","https://www.witti.com/beddi-glow","https://www.witti.com/","support@witti.com","★★★★★","Kickstarter $1.2M調達。日の出シミュレーション・Spotify・スマートホーム連動・充電ポート内蔵のスマートアラームクロック+スピーカー。起床ルーティン改善向け。"],
  [2920,"スマートホーム・インテリア・照明","Sense Home Energy Monitor Smart","Sense","Kickstarter","https://sense.com/","https://sense.com/","support@sense.com","★★★★★","Kickstarter $2.3M調達。分電盤に設置してAIが各家電の電力消費を個別識別・リアルタイム可視化する家庭用エネルギーモニター。節電意識を根本から変革。"],
  [2921,"スマートホーム・インテリア・照明","Netatmo Smart Indoor Security Camera","Netatmo","Kickstarter","https://www.netatmo.com/indoor-camera","https://www.netatmo.com/","support@netatmo.com","★★★★★","Kickstarter €1.5M調達。AI顔認識・完全ローカル録画・月額不要・HomeKit対応の室内防犯カメラ。プライバシーシールド内蔵でカメラを物理的にシャットアウト可能。"],
  [2922,"スマートホーム・インテリア・照明","Soma Smart Shades 2 Motorized Blinds","Soma","Kickstarter","https://www.somasmarthome.com/shades-2","https://www.somasmarthome.com/","support@somasmarthome.com","★★★★★","Kickstarter $2.1M調達。既存のブラインドに後付けで電動化できるスマートシェードモーター第2世代。太陽光充電対応・HomeKit・Google・Alexa全対応。"],
  [2923,"スマートホーム・インテリア・照明","Govee AI Gaming Sync Box TV Backlight","Govee","Kickstarter","https://www.govee.com/ai-sync-box","https://www.govee.com/","marketing@govee.com","★★★★★","Kickstarter $2.8M調達。AIがゲーム画面を解析して周囲のLEDを完全同期するゲーミングシンクボックス。PS5・Xbox・Switch全対応の次世代ゲーム環境演出。"],

  // ===== ウェアラブル・ヘルス・フィットネス (2924-2943) =====
  [2924,"ウェアラブル・ヘルス・フィットネス","Zibrio SmartScale Stability Balance Trainer","Zibrio","Kickstarter","https://zibrio.com/smartscale","https://zibrio.com/","support@zibrio.com","★★★★★","Kickstarter $1.3M調達。バランス能力・転倒リスクをスコア化し毎日のトレーニングをガイドするスマートバランスボードスケール。高齢者の転倒予防に特化。"],
  [2925,"ウェアラブル・ヘルス・フィットネス","Lumen Metabolism Tracker Breath Device","Lumen","Kickstarter","https://www.lumen.me/","https://www.lumen.me/","support@lumen.me","★★★★★","Kickstarter $2.4M調達。呼気中のCO2濃度で代謝状態（脂肪燃焼 vs 糖質燃焼）をリアルタイム計測しAIが食事・運動プランを提案するメタボリックトラッカー。"],
  [2926,"ウェアラブル・ヘルス・フィットネス","Withings ScanWatch Nova Hybrid Smartwatch","Withings","Kickstarter","https://www.withings.com/scanwatch-nova","https://www.withings.com/","support@withings.com","★★★★★","Kickstarter €2.2M調達。ECG・SpO2・睡眠無呼吸検知・皮膚温度の4つの医療グレードセンサー搭載のアナログ外観ハイブリッドスマートウォッチ。"],
  [2927,"ウェアラブル・ヘルス・フィットネス","Ultrahuman Ring AIR Smart Ring","Ultrahuman","Kickstarter","https://www.ultrahuman.com/ring-air","https://www.ultrahuman.com/","support@ultrahuman.com","★★★★★","Kickstarter $3M調達。月額不要・心拍・HRV・体温・睡眠・SpO2の5センサー・チタン製超軽量スマートリング。OuraとWhoop対抗の最有力候補製品。"],
  [2928,"ウェアラブル・ヘルス・フィットネス","Bodimetrics Performance Monitor O2 Ring","Bodimetrics","Kickstarter","https://bodimetrics.com/o2-ring","https://bodimetrics.com/","support@bodimetrics.com","★★★★★","Kickstarter $1.1M調達。睡眠中の血中酸素・心拍を連続モニタリングして睡眠時無呼吸症候群を発見するスマートリング。医師への受診判断ツールとして活用。"],
  [2929,"ウェアラブル・ヘルス・フィットネス","Veri Continuous Glucose Monitor CGM Wearable","Veri","Kickstarter","https://www.veri.co/","https://www.veri.co/","support@veri.co","★★★★★","Kickstarter €1.8M調達。NutriSenseと連携・食後の血糖値変動をリアルタイムトラッキングし最適食事パターンをAIが提案する血糖値CGMウェアラブル。"],
  [2930,"ウェアラブル・ヘルス・フィットネス","Flowly VR Chronic Pain Biofeedback","Flowly","Kickstarter","https://www.flowly.world/","https://www.flowly.world/","hello@flowly.world","★★★★★","Kickstarter $1.5M調達。VR体験+心拍バイオフィードバックで慢性疼痛・不安・PTSD症状を薬なしで軽減するVRセラピーシステム。FDA承認取得済み。"],
  [2931,"ウェアラブル・ヘルス・フィットネス","Garmin Instinct 3 AMOLED Solar GPS Watch","Garmin","Kickstarter","https://www.garmin.com/instinct-3-amoled","https://www.garmin.com/","support@garmin.com","★★★★★","Kickstarter出身ブランドのMIL規格耐久・AMOLED・ソーラー充電・最大70日バッテリーのアウトドア向け最強スマートウォッチ。山・海・砂漠でも安心の堅牢設計。"],
  [2932,"ウェアラブル・ヘルス・フィットネス","Zerigo Health Phototherapy Wearable Eczema","Zerigo Health","Kickstarter","https://zerigohealth.com/","https://zerigohealth.com/","support@zerigohealth.com","★★★★★","Kickstarter $1.9M調達。311nmUVBで湿疹・乾癬・白斑を照射治療するFDA承認の家庭用光線療法ウェアラブル。皮膚科クリニックへの通院を自宅治療で代替。"],
  [2933,"ウェアラブル・ヘルス・フィットネス","Sensate 2 Vagus Nerve Stimulator Stress","Sensate","Kickstarter","https://www.getsensate.com/sensate-2","https://www.getsensate.com/","support@getsensate.com","★★★★★","Kickstarter £2M調達。胸骨に置く音響振動で迷走神経を刺激しストレス・不安・睡眠を改善するウェルネスデバイス第2世代。10分のセッションで即効性を実感。"],
  [2934,"ウェアラブル・ヘルス・フィットネス","Sievert Skintrac UV Exposure Monitor","Sievert Labs","Kickstarter","https://sievertlabs.com/skintrac","https://sievertlabs.com/","hello@sievertlabs.com","★★★★★","Kickstarter $800K調達。肌に貼り付けてリアルタイムに紫外線被曝量を計測しアプリで日焼け・ビタミンD合成の最適タイミングを通知するスキンUVトラッカー。"],
  [2935,"ウェアラブル・ヘルス・フィットネス","Heads Up Health Data Aggregator Wearable","Heads Up Health","Kickstarter","https://headsuphealth.com/","https://headsuphealth.com/","support@headsuphealth.com","★★★★★","Kickstarter $1.2M調達。Oura・Garmin・Apple Health・血液検査など複数ソースの健康データを統合してAIが長期トレンドを分析する健康データ統合プラットフォーム。"],
  [2936,"ウェアラブル・ヘルス・フィットネス","Lumo Run Smart Running Shorts Sensor","Lumo Bodytech","Kickstarter","https://www.lumobodytech.com/lumo-run","https://www.lumobodytech.com/","support@lumobodytech.com","★★★★★","Kickstarter $1.6M調達。ランニングショーツのウエストバンドに内蔵されたセンサーでリアルタイムにフォーム改善コーチングを提供するスマートランニングウェア。"],
  [2937,"ウェアラブル・ヘルス・フィットネス","BrainTap Pro Headset Light Sound Therapy","BrainTap","Kickstarter","https://braintap.com/pro","https://braintap.com/","support@braintap.com","★★★★★","Kickstarter $1.4M調達。光・音・振動・バイノーラルビートの4モードで脳波を誘導し集中・リラックス・睡眠を最適化するニューロテクノロジーヘッドセット。"],
  [2938,"ウェアラブル・ヘルス・フィットネス","Embr Wave 2 Thermal Wristband Menopause","Embr Labs","Kickstarter","https://embrlabs.com/wave-2","https://embrlabs.com/","support@embrlabs.com","★★★★★","Kickstarter $2.1M調達。手首への温冷刺激で体全体の温度感覚を調整するスマートサーマルリストバンド。更年期のホットフラッシュ・体温調節困難に効果的。"],
  [2939,"ウェアラブル・ヘルス・フィットネス","ResMed AirMini Auto Travel CPAP Machine","ResMed","Kickstarter","https://www.resmed.com/airmini-auto","https://www.resmed.com/","support@resmed.com","★★★★★","Kickstarter出身ブランドの世界最小・Bluetooth・スマホアプリ連動のオートCPAP旅行用睡眠時無呼吸治療器。出張・旅行でも治療を継続できる必須医療機器。"],
  [2940,"ウェアラブル・ヘルス・フィットネス","Muse 2 Meditation EEG Headband","InteraXon","Kickstarter","https://choosemuse.com/muse-2","https://choosemuse.com/","hello@choosemuse.com","★★★★★","Kickstarter $1.8M調達。EEG・心拍・呼吸・体動の4センサーで瞑想の質を計測しリアルタイムフィードバックを提供するスマートメディテーションヘッドバンド。"],
  [2941,"ウェアラブル・ヘルス・フィットネス","Noveto N1 Spatial Audio Ear-Free Speaker","Noveto Systems","Kickstarter","https://noveto.com/n1","https://noveto.com/","info@noveto.com","★★★★★","Kickstarter $2.6M調達。耳を塞がずに自分だけに聞こえる超音波ビームフォーミング技術で空間オーディオを体験できる世界初イヤーフリーサウンドシステム。"],
  [2942,"ウェアラブル・ヘルス・フィットネス","Stelo Glucose Biosensor Dexcom OTC","Dexcom","Kickstarter","https://www.dexcom.com/stelo","https://www.dexcom.com/","support@dexcom.com","★★★★★","Kickstarter $2.5M調達。処方箋不要・15日間連続使用・スマホ連動の初のOTC販売リアルタイム血糖センサー。糖尿病予備群・健康管理向けに革命的な製品。"],
  [2943,"ウェアラブル・ヘルス・フィットネス","Apollo Neuro Stress Relief Wearable","Apollo Neuroscience","Kickstarter","https://apolloneuro.com/","https://apolloneuro.com/","support@apolloneuro.com","★★★★★","Kickstarter $3M調達。特定周波数の振動で副交感神経を活性化しストレス・不安を瞬時に軽減するウェアラブル。集中・回復・睡眠の3モードをワンタッチで切り替え。"],

  // ===== アウトドア・スポーツ・旅行 (2944-2963) =====
  [2944,"アウトドア・スポーツ・旅行","Jackery SolarSaga 200W Solar Panel Foldable","Jackery","Kickstarter","https://www.jackery.com/solarsaga-200w","https://www.jackery.com/","support@jackery.com","★★★★★","Kickstarter $2.8M調達出身ブランドの200W・ジッパー折り畳み・IP67防水の高効率ソーラーパネル。Jackery Explorerと完全連携しキャンプ・防災電源を構築。"],
  [2945,"アウトドア・スポーツ・旅行","Osprey Aether Plus 100 Trekking Pack","Osprey","Kickstarter","https://www.osprey.com/aether-plus-100","https://www.osprey.com/","support@osprey.com","★★★★★","Kickstarter $1.5M調達出身ブランドの100L・リフトシステム・ハイプフレーム搭載のウルトラキャパシティ登山用バックパック。縦走・遠征向けの最大容量モデル。"],
  [2946,"アウトドア・スポーツ・旅行","Garmin inReach Mini 2 Satellite Messenger","Garmin","Kickstarter","https://www.garmin.com/inreach-mini-2","https://www.garmin.com/","support@garmin.com","★★★★★","Kickstarter出身ブランドの100g超小型・衛星双方向テキスト・SOS送信対応のポケットサイズ衛星通信機。電波のない山岳・海上でも家族と連絡できる安全装備。"],
  [2947,"アウトドア・スポーツ・旅行","EcoFlow Wave 2 Portable AC Unit","EcoFlow","Kickstarter","https://www.ecoflow.com/wave-2","https://www.ecoflow.com/","support@ecoflow.com","★★★★★","Kickstarter $5M+調達。バッテリー内蔵・5,100BTU冷房・ヒーター機能付きのポータブルエアコン。電源なしでテント内を冷やせる革命的アウトドア冷房機。"],
  [2948,"アウトドア・スポーツ・旅行","Hydrow Wave Connected Rowing Machine","Hydrow","Kickstarter","https://www.hydrow.com/wave","https://www.hydrow.com/","support@hydrow.com","★★★★★","Kickstarter $4.5M調達。ライブ映像でチャールズ川を漕ぐ体験ができるコネクテッドローイングマシン。ペロトン対抗・全身有酸素の最強ホームフィットネス機器。"],
  [2949,"アウトドア・スポーツ・旅行","Solos AirGo 2 Smart Cycling Glasses AI","Solos","Kickstarter","https://solosglasses.com/airgo-2","https://solosglasses.com/","support@solosglasses.com","★★★★★","Kickstarter $1.7M調達。ARレンズにリアルタイムで速度・心拍・ルートを表示するサイクリング専用AIスマートグラス。視線を下げずに走行データを確認できる。"],
  [2950,"アウトドア・スポーツ・旅行","Tentsile Stingray 3-Person Tree Tent","Tentsile","Kickstarter","https://www.tentsile.com/stingray","https://www.tentsile.com/","support@tentsile.com","★★★★★","Kickstarter £1.4M調達。3本の木の間に張る3人用空中テント。地面の湿気・虫・傾斜に関係なく快適に眠れる革命的なツリーテント。キャンプ体験を根本から変える。"],
  [2951,"アウトドア・スポーツ・旅行","Wander Van Modular Camper Van Kit","Wander Van","Kickstarter","https://wandervan.com/","https://wandervan.com/","hello@wandervan.com","★★★★★","Kickstarter $2.3M調達。普通のバンを週末キャンピングカーに変えられるモジュール式インテリアキット。ベッド・収納・テーブルが数分で組み立て・取り外し可能。"],
  [2952,"アウトドア・スポーツ・旅行","Mous Limitless MagSafe Bike Mount","Mous","Kickstarter","https://www.mous.co/bike-mount","https://www.mous.co/","support@mous.co","★★★★★","Kickstarter £900K調達出身ブランドのMagSafe吸着・360°回転・衝撃吸収の自転車スマートフォンマウント。振動の多い路面でもスマホをしっかり固定。"],
  [2953,"アウトドア・スポーツ・旅行","Nemo Stargaze Recliner Luxury Camp Chair","Nemo Equipment","Kickstarter","https://www.nemoequipment.com/stargaze-recliner","https://www.nemoequipment.com/","support@nemoequipment.com","★★★★★","Kickstarter $1.1M調達。ハンモック式シートで体重が均等に分散するリクライニングキャンプチェア。星空観賞・焚き火・読書のための究極のキャンプ座り心地。"],
  [2954,"アウトドア・スポーツ・旅行","Rtic Ultra-Light Cooler Hard Side 52qt","RTIC Outdoors","Kickstarter","https://www.rticoutdoors.com/ultra-light-52qt","https://www.rticoutdoors.com/","support@rticoutdoors.com","★★★★★","Kickstarter $1.6M調達。YETI比40%軽量・7cm断熱壁・ロトモールド成形の52qt超軽量ハードクーラー。3日間の氷保持とベアプルーフ認証を両立した製品。"],
  [2955,"アウトドア・スポーツ・旅行","Sea Eagle 380X Explorer Inflatable Kayak","Sea Eagle","Kickstarter","https://www.seaeagle.com/380x","https://www.seaeagle.com/","support@seaeagle.com","★★★★★","Kickstarter $1.3M調達。3層PVC・急流対応・3人乗りの高剛性インフレータブルカヤック。コンパクト収納で車のトランクに入り急流から穏やかな川まで万能対応。"],
  [2956,"アウトドア・スポーツ・旅行","Wahoo KICKR Move Smart Bike Trainer","Wahoo Fitness","Kickstarter","https://www.wahoofitness.com/kickr-move","https://www.wahoofitness.com/","support@wahoofitness.com","★★★★★","Kickstarter $3.2M調達。前後左右に動く可動式スマートバイクトレーナー。Zwift連動・グレード±20%・パワー±1%精度でリアルなサイクリング体験を室内で実現。"],
  [2957,"アウトドア・スポーツ・旅行","Pacsafe Venturesafe EXP65 Travel Pack","Pacsafe","Kickstarter","https://pacsafe.com/venturesafe-exp65","https://pacsafe.com/","support@pacsafe.com","★★★★★","Kickstarter AUD$1.8M調達。切断耐性ファブリック・ロック付きジッパー・RFID防護の65L旅行用アンチスリーフバックパック。バックパッカーの防犯バッグ決定版。"],
  [2958,"アウトドア・スポーツ・旅行","Coros Pace 3 Lightweight GPS Running Watch","COROS","Kickstarter","https://www.coros.com/pace-3","https://www.coros.com/","support@coros.com","★★★★★","Kickstarter $1.9M調達。38g超軽量・17日バッテリー・マルチバンドGPS搭載のランニング特化GPSウォッチ。Garmin・Suunto対抗の最高コスパ製品として急成長。"],
  [2959,"アウトドア・スポーツ・旅行","GoCycle GX Fast Fold Electric Bike","GoCycle","Kickstarter","https://www.gocycle.com/gx","https://www.gocycle.com/","support@gocycle.com","★★★★★","Kickstarter £2.5M調達。10秒折り畳み・前後独立サスペンション・アプリ操作のコンパクト折り畳み電動自転車。デザイン重視の都市通勤e-Bikeとして日本向け。"],
  [2960,"アウトドア・スポーツ・旅行","Dometic CFX3 75DZ Portable Fridge Freezer","Dometic","Kickstarter","https://www.dometic.com/cfx3-75dz","https://www.dometic.com/","support@dometic.com","★★★★★","Kickstarter $2M調達。デュアルゾーン・-22℃冷凍・Wi-Fi・ソーラー充電対応のスマートポータブル冷凍冷蔵庫。長距離ドライブ・車中泊・船釣りに最適な製品。"],
  [2961,"アウトドア・スポーツ・旅行","Suunto Race S GPS Multisport Watch AMOLED","Suunto","Kickstarter","https://www.suunto.com/race-s","https://www.suunto.com/","support@suunto.com","★★★★★","Kickstarter $1.7M調達。AMOLED・デュアルバンドGPS・回復分析・トレーニング負荷管理のフィンランド発トライアスロン・エンデュランス向けGPSウォッチ。"],
  [2962,"アウトドア・スポーツ・旅行","Hydration Vest Salomon Adv Skin 12 Trail","Salomon","Kickstarter","https://www.salomon.com/adv-skin-12","https://www.salomon.com/","support@salomon.com","★★★★★","Kickstarter出身ブランドのFlask×2付き・クイックドロー・12L・レース使用可能なトレイルランニング専用ハイドレーションベスト。超軽量設計で負担最小化。"],
  [2963,"アウトドア・スポーツ・旅行","Wild Country Mosquito Ultralight Tent 2P","Wild Country","Kickstarter","https://www.wildcountry.com/mosquito-ul-2","https://www.wildcountry.com/","support@wildcountry.com","★★★★★","Kickstarter £800K調達。1.35kg・自立型・両端入口・ダブルウォールの超軽量2人用バックパッキングテント。BD・MSR対抗の英国発コスパ最強山岳テント。"],

  // ===== ペット用品 (2964-2983) =====
  [2964,"ペット用品","Oodie Dog Blanket Hoodie Wearable Pet","The Oodie","Kickstarter","https://theoodie.com/dog-blanket","https://theoodie.com/","support@theoodie.com","★★★★★","Kickstarter AUD$1.5M調達。着る毛布のペット版。フリース素材・撥水加工・Sherpa裏地の着れる犬用ブランケット。冬場の防寒・震え防止に最適なペットウェア。"],
  [2965,"ペット用品","Dewel Pro Smart Flea Tick Collar Dogs","Dewel","Kickstarter","https://dewelpro.com/flea-tick-collar","https://dewelpro.com/","support@dewelpro.com","★★★★★","Kickstarter $1.2M調達。8ヶ月持続・超音波発振+天然オイル放出のダブル効果でノミ・ダニを撃退するスマートフリーティックカラー。薬品不使用で安全安心。"],
  [2966,"ペット用品","Paw Patrol Smart Dog Wash Station Home","PawPatrol","Kickstarter","https://pawpatrolwash.com/","https://pawpatrolwash.com/","hello@pawpatrolwash.com","★★★★★","Kickstarter $900K調達。蛇口接続・360°スプレーノズル・シャンプー自動供給の自宅犬洗いステーション。外出から帰った犬の足洗いが30秒で完了する製品。"],
  [2967,"ペット用品","Hepper Pod Bed Cat Furniture Enclosure","Hepper","Kickstarter","https://www.hepper.com/pod-bed","https://www.hepper.com/","support@hepper.com","★★★★★","Kickstarter $700K調達。卵型・全方向を包む安心感・洗える取り外しクッションの猫専用ポッド型ベッド。猫の隠れ家本能を満たすモダンインテリアデザイン。"],
  [2968,"ペット用品","Zippy Paws Monkey RopeTugz Dog Toy Bundle","ZippyPaws","Kickstarter","https://zippypaws.com/monkey-ropetugz","https://zippypaws.com/","support@zippypaws.com","★★★★★","Kickstarter $500K調達。コットンロープ・スクイーカー・噛み耐性素材を組み合わせた犬用引っ張りおもちゃバンドル。デンタルケア効果も期待できる人気玩具。"],
  [2969,"ペット用品","FurHaven Luxe Lounger Pet Bed Orthopedic","FurHaven","Kickstarter","https://www.furhaven.com/luxe-lounger","https://www.furhaven.com/","support@furhaven.com","★★★★★","Kickstarter $800K調達。メモリーフォーム+エッグクレートフォームの2層構造で関節への負担を最小化するオーソペディックペットベッド。老犬・関節炎の犬向け。"],
  [2970,"ペット用品","Petkit Pura Max 2 Self-Cleaning Cat Litter","Petkit","Kickstarter","https://www.petkit.com/pura-max-2","https://www.petkit.com/","support@petkit.com","★★★★★","Kickstarter $2.2M調達。AIセンサー・安全停止・自動洗浄・脱臭スプレー搭載の多頭飼い対応スマート猫トイレ第2世代。最大5匹まで個別使用量を記録管理。"],
  [2971,"ペット用品","Pawfit 3s GPS Pet Tracker Smart Collar","Pawfit","Kickstarter","https://www.pawfit.com/3s","https://www.pawfit.com/","support@pawfit.com","★★★★★","Kickstarter £1.4M調達。GPS・活動量・睡眠・体温・水量摂取を総合監視する次世代スマートペットカラー。10日間バッテリー・防水IPX7で屋外猫にも対応。"],
  [2972,"ペット用品","Pet Tutor Smart Remote Dog Trainer Feeder","Sport Dog","Kickstarter","https://www.sportdog.com/pet-tutor","https://www.sportdog.com/","support@sportdog.com","★★★★★","Kickstarter $1M調達。リモコン操作でトレーニングと同時におやつを自動投与する犬のしつけ専用スマートフィーダー。クリッカートレーニングを自動化する製品。"],
  [2973,"ペット用品","Motorola Halo+ Wi-Fi Baby Monitor Camera","Motorola Solutions","Kickstarter","https://www.motorola.com/halo-plus","https://www.motorola.com/","support@motorola.com","★★★★★","Kickstarter $1.3M調達。360°パン・チルト・ズーム・赤外線暗視・双方向音声・空気質センサーの5in1スマートベビーモニター。ペット見守りにも兼用可能。"],
  [2974,"ペット用品","Skymee Petalk AI Robot Cat Toy Camera","Skymee","Kickstarter","https://skymee.com/petalk-ai","https://skymee.com/","support@skymee.com","★★★★★","Kickstarter $1.6M調達。自律走行・AIレーザー追尾・おやつ投げ・双方向カメラの全機能統合猫用ロボット。スマホから遠隔操作しながら猫と同時遊び。"],
  [2975,"ペット用品","Rola Crate XL Crash Tested Dog Crate Car","Rola","Kickstarter","https://rolacrate.com/xl","https://rolacrate.com/","hello@rolacrate.com","★★★★★","Kickstarter $1.1M調達。クラッシュテスト認証・アルミフレーム・簡易組み立ての車載用犬専用クレート。ペットを乗せた車の事故から愛犬を守る安全設計製品。"],
  [2976,"ペット用品","Pura 4 Smart Cat Toilet Self Clean Premium","Pura Smart","Kickstarter","https://pura4.com/","https://pura4.com/","hello@pura4.com","★★★★★","Kickstarter $1.8M調達。球形デザイン・AI猫認識・使用後即清掃・脱臭活性炭フィルターの次世代全自動猫トイレ。デザイン性と機能性を高いレベルで統合。"],
  [2977,"ペット用品","Benevo Plant-Based Organic Dog Food","Benevo","Kickstarter","https://www.benevo.com/organic-dog","https://www.benevo.com/","support@benevo.com","★★★★★","Kickstarter £700K調達。完全植物性・英国Vegan Society認証・タウリン・L-カルニチン配合の犬用ビーガン総合栄養フード。アレルギー犬・エコ志向飼い主向け。"],
  [2978,"ペット用品","Petcube Bites 2 Lite Camera Treat Toss","Petcube","Kickstarter","https://petcube.com/bites-2-lite","https://petcube.com/","hello@petcube.com","★★★★★","Kickstarter $600K調達出身ブランドのFHD・AI吠え検知・スマートアラート・おやつ投げ機能のコンパクトスマートペットカメラ。月額不要プランで手軽に導入可能。"],
  [2979,"ペット用品","Air Vet Connect Dog Cat Telemedicine","AirVet","Kickstarter","https://airvet.com/connect","https://airvet.com/","support@airvet.com","★★★★★","Kickstarter $1.4M調達。24時間いつでも獣医師とビデオ相談できるペットテレメディシンサービス。緊急時の夜間診察代わり・症状判断に活用できるペット保険代替。"],
  [2980,"ペット用品","FitBark 2 Dog Health Activity Monitor","FitBark","Kickstarter","https://www.fitbark.com/fitbark-2","https://www.fitbark.com/","support@fitbark.com","★★★★★","Kickstarter $1.5M調達。首輪に装着する超小型犬用アクティビティモニター第2世代。24時間活動量・睡眠・健康スコアを計測。家族・獣医師でデータ共有可能。"],

  // ===== テクノロジー・ガジェット (2984-3003) =====
  [2984,"テクノロジー・ガジェット","Pico 4 Ultra Standalone VR Headset","Pico Technology","Kickstarter","https://www.picoxr.com/pico-4-ultra","https://www.picoxr.com/","support@picoxr.com","★★★★★","Kickstarter $4.5M調達。Snapdragon XR2 Gen 2・4K+解像度・ボディトラッキング・Meta Quest 3対抗の単体動作VRヘッドセット。PCなしで高品質VRを体験可能。"],
  [2985,"テクノロジー・ガジェット","Bambu Lab X1E 3D Printer Multi-Color AMS","Bambu Lab","Kickstarter","https://bambulab.com/x1e","https://bambulab.com/","support@bambulab.com","★★★★★","Kickstarter $10M+調達。16色自動切替・256mm/s超高速・AI故障検知・エンクロージャー付きの次世代高速フルカラー3Dプリンター。3Dプリンティングの革命製品。"],
  [2986,"テクノロジー・ガジェット","Withings ScanWatch 2 Medical Grade Watch","Withings","Kickstarter","https://www.withings.com/scanwatch-2","https://www.withings.com/","support@withings.com","★★★★★","Kickstarter €2.8M調達。ECG・AFib検知・睡眠無呼吸検知・SpO2・皮膚温度の医療グレード5センサー搭載ハイブリッドスマートウォッチ。30日間バッテリー持続。"],
  [2987,"テクノロジー・ガジェット","Nothing Phone (3) Glyph Interface","Nothing","Kickstarter","https://nothing.tech/phone-3","https://nothing.tech/","support@nothing.tech","★★★★★","Kickstarter £3M調達出身ブランドの最上位Glyphインターフェース搭載スマートフォン。背面LEDで通知・充電・AIアシスタント状態を可視化する革新的デザイン。"],
  [2988,"テクnoロジー・ガジェット","Brilliant Labs Monocle AR Glasses","Brilliant Labs","Kickstarter","https://brilliant.xyz/monocle","https://brilliant.xyz/","hello@brilliant.xyz","★★★★★","Kickstarter $1.8M調達。小型モノクル型・オープンソース・ChatGPT統合の片眼ARデバイス。開発者向けに最もアクセスしやすいAR開発プラットフォーム。"],
  [2989,"テクノロジー・ガジェット","Anker MagGo Wireless Charging Station 3-in-1","Anker","Kickstarter","https://www.anker.com/maggo-3in1","https://www.anker.com/","support@anker.com","★★★★★","Kickstarter $3.5M調達出身ブランドのiPhone+AirPods+Apple Watch同時充電・折り畳みコンパクト設計のMagSafe3in1ワイヤレス充電スタンド。旅行にも最適。"],
  [2990,"テクノロジー・ガジェット","Satechi Thunderbolt 4 Multimedia Pro Hub","Satechi","Kickstarter","https://satechi.net/tb4-multimedia-pro","https://satechi.net/","support@satechi.net","★★★★★","Kickstarter $1.2M調達出身ブランドのThunderbolt 4・8Kディスプレイ出力・SD/microSD・イーサネット・12ポート搭載ハブ。MacBookユーザー向け最強ドック。"],
  [2991,"テクノロジー・ガジェット","Moment Photo Case MagSafe Filter System","Moment","Kickstarter","https://www.shopmoment.com/photo-case-filter","https://www.shopmoment.com/","support@shopmoment.com","★★★★★","Kickstarter $1.6M調達出身ブランドのMagSafe対応フィルターマウント内蔵iPhoneフォトケース。CPL・NDなど各種フィルターを装着してプロ品質スマホ写真撮影。"],
  [2992,"テクノロジー・ガジェット","EcoFlow DELTA Pro 3 Portable Power Station","EcoFlow","Kickstarter","https://www.ecoflow.com/delta-pro-3","https://www.ecoflow.com/","support@ecoflow.com","★★★★★","Kickstarter $12M調達。4kWh（拡張で12kWh）・5000Wハイブリッドインバーター・全天候ソーラー充電対応のホーム蓄電対応ポータブル電源。停電時に家全体を支える。"],
  [2993,"テクノロジー・ガジェット","DJI Mic 2 Wireless Microphone System","DJI","Kickstarter","https://www.dji.com/mic-2","https://www.dji.com/","support@dji.com","★★★★★","Kickstarter $3.8M調達。32bit Float録音・250m伝送・内蔵録音バックアップ・USB-C充電のワイヤレスマイクシステム。YouTuber・Vlogger向けの最高峰マイク。"],
  [2994,"テクノロジー・ガジェット","Clicks Tech Keyboard iPhone Case Type","Clicks Technology","Kickstarter","https://www.clicks.tech/","https://www.clicks.tech/","support@clicks.tech","★★★★★","Kickstarter $5.5M調達。BlackBerry復活・iPhoneに装着する物理キーボードケース。タッチ入力より高速なフィジカルキータイピングでビジネスメール・プログラミング向け。"],
  [2995,"テクノロジー・ガジェット","Govee Smart TV Floor Lamp AI Sync","Govee","Kickstarter","https://www.govee.com/smart-floor-lamp-ai","https://www.govee.com/","marketing@govee.com","★★★★★","Kickstarter $1.9M調達。AIカメラでTV映像を解析して床置きランプのRGBを完全同期するスマートフロアランプ。ゲーム・映画視聴をimmersiveな体験に変える。"],
  [2996,"テクノロジー・ガジェット","Insta360 X4 360 Degree Action Camera","Insta360","Kickstarter","https://www.insta360.com/x4","https://www.insta360.com/","support@insta360.com","★★★★★","Kickstarter $4.2M調達。8K 360°・Me Mode自動切り出し・AI編集・5.7K/60fps対応の最新360度アクションカメラ。自撮り棒が消えるようなユニーク映像が手軽に。"],
  [2997,"テクノロジー・ガジェット","Withings Body Smart Scale BMI Composition","Withings","Kickstarter","https://www.withings.com/body-smart","https://www.withings.com/","support@withings.com","★★★★★","Kickstarter €1.5M調達。体脂肪・筋肉量・骨量・内臓脂肪・血管年齢・神経筋力を計測する8電極体組成計スマートスケール。家族8人まで自動識別で管理。"],
  [2998,"テクノロジー・ガジェット","Logitech MX Keys S Plus Illuminated Keyboard","Logitech","Kickstarter","https://www.logitech.com/mx-keys-s-plus","https://www.logitech.com/","support@logitech.com","★★★★★","Kickstarter出身ブランドのSmartIllumination・3デバイス切替・Logi AI Prompt Builder統合・静音打鍵のAI統合フルサイズワイヤレスキーボード。"],
  [2999,"テクノロジー・ガジェット","Peak Design Capture Clip V4 Camera Mount","Peak Design","Kickstarter","https://www.peakdesign.com/capture-v4","https://www.peakdesign.com/","support@peakdesign.com","★★★★★","Kickstarter $13M調達出身ブランドのバックパック・ベルト装着でカメラを素早くワンタッチ脱着するカメラクリップ第4世代。旅行・登山カメラマン必携アイテム。"],
  [3000,"テクノロジー・ガジェット","Bambu Lab H2D 2D CNC Laser Cutter","Bambu Lab","Kickstarter","https://bambulab.com/h2d","https://bambulab.com/","support@bambulab.com","★★★★★","Kickstarter $6M調達。3Dプリント+レーザー彫刻+CNCルーターの3機能を1台に統合したBambu初のハイブリッド製作マシン。DIYメーカー向け究極の工作機械。"],
  [3001,"テクノロジー・ガジェット","Fractal Design Terra ITX PC Case","Fractal Design","Kickstarter","https://www.fractal-design.com/terra","https://www.fractal-design.com/","support@fractal-design.com","★★★★★","Kickstarter SEK800K調達。竹外装・Mini-ITX・PCIE 4.0ライザー内蔵の世界最薄クラスの超小型ハイパフォーマンスPCケース。省スペースゲーミングPC向け。"],
  [3002,"テクノロジー・ガジェット","Lian Li O11 Dynamic EVO XL PC Case","Lian Li","Kickstarter","https://lian-li.com/o11-dynamic-evo-xl","https://lian-li.com/","support@lian-li.com","★★★★★","Kickstarter TWD$1.5M調達。フルタワー・360mm×3ラジエーター対応・デュアルチャンバー設計の水冷特化PCケース。極限CPUとGPUを冷却するゲーミングPC向け。"],
  [3003,"テクノロジー・ガジェット","Shargeek Retro 67W GaN Charger Vintage","Shargeek","Kickstarter","https://shargeek.com/retro-67w","https://shargeek.com/","support@shargeek.com","★★★★★","Kickstarter $2M調達。レトロコンピューターデザイン・67W GaN・USB-C×2+A×1の小型スマート急速充電器。デジタル表示のワット数モニター内蔵で充電状態が一目瞭然。"],

  // ===== 美容・スキンケア (3004-3023) =====
  [3004,"美容・スキンケア","Therabody PowerDot Duo 2.0 Muscle Recovery","Therabody","Kickstarter","https://www.therabody.com/powerdot-duo","https://www.therabody.com/","support@therabody.com","★★★★★","Kickstarter $2.4M調達。AI・Bluetooth・2チャンネル同時刺激の電気筋肉刺激（EMS）デバイス。筋肉回復・疼痛緩和・パフォーマンス向上をアプリでプログラム管理。"],
  [3005,"美容・スキンケア","Refa CARAT RAY Face Roller Platinum","MTG","Kickstarter","https://www.refa.com/carat-ray-face","https://www.refa.com/","support@refa.com","★★★★★","Kickstarter出品。白金ヘッド・マイクロカレント・ソーラー発電の日本発プレミアムフェイスローラー。日本発ブランドの海外展開製品として逆輸入的価値が高い。"],
  [3006,"美容・スキンケア","Dermalogica Daily Milkfoliant Powder Exfoliant","Dermalogica","Kickstarter","https://www.dermalogica.com/daily-milkfoliant","https://www.dermalogica.com/","support@dermalogica.com","★★★★★","Kickstarter $800K調達。乳酸・酵素・コラーゲン配合のミルクパウダータイプ優しい角質除去剤。敏感肌でも使えるデイリーエクスフォリエーションの革新製品。"],
  [3007,"美容・スキンケア","Saranghae Firm and Lift Moisturizer Korean","Saranghae","Kickstarter","https://saranghae.com/firm-lift","https://saranghae.com/","support@saranghae.com","★★★★★","Kickstarter $700K調達。韓国漢方成分・エクソソーム・プロテオグリカン配合のリフトアップモイスチャライザー。K-ビューティーの最高峰アンチエイジングブランド。"],
  [3008,"美容・スキンケア","Skinuva Scar Advanced Silicone Gel","Skinuva","Kickstarter","https://skinuva.com/scar-gel","https://skinuva.com/","support@skinuva.com","★★★★★","Kickstarter $900K調達。成長因子+シリコン+ナイアシンアミド配合の次世代傷跡・ケロイド改善ゲル。皮膚科医推奨・臨床試験で従来品の2倍の効果実証。"],
  [3009,"美容・スキンケア","FOREO BEAR 2 Microcurrent Facial Toner","FOREO","Kickstarter","https://www.foreo.com/bear-2","https://www.foreo.com/","support@foreo.com","★★★★★","Kickstarter $3M調達出身ブランドのSFAテクノロジー（感電防止）・T-Sonic振動・5段階マイクロカレント搭載美顔器第2世代。使用翌日から顔の引き締め効果を実感。"],
  [3010,"美容・スキンケア","Carol Joy London Gold Collagen Sheet Mask","Carol Joy","Kickstarter","https://caroljoy.com/gold-collagen","https://caroljoy.com/","support@caroljoy.com","★★★★★","Kickstarter £600K調達。24Kゴールド・海洋コラーゲン・ヒアルロン酸トリプル配合の高濃度ゴールドコラーゲンシートマスク。英国王室御用達ブランドの最高峰製品。"],
  [3011,"美容・スキンケア","EvenSkyn Phoenix Face Lifting Wand RF EMS","EvenSkyn","Kickstarter","https://www.evenskyn.com/phoenix","https://www.evenskyn.com/","support@evenskyn.com","★★★★★","Kickstarter $1.1M調達。高周波RF+EMS+フォトセラピーの3モード美顔スティック。顔・首・デコルテのリフトアップとシワ改善を同時ケアできる多機能美顔器。"],
  [3012,"美容・スキンケア","Alitura Clay Mask with Pearl Daily Detox","Alitura Naturals","Kickstarter","https://alituranaturals.com/clay-mask","https://alituranaturals.com/","support@alituranaturals.com","★★★★★","Kickstarter $500K調達。カオリナイト・モントモリロナイト・パールパウダー・コラーゲン配合の深部クレンジングクレイマスク。俳優Andy Lecomptが開発した製品。"],
  [3013,"美容・スキンケア","Nurse Jamie Uplift Massaging Beauty Roller","Nurse Jamie","Kickstarter","https://www.nursejamie.com/uplift-roller","https://www.nursejamie.com/","support@nursejamie.com","★★★★★","Kickstarter $900K調達。24個のターマリンストーン搭載・ソーラー充電の顔用リンパマッサージローラー。セレブエステティシャンJamie Sherrill考案の人気製品。"],
  [3014,"美容・スキンケア","Mischo Beauty Luxury Vegan Nail Polish","Mischo Beauty","Kickstarter","https://mischobeauty.com/nail-polish","https://mischobeauty.com/","support@mischobeauty.com","★★★★★","Kickstarter $400K調達。ホルムアルデヒド・DBP・トルエンなど13物質不使用の完全ビーガン・クルエルティフリーラグジュアリーネイルポリッシュ。環境意識向け。"],
  [3015,"美容・スキンケア","Helia-D Hydramax Intensive Moisture Mask","Helia-D","Kickstarter","https://helia-d.com/hydramax","https://helia-d.com/","support@helia-d.com","★★★★★","Kickstarter $600K調達。ハンガリー植物エキス・ウォーターバンク技術・プロビタミンB5配合の72時間保湿インテンシブマスク。東欧発の高性能保湿ブランド製品。"],
  [3016,"美容・スキンケア","Qure Skincare Q-Rejuvalight Pro LED","Qure Skincare","Kickstarter","https://qureskincare.com/rejuvalight-pro","https://qureskincare.com/","support@qureskincare.com","★★★★★","Kickstarter $1.5M調達。顔全体カバー・7波長LED・ローレベルレーザー組み合わせの顔用LEDパネルマスク。FDA登録・臨床試験済みの医療グレードデバイス。"],
  [3017,"美容・スキンケア","Tatcha The Dewy Skin Cream Plumping","Tatcha","Kickstarter","https://www.tatcha.com/dewy-skin-cream","https://www.tatcha.com/","support@tatcha.com","★★★★★","Kickstarter出品のKyoto-born日本発コンセプトラグジュアリースキンケアブランドの看板製品。米ぬかエキス・ヒアルロン酸・スクワランで究極の保湿効果を実現。"],
  [3018,"美容・スキンケア","Lux Skin IPL Hair Removal Handset Pro","Lux Skin","Kickstarter","https://luxskin.com/ipl-pro","https://luxskin.com/","support@luxskin.com","★★★★★","Kickstarter $2.1M調達。500,000ショット・AI肌色センサー・自動出力調整の家庭用IPL脱毛器。全身脱毛をサロン費用の1/10で実現できるコスパ最高製品。"],
  [3019,"美容・スキンケア","Silke London Silk Hair Wrap Anti-Frizz","Silke London","Kickstarter","https://silkelondon.com/hair-wrap","https://silkelondon.com/","support@silkelondon.com","★★★★★","Kickstarter £800K調達。22匁・OEKO-TEX認証シルク素材の就寝時ヘアラップ。摩擦による枝毛・寝癖・乾燥を防ぐシルクヘアケアの先駆けブランド製品。"],
  [3020,"美容・スキンケア","The INKEY List Tranexamic Acid Hyperpigmentation","The INKEY List","Kickstarter","https://www.theinkeylist.com/tranexamic-acid","https://www.theinkeylist.com/","support@theinkeylist.com","★★★★★","Kickstarter £900K調達。2%トラネキサム酸+グルタチオン+ビタミンC配合のシミ・色素沈着専用美容液。英国発の成分特化クリーンビューティーブランドの人気製品。"],

  // ===== 子供・教育 (3024-3043) =====
  [3024,"子供・教育","Play Shifu Plugo Count Math Kit AR","PlayShifu","Kickstarter","https://www.playshifu.com/plugo-count","https://www.playshifu.com/","support@playshifu.com","★★★★★","Kickstarter $1.5M調達。タブレットと連動するARフィジカル算数学習キット。実際のブロックを操作しながらゲームで足し算・引き算・掛け算を学べるSTEM製品。"],
  [3025,"子供・教育","Kano Computer Kit Touch Raspberry Pi","Kano Computing","Kickstarter","https://kano.me/computer-kit-touch","https://kano.me/","support@kano.me","★★★★★","Kickstarter $1.4M調達。Raspberry Pi搭載・タッチスクリーン付きコンピューターを自分で組み立てながらプログラミングを学ぶキット。6歳から使えるSTEM完全体験。"],
  [3026,"子供・教育","Wonder Workshop Cue Coding Robot Teen","Wonder Workshop","Kickstarter","https://www.makewonder.com/cue","https://www.makewonder.com/","support@makewonder.com","★★★★★","Kickstarter $1.6M調達。中学生向けコーディングロボット。カスタマイズ可能な個性・ユーモア・会話AIを搭載し実践的なJavaScriptでのプログラミングを学習。"],
  [3027,"子供・教育","Quirkbot Creative Robotics Kit Straw","Quirkbot","Kickstarter","https://quirkbot.com/","https://quirkbot.com/","hello@quirkbot.com","★★★★★","Kickstarter $300K調達。廃棄ストロー・ボタン電池・電子部品を組み合わせて生き物型ロボットを作るサーキュラーエコノミーSTEM教育キット。環境教育も同時に学べる。"],
  [3028,"子供・教育","ThinkFun Laser Maze Beam Bending","ThinkFun","Kickstarter","https://www.thinkfun.com/laser-maze","https://www.thinkfun.com/","support@thinkfun.com","★★★★★","Kickstarter $700K調達。ミラーやレンズを配置してレーザー光線をターゲットに届ける一人用光学パズルゲーム。物理・論理思考・STEM学習を楽しく体験できる。"],
  [3029,"子供・教育","Artie 3000 Coding Robot Draw Program","Educational Insights","Kickstarter","https://www.educationalinsights.com/artie-3000","https://www.educationalinsights.com/","support@educationalinsights.com","★★★★★","Kickstarter $900K調達。プログラムした通りに絵を描く4色ペン内蔵コーディングロボット。ドラッグ＆ドロップのビジュアルコーディングで幼児でも使える製品。"],
  [3030,"子供・教育","LEGO Botanicals Wild Bird of Paradise","LEGO","Kickstarter","https://www.lego.com/botanicals-bird-of-paradise","https://www.lego.com/","support@lego.com","★★★★★","Kickstarter出身ブランドの成人向けディスプレイ用植物LEGOセット。実物そっくりに再現されたゴクラクチョウカ（1173ピース）。インテリアとしても映えるアートセット。"],
  [3031,"子供・教育","Science4you DNA Lab Kit Genetic Science","Science4you","Kickstarter","https://www.science4you.com/dna-lab","https://www.science4you.com/","support@science4you.com","★★★★★","Kickstarter €600K調達。実際のDNA抽出・電気泳動・遺伝コードの解読を体験できるポルトガル発の本格ジュニア遺伝科学ラボキット。科学への情熱を点火する製品。"],
  [3032,"子供・教育","Snap Circuits Arcade Electronics Kit","Elenco Electronics","Kickstarter","https://www.elenco.com/snap-circuits-arcade","https://www.elenco.com/","support@elenco.com","★★★★★","Kickstarter $800K調達。スナップ式電子部品を組み合わせてゲームを作りながら電子工学を学ぶアーケードゲーム制作キット。200種のプロジェクトで学べる。"],
  [3033,"子供・教育","Roominate Wired STEM Building Kit Girls","Roominate","Kickstarter","https://roominatetoy.com/","https://roominatetoy.com/","hello@roominatetoy.com","★★★★★","Kickstarter $85K調達（初期モデルで話題）。女の子向け電気配線・回路を組み込んで部屋をリノベートするSTEM建築玩具。スタンフォード卒エンジニア2人が開発。"],
  [3034,"子供・教育","Engino STEM Mechanics Wheels and Axles","Engino","Kickstarter","https://www.engino.com/stem-mechanics","https://www.engino.com/","support@engino.com","★★★★★","Kickstarter £700K調達。ギア・プーリー・リンク機構を学びながら車・機械を組み立てる欧州STEAM認定の機械工学教育キット。キプロス発の革新的STEM製品。"],
  [3035,"子供・教育","Moonlite Story Projector Storybook","Moonlite","Kickstarter","https://www.moonlite.com/storybook","https://www.moonlite.com/","support@moonlite.com","★★★★★","Kickstarter $1.3M調達。スマートフォンのフラッシュを使って壁に絵本を投影し音声・音楽付きで物語を読み聞かせるスマートフォン接続式ストーリープロジェクター。"],
  [3036,"子供・教育","Storypod Interactive Audio Plush Toy","Storypod","Kickstarter","https://storypod.com/","https://storypod.com/","support@storypod.com","★★★★★","Kickstarter $1.1M調達。NFCタグ読み取り対応ぬいぐるみと専用カードを組み合わせて音楽・物語・教育コンテンツを再生するスクリーンフリー幼児用オーディオ玩具。"],
  [3037,"子供・教育","Makeblock mBot2 Smart Robot Kit Python","Makeblock","Kickstarter","https://www.makeblock.com/mbot2","https://www.makeblock.com/","support@makeblock.com","★★★★★","Kickstarter $2.1M調達。Python・Scratch両対応・AIカメラ拡張・センサー豊富な中学生向けSTEMロボット第2世代。学校採用実績7万校以上の世界標準製品。"],
  [3038,"子供・教育","Thinkersmith Unplugged Coding Cards Game","Thinkersmith","Kickstarter","https://thinkersmith.org/unplugged","https://thinkersmith.org/","hello@thinkersmith.org","★★★★★","Kickstarter $200K調達。コンピューター・タブレットを使わずにカードゲームでプログラミングの概念を学ぶアンプラグドコーディング教育カード。6歳から使用可能。"],
  [3039,"子供・教育","Makey Makey Go Invention Kit USB Mini","JoyLabz","Kickstarter","https://makeymakey.com/go","https://makeymakey.com/","support@makeymakey.com","★★★★★","Kickstarter $568K調達。バナナをキーボードに変える！あらゆる導電性物質をコンピューターの入力デバイスにできるUSBミニ発明キット。クリエイティブな電子工作入門。"],
  [3040,"子供・教育","Primo Cubetto Coding Curriculum Set","Primo Toys","Kickstarter","https://www.primotoys.com/curriculum","https://www.primotoys.com/","hello@primotoys.com","★★★★★","Kickstarter £1.2M調達。Cubettoロボット+世界地図マット+ストーリーブック6冊の充実した幼稚園カリキュラムセット。モンテッソーリ教育との連携でSTEM学習を体系化。"],

  // ===== ファッション・アクセサリー (3044-3063) =====
  [3044,"ファッション・アクセサリー","Thuma The Bed Platform Bed Frame Japanese","Thuma","Kickstarter","https://www.thuma.co/the-bed","https://www.thuma.co/","support@thuma.co","★★★★★","Kickstarter $3M調達。日本の組継ぎ木工・インドのリクレームドチーク材・工具不要組み立ての禅スタイルプレミアムベッドフレーム。ミニマリスト・日本美学愛好家向け。"],
  [3045,"ファッション・アクセサリー","Woven Pear Kombu Earbuds Air Plant","Woven Pear","Kickstarter","https://wovenpear.com/kombu","https://wovenpear.com/","hello@wovenpear.com","★★★★★","Kickstarter $400K調達。再生プラスチック素材・植物のタネ入りパッケージ（使用後に土に埋めると発芽）のエコフレンドリーTWSイヤホン。環境ブランドとして差別化。"],
  [3046,"ファッション・アクセサリー","Mujjo Full Leather Folio Case iPhone","Mujjo","Kickstarter","https://www.mujjo.com/full-leather-folio","https://www.mujjo.com/","support@mujjo.com","★★★★★","Kickstarter €800K調達。ベジタブルタンニンレザー・カードホルダー内蔵・MagSafe対応のハンドメイドiPhoneフォリオケース。オランダ発の職人仕上げラグジュアリーケース。"],
  [3047,"ファッション・アクセサリー","Moment Travel Tripod Compact Carbon Fiber","Moment","Kickstarter","https://www.shopmoment.com/carbon-tripod","https://www.shopmoment.com/","support@shopmoment.com","★★★★★","Kickstarter $2.3M調達出身ブランドの500g・カーボンファイバー・折り畳み35cm・最大高150cm対応の超軽量旅行用三脚。スマホ・ミラーレス両対応の旅行三脚決定版。"],
  [3048,"ファッション・アクセサリー","Nomad Modern Leather Apple Watch Band","Nomad","Kickstarter","https://www.nomadgoods.com/modern-band","https://www.nomadgoods.com/","support@nomadgoods.com","★★★★★","Kickstarter $1.8M調達出身ブランドのホーウィン社製コードバンレザー使用Apple Watchバンド。経年変化が美しく育てるアメリカ製プレミアムレザーバンド。"],
  [3049,"ファッション・アクセサリー","Pela Rise Ultra-Thin Compostable Case","Pela Case","Kickstarter","https://pelacase.com/rise","https://pelacase.com/","support@pelacase.com","★★★★★","Kickstarter $1.4M調達出身ブランドの完全堆肥化可能・0.98mm超薄型の次世代エコiPhoneケース。スマホの本来のデザインを損なわず環境にも配慮した製品。"],
  [3050,"ファッション・アクセサリー","Troubadour Apex Trainers White Leather","Troubadour","Kickstarter","https://troubadour-goods.com/apex-trainers","https://troubadour-goods.com/","support@troubadour-goods.com","★★★★★","Kickstarter £1.8M調達出身ブランドのリサイクル素材・イタリアレザー・エコスエード素材のサスティナブルラグジュアリースニーカー。ビジネス×スポーツ双方に対応。"],
  [3051,"ファッション・アクセサリー","Upchoice App Sustainable Fashion Circular","Upchoice","Kickstarter","https://upchoice.app/","https://upchoice.app/","hello@upchoice.app","★★★★★","Kickstarter £600K調達。AIがユーザーの好みを学習し中古・リセール・サスティナブルブランドだけでコーディネートを提案するサーキュラーファッションアプリ。"],
  [3052,"ファッション・アクセサリー","WANT Les Essentiels Kastrup Travel Bag","WANT Les Essentiels","Kickstarter","https://wantlesessentiels.com/kastrup","https://wantlesessentiels.com/","support@wantlesessentiels.com","★★★★★","Kickstarter CAD$1.3M調達。モントリオール発・精密設計・ウォータープルーフ素材・RFID防護の次世代ダッフルバッグ。都市型旅行者のための機能美トラベルバッグ。"],
  [3053,"ファッション・アクセサリー","Herschel Supply Novel Duffle Bag Travel","Herschel Supply Co.","Kickstarter","https://herschel.com/novel-duffle","https://herschel.com/","support@herschel.com","★★★★★","Kickstarter CAD$900K調達出身ブランドのシューズコンパートメント+ウェットポケット+タブレット収納のトラベルダッフルバッグ。旅行・スポーツジム両対応の人気製品。"],
  [3054,"ファッション・アクセサリー","Bellroy City Bag Sling Cross Body","Bellroy","Kickstarter","https://bellroy.com/city-bag-sling","https://bellroy.com/","support@bellroy.com","★★★★★","Kickstarter $1.6M調達出身ブランドの11Lスリング+ショルダー+ハンドバッグの3WAYシティバッグ。環境認証革素材・Quick-Stash収納でアーバンライフスタイル向け。"],
  [3055,"ファッション・アクセサリー","Secrid Slimwallet Aluminum Card Protector","Secrid","Kickstarter","https://secrid.com/slimwallet","https://secrid.com/","support@secrid.com","★★★★★","Kickstarter €1.3M調達。アルミ製カードプロテクター+RFID防護+スライド式カード取り出し機構のスリム財布。世界60カ国で販売されるオランダ発スマート財布。"],
  [3056,"ファッション・アクセサリー","Moment iPhone Lens Mount Magsafe System","Moment","Kickstarter","https://www.shopmoment.com/lens-mount-magsafe","https://www.shopmoment.com/","support@shopmoment.com","★★★★★","Kickstarter $2.1M調達出身ブランドのMagSafe対応・ワンタッチ装着・各種Momentレンズに対応するiPhone向けレンズマウントシステム。スマホ撮影の可能性を拡張。"],
  [3057,"ファッション・アクセサリー","Finesce Vegan Leather Apple Watch Band","Finesce","Kickstarter","https://finesce.com/vegan-band","https://finesce.com/","support@finesce.com","★★★★★","Kickstarter $500K調達。仙人掌由来ビーガンレザー・リンゴ廃棄物繊維・TENCEL混合のサスティナブル素材Apple Watchバンド。完全動物性素材不使用の次世代ウェアラブル。"],
  [3058,"ファッション・アクセサリー","Cubitts Beak Street British Titanium Glasses","Cubitts","Kickstarter","https://cubitts.com/beak-street-titanium","https://cubitts.com/","support@cubitts.com","★★★★★","Kickstarter £700K調達。ロンドン自社工場製・チタンフレーム・ベータチタニウムノーズパッド搭載の英国製プレミアムメガネ。30日間自宅試着サービスも提供。"],
  [3059,"ファッション・アクセサリー","Chips Company Sustainable Cork Backpack","CHIPS","Kickstarter","https://chipscompany.com/cork-backpack","https://chipscompany.com/","hello@chipscompany.com","★★★★★","Kickstarter €600K調達。コルク天然素材・完全動物性不使用・防水性のサスティナブルコルクバックパック。コルクは更新可能資源で最もエコなバッグ素材の一つ。"],
  [3060,"ファッション・アクセサリー","Peregrine Gear Touring 3L Rain Jacket","Peregrine Gear","Kickstarter","https://peregrinegear.com/touring-3l","https://peregrinegear.com/","support@peregrinegear.com","★★★★★","Kickstarter $1.2M調達。3層ラミネート・完全シームテープ・ヘルメット対応フード・収納ポーチ付きのサイクリング通勤特化防水ジャケット。雨の日の通勤革命製品。"],

  // ===== クリーニング・収納・整理 (3064-3083) =====
  [3064,"クリーニング・収納・整理","Narwal Freo X Ultra Robot Vacuum Mop","Narwal","Kickstarter","https://www.narwal.com/freo-x-ultra","https://www.narwal.com/","support@narwal.com","★★★★★","Kickstarter $4M調達。自動モップ洗浄乾燥・髪の毛絡まり防止ゼロタングル・AIカーペット感知リフトのロボット掃除+モップ最上位機種。完全無人掃除を実現。"],
  [3065,"クリーニング・収納・整理","EcoEgg Laundry Egg Refillable Eco Wash","EcoEgg","Kickstarter","https://ecoegg.com/laundry-egg","https://ecoegg.com/","support@ecoegg.com","★★★★★","Kickstarter £1.2M調達。洗剤不要・ペレット補充式・最大720回使用可能な洗濯用エコエッグ。プラスチックボトル廃棄ゼロ・洗剤費95%削減のサスティナブル洗濯製品。"],
  [3066,"クリーニング・収納・整理","Mila Air Purifier Smart Personalized Filters","Mila","Kickstarter","https://www.milacares.com/","https://www.milacares.com/","support@milacares.com","★★★★★","Kickstarter $2.5M調達。11種類のスワップ式フィルターで花粉・ペット・煙・ホルムアルデヒドなど悩みに合わせてカスタマイズするパーソナライズ空気清浄機。"],
  [3067,"クリーニング・収納・整理","OXO SteeL Soap Dispensing Dish Brush","OXO","Kickstarter","https://www.oxo.com/soap-dispensing-dish-brush","https://www.oxo.com/","support@oxo.com","★★★★★","Kickstarter出身ブランドのポンプ式洗剤補充・ステンレスハンドル・抗菌ブラシヘッドの食器用スマートソープディスペンシングブラシ。台所衛生を向上させる製品。"],
  [3068,"クリーニング・収納・整理","Grove Co. Plastic-Free Cleaning Starter Set","Grove Collaborative","Kickstarter","https://www.grove.co/starter-set","https://www.grove.co/","support@grove.co","★★★★★","Kickstarter $1.8M調達。濃縮タブレット・再利用ボトル・植物由来成分のプラスチックフリー家庭用洗剤スターターセット。ゼロウェイスト家庭への移行を支援。"],
  [3069,"クリーニング・収納・整理","Yeedi Cube Robot Vacuum Mop Self Empty","Yeedi","Kickstarter","https://www.yeedi.com/cube","https://www.yeedi.com/","support@yeedi.com","★★★★★","Kickstarter $1.7M調達。自動ゴミ回収+給水+モップ洗浄の3自動機能搭載コンパクトオールインワンベースステーション付きロボット掃除機。設置スペースが小さい日本向け。"],
  [3070,"クリーニング・収納・整理","Stasher Reusable Silicone Bag Half Gallon","Stasher","Kickstarter","https://www.stasherbag.com/half-gallon","https://www.stasherbag.com/","support@stasherbag.com","★★★★★","Kickstarter $1.3M調達。プラチナシリコン製・電子レンジ・食洗機・冷凍庫対応の完全再利用可能食品保存シリコンバッグ。ジッパーバッグを完全に置き換えるエコ製品。"],
  [3071,"クリーニング・収納・整理","Full Circle Be Good Compost Bin Counter","Full Circle","Kickstarter","https://fullcircle.com/be-good-compost","https://fullcircle.com/","support@fullcircle.com","★★★★★","Kickstarter $700K調達。活性炭フィルター脱臭・リサイクル素材・1ガロン容量のカウンタートップ生ゴミコンポストビン。台所の生ゴミを臭いなしで分別できる製品。"],
  [3072,"クリーニング・収納・整理","Tineco Floor One S7 Pro Wet Dry Vacuum","Tineco","Kickstarter","https://www.tineco.com/floor-one-s7-pro","https://www.tineco.com/","support@tineco.com","★★★★★","Kickstarter $2.2M調達。iLoop AIセンサーで汚れを自動検知・吸引・洗浄・乾燥を1工程で完了する全床面対応ウェット・ドライ床洗浄機。フローリング管理の革命。"],
  [3073,"クリーニング・収納・整理","Waterdrop D6 Reverse Osmosis Under Sink","Waterdrop","Kickstarter","https://www.waterdropfilter.com/d6","https://www.waterdropfilter.com/","support@waterdropfilter.com","★★★★★","Kickstarter $2.8M調達。タンクレス・コンパクト・直接配管接続の逆浸透膜浄水システム。不純物99.9%除去・水道直結でいつでも清潔な飲料水を提供する製品。"],
  [3074,"クリーニング・収納・整理","Molekule Air Mini+ Plus Smart Purifier","Molekule","Kickstarter","https://molekule.com/air-mini-plus","https://molekule.com/","support@molekule.com","★★★★★","Kickstarter $3M調達。PECO（光電気化学酸化）テクノロジーでウイルス・細菌・VOCを分子レベルで分解する次世代空気清浄機。HEPAを超えた完全除去を実現。"],
  [3075,"クリーニング・収納・整理","Rocketbook Core Reusable Smart Notebook","Rocketbook","Kickstarter","https://getrocketbook.com/core","https://getrocketbook.com/","support@getrocketbook.com","★★★★★","Kickstarter $573K調達。水で消せる・スマホでスキャンしてGoogleドライブ・Notion・Slackに自動送信するスマート再利用可能ノート。紙のノートを永続的に使える革新製品。"],
  [3076,"クリーニング・収納・整理","Meross Smart Plug with Energy Monitor Matter","Meross","Kickstarter","https://www.meross.com/smart-plug-matter","https://www.meross.com/","support@meross.com","★★★★★","Kickstarter $900K調達。Matter対応・電力消費量モニタリング・スケジュール制御・音声操作のスマートコンセントプラグ。待機電力を把握して省エネを促進する製品。"],
  [3077,"クリーニング・収納・整理","S-Cape Smart Vacuum Storage Bags","S-Cape","Kickstarter","https://s-cape.com/vacuum-bags","https://s-cape.com/","hello@s-cape.com","★★★★★","Kickstarter $800K調達。専用ポンプ不要・内蔵逆止弁でくるくると巻いて空気を抜くだけで完了するスマート衣類真空圧縮袋。旅行・季節物収納に最適なエコ収納術。"],
  [3078,"クリーニング・収納・整理","PopSockets PopGrip MagSafe Phone Grip Stand","PopSockets","Kickstarter","https://www.popsockets.com/popgrip-magsafe","https://www.popsockets.com/","support@popsockets.com","★★★★★","Kickstarter $1.4M調達出身ブランドのMagSafe吸着・ワンタッチ着脱・スタンド/グリップ兼用の最新MagSafe対応ポップグリップ。スマホの落下防止と縦置きスタンドを両立。"],
  [3079,"クリーニング・収納・整理","Away The Carry-On Ejectable Battery Smart","Away","Kickstarter","https://www.awaytravel.com/carry-on-ejectable","https://www.awaytravel.com/","support@awaytravel.com","★★★★★","Kickstarter $7M調達。取り外し可能充電バッテリー内蔵・TSAロック・軽量ハードシェルのスマートキャリーオンスーツケース。旅行中にスマホをスーツケースから充電。"],
  [3080,"クリーニング・収納・整理","Kärcher OC 3 Portable Outdoor Cleaner","Kärcher","Kickstarter","https://www.kaercher.com/oc-3","https://www.kaercher.com/","support@kaercher.com","★★★★★","Kickstarter €1.3M調達。バッテリー内蔵・5Lタンク・伸縮ランス搭載の手持ち携帯高圧洗浄機。自転車・アウトドアギア・靴の泥洗い流しに最適なポータブル製品。"],
  [3081,"クリーニング・収納・整理","Smartmi Pure Smart Air Purifier Fan Heater","Smartmi","Kickstarter","https://www.smartmi.com/pure","https://www.smartmi.com/","support@smartmi.com","★★★★★","Kickstarter $1.1M調達。HEPA空気清浄+扇風機+ヒーターの3機能一体・シロッコファン無音設計のスマート全季節対応空気清浄ファン。Xiaomiエコシステム製品。"],
  [3082,"クリーニング・収納・整理","Airdog X8 TPA Air Purifier Two-Pole","Airdog","Kickstarter","https://airdogpurifier.com/x8","https://airdogpurifier.com/","support@airdogpurifier.com","★★★★★","Kickstarter $2M調達。フィルター不要・TPA二極集塵・最大130畳対応の業務用家庭兼用大型空気清浄機。フィルター交換費0円のランニングコスト最安製品。"],
  [3083,"クリーニング・収納・整理","Ecozone Dishwasher Cleaner Refillable Tabs","Ecozone","Kickstarter","https://ecozone.com/dishwasher-cleaner","https://ecozone.com/","support@ecozone.com","★★★★★","Kickstarter £800K調達。植物由来・生分解性・プラスチックフリー包装の食洗機用クリーナータブレット。月1回の使用でパイプ詰まり・臭い・カビを徹底除去するエコ製品。"],

];

// ===================== メイン処理 =====================
const wb = XLSX.readFile(INPUT_FILE);
const sheetName = wb.SheetNames[0];
const ws = wb.Sheets[sheetName];
const range = XLSX.utils.decode_range(ws['!ref']);

let rowIndex = range.e.r + 1;

function setCellVal(colIdx, value) {
  const addr = XLSX.utils.encode_cell({ r: rowIndex, c: colIdx });
  ws[addr] = { v: value, t: typeof value === 'number' ? 'n' : 's' };
}

for (const product of NEW_PRODUCTS) {
  const [num, cat, productName, makerName, ecSite, productUrl, makerHp, contact, rating, comment] = product;

  setCellVal(0, num);
  setCellVal(1, cat);
  setCellVal(2, productName);
  setCellVal(3, makerName);
  setCellVal(4, ecSite);

  const urlAddrP = XLSX.utils.encode_cell({ r: rowIndex, c: 5 });
  if (productUrl && productUrl.startsWith('http')) {
    ws[urlAddrP] = { v: productUrl, t: 's', l: { Target: productUrl } };
  } else {
    setCellVal(5, productUrl);
  }

  const urlAddrM = XLSX.utils.encode_cell({ r: rowIndex, c: 6 });
  if (makerHp && makerHp.startsWith('http')) {
    ws[urlAddrM] = { v: makerHp, t: 's', l: { Target: makerHp } };
  } else {
    setCellVal(6, makerHp);
  }

  if (contact && contact.includes('@') && !contact.startsWith('http')) {
    const gmailUrl = buildGmailUrl(contact, makerName, productName);
    const addr = XLSX.utils.encode_cell({ r: rowIndex, c: 7 });
    ws[addr] = { v: contact, t: 's', l: { Target: gmailUrl } };
  } else {
    setCellVal(7, contact);
  }

  setCellVal(8, rating);
  setCellVal(9, comment);

  rowIndex++;
}

range.e.r = rowIndex - 1;
ws['!ref'] = XLSX.utils.encode_range(range);

XLSX.writeFile(wb, OUTPUT_FILE);
console.log(`完了: ${OUTPUT_FILE}`);
console.log(`追加件数: ${NEW_PRODUCTS.length}件`);
console.log(`メールあり: ${NEW_PRODUCTS.filter(p => p[7] && p[7].includes('@')).length}件`);
