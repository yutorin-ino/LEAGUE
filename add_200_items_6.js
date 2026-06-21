/**
 * add_200_items_6.js — No.1684〜1883 (200件追加)
 * node add_200_items_6.js
 */

const XLSX = require('xlsx');
const path = require('path');

const INPUT_FILE  = path.join(__dirname, '海外便利グッズリスト_日本未上陸1705件_評価付.xlsx');
const OUTPUT_FILE = path.join(__dirname, '海外便利グッズリスト_日本未上陸1905件_評価付.xlsx');
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

  // ===== キッチン・調理器具 (1684-1703) =====
  [1684,"キッチン・調理器具","wan AIChef Ultra","WAN AICHEF","WAN AICHEF公式","https://www.global.wanaichef.com/products/wan-aichef-ultra","https://www.global.wanaichef.com/","liangwanping@wanaichef.com","★★★★★","CES 2026出展。AI「Zhurong Cooking Model」搭載のロボットシェフ。食材認識・調理・盛り付けまで完全自動化。料理自動化ニーズが高い日本市場に強く訴求。"],
  [1685,"キッチン・調理器具","LUMO AI Infrared Indoor Grill","Cozytime","Cozytime公式","https://cozytime-tech.com/pages/lumo","https://cozytime-tech.com/","https://cozytime-tech.com/pages/contact-us","★★★★★","CES 2026出展。赤外線光でBBQ風味を室内で再現。AIが食材タイプ・厚み・重量を自動検知して火加減調整。煙が出ないため日本の住宅事情に最適。"],
  [1686,"キッチン・調理器具","SmartVoice Air Fryer","Emerson","Emerson公式","https://www.emerson.com/","https://www.emerson.com/","info@emerson.com","★★★★☆","CES 2026出展。1,000以上の音声コマンド対応の音声操作エアフライヤー。ハンズフリー調理を実現。高齢者・忙しい共働き世帯向け。"],
  [1687,"キッチン・調理器具","C-200 Ultrasonic Knife","Seattle Ultrasonics","Seattle Ultrasonics公式","https://seattleultrasonics.com/","https://seattleultrasonics.com/","info@seattleultrasonics.com","★★★★★","CES 2026出展。毎秒30,000回以上振動する超音波ナイフ。通常の包丁より50%少ない力で切断可能。プロシェフ・調理器具愛好家向け。"],
  [1688,"キッチン・調理器具","Bosch Cook AI Cooking Guidance","Bosch Home Appliances","Bosch公式","https://www.bosch-home.com/","https://www.bosch-home.com/","https://www.bosch-home.com/en/contact","★★★★☆","CES 2026出展。生成AI・センサー・カメラ統合で動的な個人向け料理ガイダンスを提供。既存のBosch調理器具と連携。"],
  [1689,"キッチン・調理器具","Fellow Aiden Precision Coffee Maker","Fellow","Fellow公式","https://fellowproducts.com/products/aiden-pour-over-coffee-maker","https://fellowproducts.com/","support@fellowproducts.com","★★★★★","クラフトコーヒーをボタン一つで再現するプレシジョンコーヒーメーカー。ドリップ温度・流量を精密制御。コーヒー文化が根付く日本市場に最適。"],
  [1690,"キッチン・調理器具","CHEF iQ Smart Thermometer Pro","CHEF iQ","Amazon US","https://chefiq.com/products/smart-thermometer","https://chefiq.com/","support@chefiq.com","★★★★★","AIでリアルタイム調理ガイドを提供するワイヤレス料理用温度計。スマホ連動で揚げ物・ステーキを失敗なく調理。"],
  [1691,"キッチン・調理器具","Anova Precision Cooker Pro","Anova Culinary","Amazon US","https://anovaculinary.com/products/anova-precision-cooker-pro","https://anovaculinary.com/","support@anovaculinary.com","★★★★★","Wi-Fi接続・1,200W出力のプレミアム低温調理器。プロ品質の真空低温調理を家庭で実現。食への関心が高い日本市場に響く。"],
  [1692,"キッチン・調理器具","Bartesian Premium Cocktail Machine","Bartesian","Amazon US","https://bartesian.com/","https://bartesian.com/","hello@bartesian.com","★★★★☆","カプセルを入れるだけでバーテンダー品質のカクテルを自動作製するマシン。在宅飲み需要が高まる日本でヒットの可能性大。"],
  [1693,"キッチン・調理器具","Suvie Kitchen Robot","Suvie","Suvie公式","https://www.suvie.com/","https://www.suvie.com/","hello@suvie.com","★★★★★","冷蔵保存から自動調理まで一台で完結するキッチンロボット。帰宅に合わせて出来立て料理が完成。共働き世帯の多い日本に最適。"],
  [1694,"キッチン・調理器具","Lissome R1 Countertop Dishwasher","Lissome","Lissome公式","https://lissome.co/","https://lissome.co/","info@lissome.co","★★★★★","AIスマートウォッシュ技術搭載の置き型食洗機。15分で洗浄・乾燥完了。設置工事不要で日本の賃貸住宅に最適。水道接続不要モデル。"],
  [1695,"キッチン・調理器具","Hisense FreshVault Refrigerator Drawer","Hisense","Hisense公式","https://www.hisense.com/","https://www.hisense.com/","info@hisense.com","★★★★☆","真空密封ドロワーで食材を最大5日間新鮮保存するフレンチドアタイプ冷蔵庫の革新的機能。食品ロス削減意識の高い日本市場向け。"],
  [1696,"キッチン・調理器具","Ooni Volt 12 Electric Pizza Oven","Ooni","Ooni公式","https://ooni.com/products/ooni-volt","https://ooni.com/","support@ooni.com","★★★★★","450℃まで達する家庭用電気ピザオーブン。屋内外両対応でナポリタンピザを2分で焼き上げ。ピザ人気の高い日本で高需要。"],
  [1697,"キッチン・調理器具","Brava Home Light Oven","Brava Home","Brava公式","https://www.brava.com/","https://www.brava.com/","support@brava.com","★★★★★","純光（ピュアライト）で焼く革新的オーブン。食材をカメラで認識し自動で最適調理。レシピアプリ連動で失敗知らずの調理を実現。"],
  [1698,"キッチン・調理器具","Thermomix TM7 Smart Kitchen Robot","Vorwerk","Thermomix公式","https://www.thermomix.com/","https://www.thermomix.com/","info@thermomix.com","★★★★★","世界80か国以上で販売されるオールインワンスマートキッチンロボット。切る・混ぜる・煮る・蒸すを一台で完結。料理の手間を大幅削減。"],
  [1699,"キッチン・調理器具","Hestan Cue Smart Cooking System","Hestan Smart Cooking","Hestan Cue公式","https://www.hestancue.com/","https://www.hestancue.com/","hello@hestancue.com","★★★★☆","スマートフライパン+誘導クッキングシステムで料理プロセスをアプリがリアルタイムガイド。プロ品質の料理を家庭で再現。"],
  [1700,"キッチン・調理器具","Breville Joule Turbo Sous Vide","Breville","Amazon US","https://www.breville.com/us/en/products/joule/bsv600.html","https://www.breville.com/","support@breville.com","★★★★★","ターボ加熱機能搭載の高性能低温調理器。通常より最大40%速く調理完了。ステーキ・チキンをプロ品質に仕上げる。"],
  [1701,"キッチン・調理器具","Nosh AI Robo-Chef","Nosh","Kickstarter","https://hellonosh.com/","https://hellonosh.com/","hello@hellonosh.com","★★★★★","CES 2026出展。500以上のレシピから選ぶだけで完全自律調理するAIロボットシェフ。食材をセットするだけで調理完了。"],
  [1702,"キッチン・調理器具","Drinkmate OmniFizz Sparkling Water Maker","Drinkmate","Amazon US","https://drinkmate.us/","https://drinkmate.us/","info@drinkmate.us","★★★★☆","炭酸水・ジュース・ワインなどあらゆる飲料を炭酸化できる多用途炭酸メーカー。日本のコンビニ炭酸水文化との親和性が高い。"],
  [1703,"キッチン・調理器具","Zwilling Enfinigy Electric Kettle Pro","Zwilling","Amazon US","https://www.zwilling.com/us/zwilling-enfinigy-electric-kettle-pro/","https://www.zwilling.com/","support@zwilling.com","★★★★☆","ドイツZwilling製の精密温度制御電気ケトル。緑茶・紅茶・コーヒー向けに1度単位で温度調整。お茶文化の日本市場に訴求。"],

  // ===== スマートホーム・インテリア・照明 (1704-1723) =====
  [1704,"スマートホーム・インテリア・照明","Govee Sky Ceiling Light","Govee","Amazon US","https://us.govee.com/products/govee-sky-ceiling-light","https://us.govee.com/","marketing@govee.com","★★★★★","CES 2026で9メディア賞受賞。高密度ピクセルと空模様ディフューザーで自然光を室内に再現するスマート天井照明。テレワーク環境の改善にも最適。"],
  [1705,"スマートホーム・インテリア・照明","Govee Floor Lamp 3","Govee","Amazon US","https://us.govee.com/","https://us.govee.com/","distribution@govee.com","★★★★★","CES 2026で9賞受賞。1000K〜10000Kの業界最広色温度範囲とAI適応照明システム搭載フロアランプ。気分・時間帯に合わせて自動調光。"],
  [1706,"スマートホーム・インテリア・照明","Lockin V7 Max Smart Lock","Lockin","Lockin公式","https://lockly.com/products/lockin-v7-max","https://lockly.com/","info@lockly.com","★★★★★","CES 2026 Innovation Award受賞。3D顔認証・掌紋静脈・指紋静脈の3重生体認証で0.17秒解錠。光充電システム搭載で電池切れ不要。"],
  [1707,"スマートホーム・インテリア・照明","Ultraloq Bolt Sense Smart Lock","Xthings","Amazon US","https://ultraloq.com/products/ultraloq-bolt-sense","https://ultraloq.com/","support@ultraloq.com","★★★★★","CES 2026出展。0.5秒で掌紋静脈認証が完了するスマートロック。非接触・高精度で衛生面でも優れる。"],
  [1708,"スマートホーム・インテリア・照明","Ulticam HaLow 1.5-Mile Security Camera","Xthings","Amazon US","https://ultraloq.com/pages/ulticam-halow","https://ultraloq.com/","support@ultraloq.com","★★★★☆","CES 2026出展。Wi-Fi HaLow技術で1.5マイル（約2.4km）届く超長距離セキュリティカメラ。農村・広い敷地でも安定監視。"],
  [1709,"スマートホーム・インテリア・照明","Aqara Thermostat Hub W200","Aqara","Amazon US","https://www.aqara.com/eu/product/thermostat-hub-w200","https://www.aqara.com/","support@aqara.com","★★★★★","CES 2026出展。Apple「適応温度」「クリーンエネルギーガイダンス」対応の初の機器。4インチカラーで空調・セキュリティを一元管理。"],
  [1710,"スマートホーム・インテリア・照明","OVAL Smart Home Hub","OVAL","OVAL公式","https://ovalsmarthome.com/","https://ovalsmarthome.com/","hello@ovalsmarthome.com","★★★★★","CES 2026出展。物の場所を把握・文脈を理解・人物を識別できるAIスマートホームハブ。深夜の不審者と子どもの遊びを区別して通知。"],
  [1711,"スマートホーム・インテリア・照明","MAMMOTION LUBA 3 AWD Robot Mower","MAMMOTION","Amazon US","https://us.mammotion.com/products/luba-3","https://us.mammotion.com/","distribution@mammotion.com","★★★★★","CES 2026出展。360°LiDAR+デュアル1080p AI Vision+NetRTKで±1cm精度の自律芝刈りロボット。ワイヤー設置不要。"],
  [1712,"スマートホーム・インテリア・照明","Roborock RockMow X1 LiDAR Robot Mower","Roborock","Roborock公式","https://www.roborock.com/products/rockmow-x1","https://www.roborock.com/","support@roborock.com","★★★★★","CES 2026出展。Sentisphere LiDARで庭の3Dマップを生成しリアルタイム自律走行する次世代ロボット芝刈り機。"],
  [1713,"スマートホーム・インテリア・照明","NexLawn NAVIA 6000 AWD Robot Mower","NexLawn","NexLawn公式","https://nexlawn.com/","https://nexlawn.com/","info@nexlawn.com","★★★★☆","CES 2026出展。悪路・傾斜地での自律動作に特化した高性能4WDロボット芝刈り機。法面・不整地対応モデル。"],
  [1714,"スマートホーム・インテリア・照明","Dreo 765S TurboCool Misting Fan","Dreo","Amazon US","https://www.dreosmart.com/products/dreo-765s-tower-fan","https://www.dreosmart.com/","support@dreosmart.com","★★★★☆","CES 2026出展。超微細ミストと扇風機を組み合わせた冷却タワーファン。ミストが蒸発時に熱を吸収し効率的に冷却。節電志向の日本に最適。"],
  [1715,"スマートホーム・インテリア・照明","touchHL AI House Platform","touchHL","touchHL公式","https://touchhl.com/","https://touchhl.com/","info@touchhl.com","★★★★★","CES 2026 Innovation Award受賞。ユーザー行動を学習しながら温度・照明・空気質を自動最適化するAI統合スマートホームプラットフォーム。"],
  [1716,"スマートホーム・インテリア・照明","Nanoleaf Lines Squared","Nanoleaf","Amazon US","https://nanoleaf.me/en-US/products/nanoleaf-lines/","https://nanoleaf.me/","support@nanoleaf.me","★★★★★","壁に直接貼るLEDラインパネル。16Mカラー・Matter対応で音楽連動も可能。ゲーミング・インテリア好きの日本の若年層に高需要。"],
  [1717,"スマートホーム・インテリア・照明","Govee Neon Rope Light 2","Govee","Amazon US","https://us.govee.com/products/govee-neon-rope-light-2","https://us.govee.com/","marketing@govee.com","★★★★☆","自由に形が作れるネオン風LEDロープライト。Matter対応・AIシーン自動生成機能付き。インテリア照明として日本でも高い需要。"],
  [1718,"スマートホーム・インテリア・照明","Eve Energy Smart Plug with Energy Monitor","Eve Systems","Amazon US","https://www.evehome.com/en/eve-energy","https://www.evehome.com/","support@evehome.com","★★★★☆","Matter対応・電力消費量モニタリング機能付きスマートプラグ。Thread搭載で超低遅延動作。節電意識の高い日本市場に最適。"],
  [1719,"スマートホーム・インテリア・照明","WiZ Smart Color Bulb A21","WiZ Connected","Amazon US","https://www.wizconnected.com/","https://www.wizconnected.com/","info@wizconnected.com","★★★★☆","設定不要・Wi-Fi直接接続のスマートLED電球。1,600万色・音楽連動・スケジュール設定対応。スマート照明入門として最適。"],
  [1720,"スマートホーム・インテリア・照明","LIFX Candle Color Smart Bulb","LIFX","Amazon US","https://www.lifx.com/products/lifx-candle-color","https://www.lifx.com/","support@lifx.com","★★★★☆","ハブ不要でWi-Fi直接接続の高輝度カラースマートLED電球。HomeKit・Alexa・Google対応。高演色性でインテリア照明に最適。"],
  [1721,"スマートホーム・インテリア・照明","Sylvox Deck Pro Outdoor TV","Sylvox","Sylvox公式","https://www.sylvox.com/","https://www.sylvox.com/","support@sylvox.com","★★★★☆","CES 2026出展。屋外用スーパーナローベゼル4K TV。明るさ1,000nit以上で日中屋外でも鮮明表示。テラス・バルコニー設置向け。"],
  [1722,"スマートホーム・インテリア・照明","IKEA Kajplats Smart LED Bulb","IKEA","IKEA公式","https://www.ikea.com/","https://www.ikea.com/","https://www.ikea.com/us/en/customer-service/contact-us/","★★★★☆","CES 2026出展。Matter over Thread標準対応の次世代IKEAスマート電球。Apple Home・Google Home・Amazonと直接連携可能。"],
  [1723,"スマートホーム・インテリア・照明","Arlo Ultra 2 XL Spotlight Camera","Arlo","Amazon US","https://www.arlo.com/en-us/cameras/ultra/VMC5282-200NAS.html","https://www.arlo.com/","support@arlo.com","★★★★★","4K HDR・180度視野・カラーナイトビジョン・スポットライト一体型ワイヤレスセキュリティカメラ。防犯意識が高い日本市場に強く訴求。"],

  // ===== ウェアラブル・ヘルス・フィットネス (1724-1743) =====
  [1724,"ウェアラブル・ヘルス・フィットネス","Garmin Venu 4 Smartwatch","Garmin","Garmin公式","https://www.garmin.com/en-US/p/1043921","https://www.garmin.com/","support@garmin.com","★★★★★","CES 2026 Innovation Award受賞。HRV・呼吸・睡眠を追跡しベースラインから外れた時に通知する「健康ステータス」機能搭載スマートウォッチ。"],
  [1725,"ウェアラブル・ヘルス・フィットネス","Withings BodyScan 2 Smart Scale","Withings","Withings公式","https://www.withings.com/us/en/bodyscan","https://www.withings.com/","support@withings.com","★★★★★","CES 2026 Innovation Award受賞。神経血管スコア・体重・体組成・心電図・脈波伝播を一度に測定する次世代スマート体重計。"],
  [1726,"ウェアラブル・ヘルス・フィットネス","Speediance Strap Fitness Tracker","Speediance","Speediance公式","https://www.speediance.com/products/strap","https://www.speediance.com/","after-sales@speediance.com","★★★★★","CES 2026出展。画面なし・サブスクなしのフィットネストラッカー。心拍・睡眠・体温・HRVを追跡しレディネススコアを算出。"],
  [1727,"ウェアラブル・ヘルス・フィットネス","Bond Ring Smart Ring","WilderTech","Bond Ring公式","https://www.bondring.com/","https://www.bondring.com/","shop@bondring.com","★★★★★","CES 2026出展。体熱で自己充電する世界初の充電不要スマートリング。血圧・血糖傾向・HRV・睡眠・生理周期など14種以上のバイオマーカーを追跡。"],
  [1728,"ウェアラブル・ヘルス・フィットネス","Circular Ring 2 Smart Ring","Circular","Kickstarter","https://www.circular.xyz/","https://www.circular.xyz/","contact@circular.xyz","★★★★★","CES 2025受賞・Kickstarter4分完売。精密ECG・AI健康インサイト・13種以上の健康機能を搭載したスマートリング。"],
  [1729,"ウェアラブル・ヘルス・フィットネス","Ultrahuman Ring Pro","Ultrahuman","Ultrahuman公式","https://www.ultrahuman.com/ring-pro","https://www.ultrahuman.com/","support@ultrahuman.com","★★★★★","チタン製・4カラー展開のサブスクなしスマートリング。デュアルコアプロセッサ搭載で睡眠追跡精度が大幅向上。サブスク不要で経済的。"],
  [1730,"ウェアラブル・ヘルス・フィットネス","airing Smart Ring for Hormones","airing smart","Kickstarter","https://www.kickstarter.com/projects/2016344821/a-smart-ring-that-understands-hormones-emotions-and-health","https://airingring.com/","hello@airingring.com","★★★★★","世界初のホルモン・感情・健康対応スマートリング。女性の生理周期・排卵・感情変化を追跡。女性ヘルスケア分野で高需要。"],
  [1731,"ウェアラブル・ヘルス・フィットネス","NuraLogix Longevity Mirror","NuraLogix","NuraLogix公式","https://www.nuralogix.ai/","https://www.nuralogix.ai/","info@nuralogix.ai","★★★★★","CES 2026出展。30秒の「ビデオ自撮り」で20年先の健康リスクを予測するウェルネスデバイス。ヘルスコンシャスな日本市場に革命的。"],
  [1732,"ウェアラブル・ヘルス・フィットネス","Cionic Neural Sleeve 2","Cionic","Cionic公式","https://www.cionic.com/","https://www.cionic.com/","info@cionic.com","★★★★★","TIME 2025年ベスト発明品。筋肉に電気刺激を与え歩行を補助するスマートリハビリスリーブ。高齢化が進む日本の介護・リハビリ分野に革命的。"],
  [1733,"ウェアラブル・ヘルス・フィットネス","Garmin Forerunner 970 Running Watch","Garmin","Garmin公式","https://www.garmin.com/en-US/p/1045127","https://www.garmin.com/","support@garmin.com","★★★★★","CES 2026 Innovation Award受賞。AIランニングコーチ・MicroLEDディスプレイ・進化したランニングダイナミクス搭載の最上位ランニングウォッチ。"],
  [1734,"ウェアラブル・ヘルス・フィットネス","Garmin Fenix 8 Pro MicroLED","Garmin","Garmin公式","https://www.garmin.com/en-US/p/1043899","https://www.garmin.com/","support@garmin.com","★★★★★","CES 2026 Innovation Award受賞。MicroLEDディスプレイ採用の最高峰マルチスポーツGPSスマートウォッチ。登山・ダイビング・ランニング対応。"],
  [1735,"ウェアラブル・ヘルス・フィットネス","Movano Evie Ring Smart Ring","Movano Health","Movano公式","https://www.movanohealth.com/evie-ring","https://www.movanohealth.com/","info@movanohealth.com","★★★★★","女性向けに設計されたFDA認定スマートリング。生理周期・排卵・心拍・睡眠・酸素濃度を追跡。女性健康意識の高い日本市場向け。"],
  [1736,"ウェアラブル・ヘルス・フィットネス","Withings ScanWatch Nova","Withings","Withings公式","https://www.withings.com/us/en/scanwatch-nova","https://www.withings.com/","support@withings.com","★★★★★","心電図・血中酸素・皮膚温度センサー搭載の医療グレードハイブリッドスマートウォッチ。クラシックな外観で30日間バッテリー持続。"],
  [1737,"ウェアラブル・ヘルス・フィットネス","BioIntelliSense BioButton Patch","BioIntelliSense","BioIntelliSense公式","https://www.biointellisense.com/","https://www.biointellisense.com/","info@biointellisense.com","★★★★★","貼るだけで24時間連続バイタル監視できる医療グレードのバイオセンサーパッチ。在宅医療・遠隔患者モニタリングに革命的。高齢化日本に最適。"],
  [1738,"ウェアラブル・ヘルス・フィットネス","InteraXon Muse S Headband Gen 2","InteraXon","Amazon US","https://choosemuse.com/muse-s/","https://choosemuse.com/","hello@choosemuse.com","★★★★★","脳波（EEG）・心拍・呼吸・体動を測定するスマートメディテーションヘッドバンド。睡眠・ストレス管理ニーズが高い日本市場に最適。"],
  [1739,"ウェアラブル・ヘルス・フィットネス","Axil XCOR SE Smart Ear Protection","Axil","Amazon US","https://www.axilgear.com/","https://www.axilgear.com/","support@axilgear.com","★★★★☆","TIME 2025年ベスト発明品。有害な大音量は遮断しつつ会話は増幅するスマートサウンドプロテクション。工事・農業・音楽業界向け。"],
  [1740,"ウェアラブル・ヘルス・フィットネス","Cosmo Robotics Bambini Kids Exoskeleton","Cosmo Robotics","Cosmo Robotics公式","https://cosmorobotics.com/","https://cosmorobotics.com/","info@cosmorobotics.com","★★★★★","CES 2026 Innovation Award受賞。足首動力機能付き世界初の地上歩行小児用外骨格。脳性まひ等の小児リハビリに革命的。"],
  [1741,"ウェアラブル・ヘルス・フィットネス","Huawei Watch GT 6 Pro","Huawei","Huawei公式","https://consumer.huawei.com/en/watches/watch-gt6-pro/","https://consumer.huawei.com/","https://consumer.huawei.com/en/support/contact/","★★★★☆","TIME 2025年ベスト発明品。スポーツ特化型マルチスポーツスマートウォッチ。TruSeen 6.0+ PPGセンサーと超高精度GPS搭載。"],
  [1742,"ウェアラブル・ヘルス・フィットネス","Oura Ring 4","Oura","Oura公式","https://ouraring.com/product/rings","https://ouraring.com/","support@ouraring.com","★★★★★","最新第4世代スマートリング。新世代センサー・18項目の健康指標追跡・6日間バッテリー。サブスクモデルだが健康データの質が最高水準。"],
  [1743,"ウェアラブル・ヘルス・フィットネス","Whoop 5.0 Performance Band","Whoop","Whoop公式","https://www.whoop.com/en-us/thelocker/whoop-5-0/","https://www.whoop.com/","support@whoop.com","★★★★☆","画面なし・AI健康コーチング特化型スポーツバンド。回復スコア・睡眠コーチで運動パフォーマンスを最大化。アスリート向け。"],

  // ===== アウトドア・スポーツ・旅行 (1744-1763) =====
  [1744,"アウトドア・スポーツ・旅行","Jackery Solar Mars Bot","Jackery","Jackery公式","https://www.jackery.com/pages/solar-mars-bot","https://www.jackery.com/","media@jackery.com","★★★★★","CES 2026出展。AI搭載コンピュータービジョンで自律走行・ユーザー追跡・日照追跡する自律ソーラー発電ロボット。キャンプ・アウトドアに革命的。"],
  [1745,"アウトドア・スポーツ・旅行","Timeli Personal Safety Device","Timeli","Timeli公式","https://timeli.com/","https://timeli.com/","info@timeli.com","★★★★★","CES 2026 Innovation Award受賞。懐中電灯・HDビデオ・大音量アラーム・GPSトラッキング・緊急通報を一台に統合した個人安全デバイス。防災意識が高い日本向け。"],
  [1746,"アウトドア・スポーツ・旅行","VIATRIX Powered Exoskeleton","VIATRIX","VIATRIX公式","https://www.viatrix.com/","https://www.viatrix.com/","info@viatrix.com","★★★★★","CES 2026出展。日常歩行・スポーツ・アウトドア探索向けの電動外骨格。歩行補助・体力増強で登山・長距離ウォーキングを楽に。高齢者向けにも最適。"],
  [1747,"アウトドア・スポーツ・旅行","AC Future AI-THt Transforming Trailer","AC Future","AC Future公式","https://www.acfuture.com/","https://www.acfuture.com/","info@acfuture.com","★★★★☆","CES 2026 Innovation Award受賞。完全展開時35平方フィート超の居住空間になる変形式AIトレーラー。Pininfaranaデザイン×AIスマート機能。"],
  [1748,"アウトドア・スポーツ・旅行","Garmin Descent S1 Diver Communication Buoy","Garmin","Garmin公式","https://www.garmin.com/en-US/p/1015817","https://www.garmin.com/","support@garmin.com","★★★★★","CES 2026 Innovation Award受賞。ダイバー間の水中コミュニケーションを可能にするブイ。ダイビング人口が多い日本の離島・沖縄エリアで需要大。"],
  [1749,"アウトドア・スポーツ・旅行","CamperKit 11-in-1 Modular Outdoor Tool","Aecooly","Kickstarter","https://www.kickstarter.com/projects/aecooly/camperkit-11-in-1-outdoor-solution-for-every-adventure","https://aecooly.com/","info@aecooly.com","★★★★★","Kickstarter 2025年人気プロジェクト。ファン・ライト・アークライター・ウォーターポンプ等11モジュールを磁気コネクタで交換できるモジュラーキャンプツール。"],
  [1750,"アウトドア・スポーツ・旅行","Powerleaf+ Solar Charger","Flextech Solar","Kickstarter","https://www.kickstarter.com/projects/flexsolar/powerleaf-a-solar-charger-perfect-for-outdoor-enth","https://flexsolar.com/","info@flexsolar.com","★★★★☆","Kickstarter 2025年出品。調整可能スタンド付きの軽量折りたたみ式ソーラーチャージャー。バックパッキング・登山に最適。USB+DC出力対応。"],
  [1751,"アウトドア・スポーツ・旅行","O-Boy Satellite Emergency Smartwatch","O-Boy","O-Boy公式","https://o-boy.com/","https://o-boy.com/","info@o-boy.com","★★★★★","圏外でも衛星通信で双方向メッセージ・SOSを送受信できる緊急対応スマートウォッチ。登山・離島・僻地アウトドアで命綱になる製品。"],
  [1752,"アウトドア・スポーツ・旅行","Garmin inReach Mini 2 Satellite Communicator","Garmin","Amazon US","https://www.garmin.com/en-US/p/765374","https://www.garmin.com/","support@garmin.com","★★★★★","世界中どこでも衛星経由で双方向通信・SOSが送れる超小型衛星通信機。登山・海外旅行中の安全対策として日本でも高い需要。"],
  [1753,"アウトドア・スポーツ・旅行","BioLite CampStove 2+","BioLite","Amazon US","https://www.bioliteenergy.com/products/campstove-2-plus","https://www.bioliteenergy.com/","support@bioliteenergy.com","★★★★★","薪燃焼の熱を電力に変換しUSB充電ができる革新的キャンプ用ストーブ。燃料不要・防災用途にも優れる。環境意識高い日本市場向け。"],
  [1754,"アウトドア・スポーツ・旅行","Goal Zero Yeti 3000X Portable Power Station","Goal Zero","Amazon US","https://www.goalzero.com/products/yeti-3000x-portable-power-station","https://www.goalzero.com/","support@goalzero.com","★★★★★","3,032Wh大容量・AC/USB/DC全出力対応のポータブル電源。大型キャンプ・防災備蓄・車中泊に最適。アウトドアブームの日本で高需要。"],
  [1755,"アウトドア・スポーツ・旅行","EcoFlow DELTA Pro 3 Power Station","EcoFlow","Amazon US","https://www.ecoflow.com/us/delta-pro-3","https://www.ecoflow.com/","support@ecoflow.com","★★★★★","4,096Wh・4,000W出力・LFPバッテリー搭載の最高性能ポータブル電源。停電時も家電フル稼働可能。日本の防災市場で最有力製品。"],
  [1756,"アウトドア・スポーツ・旅行","Insta360 X5 Action Camera","Insta360","Amazon US","https://www.insta360.com/product/insta360-x5","https://www.insta360.com/","support@insta360.com","★★★★★","8K 360度撮影対応アクションカメラ。手ブレ補正・防水・AI編集機能搭載。旅行・スポーツ動画をプロ品質で収録。SNS発信好きの日本若年層向け。"],
  [1757,"アウトドア・スポーツ・旅行","DJI Osmo Action 5 Pro","DJI","Amazon US","https://www.dji.com/osmo-action-5-pro","https://www.dji.com/","info@dji.com","★★★★★","1/1.3インチ大センサー・10m防水・4K60fps対応アクションカメラ。氷点下対応で日本のスキー・雪山登山でも活躍。"],
  [1758,"アウトドア・スポーツ・旅行","Hydrow Wave Rowing Machine","Hydrow","Hydrow公式","https://hydrow.com/pages/hydrow-wave","https://hydrow.com/","support@hydrow.com","★★★★☆","プロインストラクターのライブ・オンデマンドローイングでクラス体験ができるコネクテッドローイングマシン。在宅フィットネス需要の高い日本向け。"],
  [1759,"アウトドア・スポーツ・旅行","Aviron Tough Series Rowing Machine","Aviron","Aviron公式","https://www.aviron.com/","https://www.aviron.com/","support@aviron.com","★★★★☆","ゲーム感覚で楽しめるインタラクティブローイングマシン。ゲーミングモード搭載で運動が苦手な人でも継続しやすい。フィットネス市場向け。"],
  [1760,"アウトドア・スポーツ・旅行","Therabody RecoveryAir JetBoots","Therabody","Amazon US","https://www.therabody.com/products/recoveryair-jetboots","https://www.therabody.com/","support@therabody.com","★★★★★","空気圧マッサージで下半身の疲労回復を促進するスマートリカバリーブーツ。アスリート・スポーツ愛好者の疲労回復に革命的。"],
  [1761,"アウトドア・スポーツ・旅行","Hypervolt Go 2 Massage Gun","Hyperice","Amazon US","https://hyperice.com/products/hypervolt-go-2","https://hyperice.com/","support@hyperice.com","★★★★★","NBA・NFL公式採用の軽量コンパクトマッサージガン。静音設計・3段階強度調整で旅行携帯にも最適。スポーツ需要の高い日本向け。"],
  [1762,"アウトドア・スポーツ・旅行","Matador Beast28 2.0 Technical Backpack","Matador","Amazon US","https://matadorup.com/products/beast28-technical-backpack","https://matadorup.com/","support@matadorup.com","★★★★☆","1,000Dコーデュラ素材・完全防水・超軽量のテクニカルデイパック。タウンユース〜トレッキングまで対応する機能美型バッグ。"],
  [1763,"アウトドア・スポーツ・旅行","NEMO Tensor Insulated Sleeping Pad","NEMO Equipment","Amazon US","https://www.nemoequipment.com/products/tensor-insulated-sleeping-pad","https://www.nemoequipment.com/","support@nemoequipment.com","★★★★★","R値3.5の断熱性能と超軽量を両立した最高峰インフレータブルスリーピングパッド。三シーズン〜冬季登山まで対応。"],

  // ===== ペット用品 (1764-1783) =====
  [1764,"ペット用品","SATELLAI Collar Go GPS Dog Collar","SATELLAI","SATELLAI公式","https://satellai.com/products/collar-go","https://satellai.com/","business@satellai.com","★★★★★","CES 2026 Best of CES受賞。Petsense AIで行動の微細な変化から健康問題を早期発見する4G GPS犬用スマートカラー。関節痛・病気の早期発見に革命的。"],
  [1765,"ペット用品","Invoxia Biotracker 2026 Dog Collar","Invoxia","Invoxia公式","https://www.invoxia.com/petcare/minitailz-dog-tracker","https://www.invoxia.com/","https://www.invoxia.com/en-US/contact","★★★★★","CES 2026出展。毎秒更新のLTE-M実時間位置追跡+心肺健康モニタリング+AIを統合した世界唯一の犬用GPSカラー。"],
  [1766,"ペット用品","Netvue Birdfy Bath Pro AI Bird Bath","Netvue Technologies","Netvue公式","https://www.birdfy.com/products/birdfy-bath","https://www.birdfy.com/","info@netvue.com","★★★★★","CES 2026 Best of Innovation受賞。AI搭載スマートバードバス。6,000種以上の野鳥を自動識別・写真撮影・SNS共有機能付き。バードウォッチング愛好家向け。"],
  [1767,"ペット用品","PetPogo PetPhone Pet Wearable","uCloudlink","PetPogo公式","https://petpogo.com/","https://petpogo.com/","info@petpogo.com","★★★★★","CES 2026出展。業界初・ペット用ウェアラブル「スマートフォン」。飼い主とペットがどこからでもリアルタイム双方向通話可能。分離不安対策に革命的。"],
  [1768,"ペット用品","Tuanty Pet Fresh Food Maker","Tuanty","Tuanty公式","https://www.tuanty.com/","https://www.tuanty.com/","info@tuanty.com","★★★★★","CES 2026出展。25分以内に手作りドッグフードが完成する全自動カウンタートップ家電。蒸す・混ぜるが食洗機対応ボウル一つで完結。ペット食品安全意識が高い日本向け。"],
  [1769,"ペット用品","Sphinx Smart Wet Cat Feeder","Sphinx","Kickstarter","https://www.kickstarter.com/projects/sphinxcatfeeder/sphinx-the-worlds-smartest-cat-feeder","https://sphinxcatfeeder.com/","info@sphinxcatfeeder.com","★★★★★","Kickstarter出品。ウェットフードを自動供給・鮮度管理・リアルタイムオーディオビデオ付きスマート猫用自動給餌器。多頭飼い・留守番対策に最適。"],
  [1770,"ペット用品","ePawDen Microchip Pet Feeder","ePawDen","Kickstarter","https://www.kickstarter.com/projects/epawden/epawden-microchip-pet-feeder-solve-multi-pet-feeding-issue","https://epawden.com/","info@epawden.com","★★★★★","Kickstarter出品。マイクロチップ認証で各ペットが正しい食事のみ食べられるRFID対応自動給餌器。多頭飼い家庭の食事管理問題を解決。"],
  [1771,"ペット用品","Netvue Birdfy Feeder Vista AI Camera","Netvue Technologies","Amazon US","https://www.birdfy.com/products/birdfy-feeder-vista","https://www.birdfy.com/","info@netvue.com","★★★★☆","CES 2026 Innovation Award Honoree。野鳥自動識別AI搭載バードフィーダーカメラ。6,000種以上の鳥を識別・自動撮影しSNS投稿可能。"],
  [1772,"ペット用品","Fi Series 3 GPS Dog Collar","Fi","Fi公式","https://tryfi.com/products/series3","https://tryfi.com/","support@tryfi.com","★★★★★","LTE+Wi-Fi+GPSのトリプル接続で世界最高レベルのリアルタイム追跡精度を持つスマート犬用カラー。逃走検知・活動追跡・睡眠モニタリング付き。"],
  [1773,"ペット用品","Tractive GPS 5 Dog Tracker","Tractive","Amazon US","https://tractive.com/en/products/gps-for-dogs","https://tractive.com/","support@tractive.com","★★★★★","ヨーロッパで最も人気のGPS犬用追跡デバイス。175カ国対応・リアルタイム追跡・活動モニタリング・スマホアプリ連動。"],
  [1774,"ペット用品","Furbo 5 Dog Camera with Treat Toss","Furbo","Amazon US","https://furbo.com/products/furbo-5-dog-camera","https://furbo.com/","support@furbo.com","★★★★★","AI犬認識・360度パン機能・おやつ投げ機能付きペットカメラ。吠え検知・分離不安通知機能搭載。在宅勤務増加で需要拡大の日本市場向け。"],
  [1775,"ペット用品","Halo Collar 3 GPS + Training Collar","Halo Collar","Halo Collar公式","https://www.halocollar.com/","https://www.halocollar.com/","support@halocollar.com","★★★★☆","GPS仮想フェンス＋トレーニング機能統合スマートカラー。ドッグトレーナーCesar Millan監修。フェンスなしで犬を安全に管理。"],
  [1776,"ペット用品","Whistle Switch GPS Pet Tracker","Whistle","Amazon US","https://www.whistle.com/products/switch","https://www.whistle.com/","support@whistle.com","★★★★☆","GPS追跡+ヘルスモニタリング統合型スマートペットトラッカー。4G LTE・Wi-Fi対応でリアルタイム位置情報と健康データを一元管理。"],
  [1777,"ペット用品","Cheerble Wicked Ball EP Pet Toy","Cheerble","Amazon US","https://cheerble.com/products/wicked-ball-ep","https://cheerble.com/","support@cheerble.com","★★★★★","AI搭載の自動回避・自走・休憩モード付き猫犬兼用インタラクティブボール。留守番中のペットの運動不足解消。一人暮らしペット飼育者に最適。"],
  [1778,"ペット用品","Petlibro Granary Wi-Fi Pet Feeder","Petlibro","Amazon US","https://petlibro.com/products/granary-wifi-pet-feeder","https://petlibro.com/","support@petlibro.com","★★★★★","Wi-Fi接続・スマートフォンアプリ遠隔操作の自動給餌器。食事動画録画・詰まり検知・フレッシュロック機能付き。外出が多い日本の働き世代向け。"],
  [1779,"ペット用品","PetSafe ScoopFree Smart Self-Cleaning Litter Box","PetSafe","Amazon US","https://store.petsafe.net/scoopfree-smart","https://www.petsafe.net/","support@petsafe.net","★★★★★","健康データ追跡機能付きの自動セルフクリーニング猫トイレ。排泄回数・体重・使用時間を記録し健康異常を早期発見。猫多頭飼い家庭向け。"],
  [1780,"ペット用品","Litter-Robot 4 Self-Cleaning Cat Litter Box","Whisker","Amazon US","https://www.litter-robot.com/litter-robot-4.html","https://www.litter-robot.com/","support@litter-robot.com","★★★★★","世界最高評価の全自動猫トイレ。Wi-Fi接続・スマホアプリ健康モニタリング・消臭システム搭載。猫飼育人口が多い日本で高需要。"],
  [1781,"ペット用品","Inubox Automatic Dog Toilet","Inubox","Inubox公式","https://www.inubox.com/","https://www.inubox.com/","info@inubox.com","★★★★★","犬用の自動セルフクリーニングトイレ。センサーで排泄を検知し自動で洗浄・消臭。マンション暮らしが多い日本の犬飼育者に革命的。"],
  [1782,"ペット用品","Petcube Bites 2 Lite Pet Camera","Petcube","Amazon US","https://petcube.com/petcube-bites-2-lite/","https://petcube.com/","support@petcube.com","★★★★☆","おやつ投げ機能付き・1080p HDペットカメラ。Alexa内蔵・スマートアラート・2方向音声通話対応。ペットを見守りながらおやつを与えられる。"],
  [1783,"ペット用品","EZwhelp Self-Warming Pet Mat","EZwhelp","Amazon US","https://ezwhelp.com/","https://ezwhelp.com/","info@ezwhelp.com","★★★★☆","電力不要・体温反射で温まる高機能ペットマット。防水・洗濯機対応・滑り止め付き。老犬・術後ペット・子猫の保温に最適。"],

  // ===== テクノロジー・ガジェット (1784-1803) =====
  [1784,"テクノロジー・ガジェット","Mojie World's Lightest Full-Color AR Glasses","Mojie","CES 2026 / Mojie公式","https://www.mojie.com/","https://www.mojie.com/","info@mojie.com","★★★★★","CES 2026 Innovation Award受賞。重さわずか38gの世界最軽量フルカラーARグラス。樹脂回折導波路採用で高透過・高解像度・プライバシー保護。"],
  [1785,"テクノロジー・ガジェット","Even Realities G2 Display Smart Glasses","Even Realities","Even Realities公式","https://www.evenrealities.com/products/g2-a","https://www.evenrealities.com/","support@evenrealities.com","★★★★★","CES 2026 Innovation Award受賞。普通のメガネと見分けがつかない外観のディスプレイ搭載スマートグラス。処方レンズ対応・終日装着可能。"],
  [1786,"テクノロジー・ガジェット","Vuzix Ultralite Pro Enterprise AR Glasses","Vuzix","Vuzix公式","https://www.vuzix.com/products/vuzix-ultralite-pro","https://www.vuzix.com/","info@vuzix.com","★★★★☆","CES 2026 Innovation Award受賞。処方レンズ対応の企業向け軽量ARグラス。製造・物流・医療現場でのハンズフリー作業支援に最適。"],
  [1787,"テクノロジー・ガジェット","Povec C1 Electrochromic Sunglasses","Povec Optics","Povec公式","https://www.povecoptics.com/","https://www.povecoptics.com/","info@povecoptics.com","★★★★★","CES 2026出展。世界初の手動レンズ濃度調整型エレクトロクロミックサングラス。1秒以内にCAT1〜3の濃度を切替。バイクライダー・アウトドア向け。"],
  [1788,"テクノロジー・ガジェット","Asus ROG Xreal R1 Gaming Glasses","Asus ROG x XREAL","XREAL公式","https://www.xreal.com/rog-glasses/","https://www.xreal.com/","support@xreal.com","★★★★★","CES 2026出展。世界最高240Hzリフレッシュレート搭載のゲーミングスマートグラス。eスポーツ愛好家・ゲーマーに革命的なゲーム体験を提供。"],
  [1789,"テクノロジー・ガジェット","XREAL Air 2 Ultra AR Glasses","XREAL","Amazon US","https://www.xreal.com/air2ultra/","https://www.xreal.com/","support@xreal.com","★★★★★","電磁6DoFトラッキング搭載でAR空間に仮想画面を固定表示できるARグラス。開発者・MR体験向けに最高レベルのAR性能。"],
  [1790,"テクノロジー・ガジェット","Brilliant Labs Frame AI Smart Glasses","Brilliant Labs","Brilliant Labs公式","https://brilliant.xyz/products/frame","https://brilliant.xyz/","hello@brilliant.xyz","★★★★★","視野内にリアルタイムAI情報（翻訳・情報検索・ナビ）をオーバーレイ表示するオープンソースAIスマートグラス。旅行・学習・ビジネス向け。"],
  [1791,"テクノロジー・ガジェット","Ray-Ban Meta Smart Glasses (2025 edition)","Meta","Amazon US","https://www.meta.com/smart-glasses/","https://www.meta.com/","https://www.meta.com/help/","★★★★★","AI搭載・カメラ・スピーカー一体型スマートグラス。「What am I looking at?」と聞けばAIが即回答。日本未発売の最新モデルで高需要。"],
  [1792,"テクノロジー・ガジェット","Solos AirGo Vision AI Smart Glasses","Solos Technology","Solos公式","https://www.solosglasses.com/","https://www.solosglasses.com/","support@solosglasses.com","★★★★☆","ChatGPT統合AIアシスタント搭載スマートグラス。開放型スピーカーで耳をふさがずに音楽・通話・AIアシストを利用可能。"],
  [1793,"テクノロジー・ガジェット","Rabbit R2 AI Companion Device","Rabbit","Rabbit公式","https://www.rabbit.tech/","https://www.rabbit.tech/","support@rabbit.tech","★★★★☆","LAM（大型行動モデル）搭載のポケットサイズAIデバイス。スマホアプリを代わりに操作・情報取得・タスク実行をシームレスに実行。"],
  [1794,"テクノロジー・ガジェット","Boox Palma 2 Android E-ink Phone","BOOX","Amazon US","https://shop.boox.com/products/palma2","https://shop.boox.com/","support@boox.com","★★★★★","スマートフォンサイズのAndroid搭載E-inkデバイス。目に優しい電子ペーパー画面・読書特化設計・超長時間バッテリー。読書文化の日本市場に最適。"],
  [1795,"テクノロジー・ガジェット","Anker SOLIX F3800 Power Station","Anker","Amazon US","https://www.anker.com/products/anker-solix-f3800","https://www.anker.com/","support@anker.com","★★★★★","3,840Wh・6,000W出力・拡張最大26.9kWhのホーム対応ポータブル電源。太陽光発電と組み合わせてオフグリッド生活も可能。防災需要の日本向け。"],
  [1796,"テクノロジー・ガジェット","Bambu Lab A1 Combo 3D Printer","Bambu Lab","Bambu Lab公式","https://bambulab.com/en/a1","https://bambulab.com/","support@bambulab.com","★★★★★","自動キャリブレーション・マルチカラー印刷対応の高速家庭用3Dプリンター。複雑なパーツを高品質・高速で造形。ものづくり好きな日本市場向け。"],
  [1797,"テクノロジー・ガジェット","DJI Neo 2 Lightweight Drone","DJI","Amazon US","https://www.dji.com/neo-2","https://www.dji.com/","info@dji.com","★★★★★","135g超軽量・掌離陸・AIアクティブトラック搭載の初心者〜中級者向けドローン。規制緩和が進む日本でも航空法適合で飛行可能。"],
  [1798,"テクノロジー・ガジェット","Starlink Mini Satellite Internet Kit","SpaceX Starlink","Starlink公式","https://www.starlink.com/mini","https://www.starlink.com/","https://www.starlink.com/support","★★★★★","ノートPC並みの超コンパクトサイズの携帯型衛星インターネットキット。離島・山岳地・農村での高速ネット接続に革命的。日本の地方通信インフラ不足を補う。"],
  [1799,"テクノロジー・ガジェット","Anker SOLIX Solarbank 2 Pro E1600","Anker","Amazon US","https://www.anker.com/products/anker-solix-solarbank-2-pro","https://www.anker.com/","support@anker.com","★★★★★","ソーラーパネルと組み合わせてベランダ・屋外で使えるスマートバッテリー蓄電システム。電気代節約・防災備蓄に最適。"],
  [1800,"テクノロジー・ガジェット","Humane AI Pin Wearable AI Device","Humane","Humane公式","https://hu.ma.ne/aipin","https://hu.ma.ne/","support@hu.ma.ne","★★★★☆","スマートフォン不要のウェアラブルAIデバイス。衣服にクリップして装着しAIアシスタント・翻訳・撮影・通話が可能。新世代ガジェットとして注目。"],
  [1801,"テクノロジー・ガジェット","Logitech MX Ink Stylus for VR","Logitech","Amazon US","https://www.logitech.com/en-us/products/vr/mx-ink.html","https://www.logitech.com/","support@logitech.com","★★★★☆","Meta Quest対応のVR空間でリアルな書き心地を再現するスタイラスペン。3D設計・アート・教育向けVRクリエイティブ作業を革新。"],
  [1802,"テクノロジー・ガジェット","AYANEO Pocket S Gaming Handheld","AYANEO","AYANEO公式","https://www.ayaneo.com/product/AYANEO-POCKET-S","https://www.ayaneo.com/","info@ayaneo.com","★★★★★","Snapdragon G3x Gen 2搭載の最高性能Android携帯ゲーム機。Steam Deckより薄く軽量で持ち運び優秀。携帯ゲーム文化の日本で高需要。"],
  [1803,"テクノロジー・ガジェット","Anker Prime Power Bank 27650mAh","Anker","Amazon US","https://www.anker.com/products/anker-prime-power-bank","https://www.anker.com/","support@anker.com","★★★★★","双方向250W出力・27,650mAh大容量のフラッグシップモバイルバッテリー。スマホ・タブレット・ラップトップを同時急速充電可能。"],

  // ===== 美容・スキンケア (1804-1823) =====
  [1804,"美容・スキンケア","L'Oréal LED Face Mask","L'Oréal","L'Oréal公式","https://www.loreal.com/en/press-release/group/loreal-advances-leadership-in-beauty-tech-at-ces-2026/","https://www.loreal.com/","https://www.loreal.com/en/contact-us/","★★★★★","CES 2026 Innovation Award受賞。超薄型シリコン製フレキシブルLEDマスク。赤色光+近赤外線光で細胞活性化・抗老化効果。2027年発売予定。美容大国日本で最高需要。"],
  [1805,"美容・スキンケア","Kolmar Korea Scar Beauty Device","Kolmar Korea","Kolmar公式","https://www.kolmarkorea.com/","https://www.kolmarkorea.com/","info@kolmarkorea.com","★★★★★","CES 2026 Best of Innovation受賞（Beauty Tech＆Digital Health両部門）。AIで傷跡を12種類に分類し最適化治療液を直接投与するデバイス。"],
  [1806,"美容・スキンケア","Appotronics Laser Hair Regrowth Cap","Appotronics","Appotronics公式","https://www.appotronics.com/en/","https://www.appotronics.com/","info@appotronics.com","★★★★★","CES 2026出展。650nm低出力レーザー120本搭載・165g超軽量の発毛促進キャップ。1日20分装着で発毛効果を実現。薄毛に悩む日本人男性向けに高需要。"],
  [1807,"美容・スキンケア","Appotronics Laser Hair Growth Band","Appotronics","Appotronics公式","https://www.appotronics.com/en/","https://www.appotronics.com/","info@appotronics.com","★★★★★","CES 2026出展。650nmレーザー搭載のヘアバンド型発毛デバイス。外出中にも帽子の下で装着可能。通勤中でも発毛ケアが継続できる革新設計。"],
  [1808,"美容・スキンケア","Appotronics Multi-Function Laser Beauty Device","Appotronics","Appotronics公式","https://www.appotronics.com/en/","https://www.appotronics.com/","info@appotronics.com","★★★★★","CES 2026出展。650nmレーザーで肌深層にアプローチするマルチファンクションレーザー美容デバイス。シミ・しわ・毛穴ケアを一台で実現。"],
  [1809,"美容・スキンケア","Amorepacific Skinsight Electronic Skin Platform","Amorepacific","Amorepacific公式","https://www.apgroup.com/","https://www.apgroup.com/","https://www.apgroup.com/int/en/contact.html","★★★★★","CES 2026 Innovation Award受賞（Beauty Tech）。MIT共同開発の電子肌プラットフォーム。皮膚状態をリアルタイム分析しパーソナライズドスキンケアを実現。"],
  [1810,"美容・スキンケア","FOREO LUNA 4 Plus Smart Facial Cleansing","FOREO","Amazon US","https://www.foreo.com/luna-4-plus","https://www.foreo.com/","hello@foreo.com","★★★★★","AIパーソナライズド機能搭載の第4世代スマートフェイシャルクレンジングデバイス。16種類のT-Sonic振動モードを肌タイプ別に自動調整。"],
  [1811,"美容・スキンケア","CurrentBody Skin LED Light Therapy Mask","CurrentBody","Amazon US","https://www.currentbody.com/products/currentbody-skin-led-light-therapy-mask","https://www.currentbody.com/","support@currentbody.com","★★★★★","クリニカルグレードの赤色光+近赤外線LEDマスク。しわ・ニキビ・肌質改善に効果的。アンバサダーがVictoria Beckhamでブランド力高い。"],
  [1812,"美容・スキンケア","Therabody TheraFace Mask","Therabody","Amazon US","https://www.therabody.com/products/theraface-mask","https://www.therabody.com/","support@therabody.com","★★★★★","赤色光+近赤外線光+温熱+振動を統合したオールインワンフェイスケアマスク。NBA公式スポーツリカバリーブランドによる本格美顔マスク。"],
  [1813,"美容・スキンケア","NuFACE Trinity+ Facial Toning Device","NuFACE","Amazon US","https://www.mynuface.com/products/trinity-plus-facial-toning-device","https://www.mynuface.com/","support@mynuface.com","★★★★★","FDA認定の微弱電流（マイクロカレント）フェイシャルトーニングデバイス。たるみ・リフトアップ・輪郭引き締めに効果的。美顔意識の高い日本女性向け。"],
  [1814,"美容・スキンケア","Omnilux Contour Face LED Mask","Omnilux","Amazon US","https://www.omniluxled.com/products/omnilux-contour-face","https://www.omniluxled.com/","support@omniluxled.com","★★★★★","医療グレードのフレキシブルLEDフェイスマスク。赤色光+近赤外線光でシワ・ニキビ・肌再生を促進。皮膚科医推奨製品。美容意識の高い日本女性向け。"],
  [1815,"美容・スキンケア","HigherDOSE Infrared Sauna Blanket V3","HigherDOSE","Amazon US","https://higherdose.com/products/infrared-sauna-blanket","https://higherdose.com/","support@higherdose.com","★★★★★","遠赤外線サウナ効果を自宅寝室で体験できるスマートブランケット。デトックス・代謝促進・ストレス解消。サウナブームの日本市場に最適。"],
  [1816,"美容・スキンケア","Solawave Radiant Renewal Skincare Wand","Solawave","Amazon US","https://solawave.co/products/radiant-renewal-skincare-wand","https://solawave.co/","support@solawave.co","★★★★★","赤色光+微弱電流+温熱+マッサージ振動の4機能統合スキンケアワンド。5分間のケアでクリニック級の肌改善効果。忙しい日本女性向け。"],
  [1817,"美容・スキンケア","Dr. Dennis Gross DRx SpectraLite FaceWare Pro","Dr. Dennis Gross Skincare","Amazon US","https://www.drdennisgross.com/products/drx-spectralite-faceware-pro","https://www.drdennisgross.com/","support@drdennisgross.com","★★★★★","FDA認定162灯LED搭載フルフェイスマスク。赤色光（しわ改善）+青色光（ニキビ撃退）の組み合わせ。米国皮膚科医開発の臨床グレード製品。"],
  [1818,"美容・スキンケア","Nira Laser Skincare Anti-Aging Device","Nira Skincare Laser","Amazon US","https://niraskincare.com/products/nira-precision-laser","https://niraskincare.com/","support@niraskincare.com","★★★★★","FDA認定の家庭用非剥脱性ダイオードレーザー。コラーゲン産生を促進し深いシワを改善。クリニック施術の代替として美容意識の高い層に高需要。"],
  [1819,"美容・スキンケア","Silk'n Titan Anti-Aging Skin Tightening","Silk'n","Amazon US","https://www.silkn.com/products/titan","https://www.silkn.com/","support@silkn.com","★★★★★","光エネルギー+双極RF+接触冷却の3技術統合で肌を引き締めるアンチエイジングデバイス。クリニック品質の肌引き締めを自宅で実現。"],
  [1820,"美容・スキンケア","Nood The Flasher 2.0 IPL Hair Removal","Nood","Amazon US","https://www.nood.com/products/the-flasher-2","https://www.nood.com/","support@nood.com","★★★★★","サロン品質の永久脱毛効果を自宅で実現するIPL光脱毛デバイス。全身対応・60万フラッシュ搭載・肌色センサー安全設計。脱毛ニーズが高い日本女性向け。"],
  [1821,"美容・スキンケア","Ulike Air 10 IPL Hair Removal Handset","Ulike","Amazon US","https://www.ulike.com/products/air10","https://www.ulike.com/","support@ulike.com","★★★★★","サファイア冷却技術搭載・第10世代IPL光脱毛器。痛みゼロで全身脱毛可能。メンズ・レディース対応。医療機関品質を家庭で実現。"],
  [1822,"美容・スキンケア","Lyma Laser Skincare System","Lyma","Lyma公式","https://lyma.life/products/lyma-laser","https://lyma.life/","support@lyma.life","★★★★★","世界最高出力の家庭用半導体レーザー搭載スキンケアシステム。ヒアルロン酸を肌深層に直接浸透させコラーゲン産生を促進。英国製・医師推奨製品。"],
  [1823,"美容・スキンケア","Medicube Age-R AMP Shot Booster","Medicube","Amazon US","https://medicube.us/products/age-r-amp-shot","https://medicube.us/","support@medicube.us","★★★★★","韓国発・電気穿孔技術（エレクトロポレーション）でスキンケア成分を肌深層に浸透させるホームケアデバイス。K-ビューティー製品として日本でも高い注目度。"],

  // ===== 子供・教育 (1824-1843) =====
  [1824,"子供・教育","Seedpace AI Story Robot","Seedpace","Seedpace公式","https://us.seedpace.com/","https://us.seedpace.com/","https://us.seedpace.com/pages/contact","★★★★★","CES 2026出展。子どもがストーリーに参加・選択肢選び・会話ができるAIインタラクティブストーリーロボット。3〜8歳対象・想像力・創造性を育む。"],
  [1825,"子供・教育","Toniebox 2 Audio Learning Player","Tonies","Toniebox公式","https://us.tonies.com/products/toniebox","https://us.tonies.com/","support@tonies.com","★★★★★","CES 2026出展。1〜9歳対象のスクリーンフリー音声プレーヤー。トニーキャラクターをのせるだけで物語・音楽・ゲームが起動。子どもの自主性と集中力を育む。"],
  [1826,"子供・教育","Baracoda Kolibree Gamified Toothbrush","Baracoda","Baracoda公式","https://www.baracoda.com/kolibree","https://www.baracoda.com/","info@baracoda.com","★★★★★","CES 2026出展。ゲーミフィケーションで子どもが自発的に歯磨きを楽しめるスマート歯ブラシ。仕上げ磨き不要の習慣づけを実現。"],
  [1827,"子供・教育","Hex Bots Wall Crawler STEM Robot","Hexnub","Amazon US","https://hexnub.com/","https://hexnub.com/","info@hexnub.com","★★★★★","CES 2026出展。AIチャレンジと実組み立てを組み合わせたSTEM特化型壁面走行ロボット。プログラミング入門から発展学習まで対応。"],
  [1828,"子供・教育","Hasbro Furby DJ AI Interactive Toy","Hasbro","Amazon US","https://shop.hasbro.com/en-us/category/furby","https://www.hasbro.com/","info@hasbro.com","★★★★☆","CES 2026出展。触れる・音・感情に反応するAI搭載インタラクティブFurby DJ。子どもの感情表現・コミュニケーション能力を育む。"],
  [1829,"子供・教育","Namibox AI Smart Learning Glasses","Namibox","Namibox公式","https://www.namibox.com/","https://www.namibox.com/","info@namibox.com","★★★★★","CES 2026出展。子ども向けAI学習スマートグラス。見たものをリアルタイムでAIが説明・翻訳・解説。好奇心旺盛な子どもの探究心を刺激。"],
  [1830,"子供・教育","Wonder Workshop Dash Robot","Wonder Workshop","Amazon US","https://www.makewonder.com/robots/dash/","https://www.makewonder.com/","support@makewonder.com","★★★★★","6歳以上向けのビジュアルプログラミング学習ロボット。Blockly・JavaScript・Python対応で段階的にプログラミングを学習。学校教育でも採用実績多数。"],
  [1831,"子供・教育","Sphero BOLT+ Programmable Robot","Sphero","Amazon US","https://sphero.com/products/sphero-bolt-plus","https://sphero.com/","support@sphero.com","★★★★★","透明シェル・LED光通信・AGLセンサー搭載のプログラミング学習ロボットボール。STEM教育ツールとして世界中の学校で採用実績。"],
  [1832,"子供・教育","Matatalab Pro Set Coding Robot","Matatalab","Amazon US","https://www.matatalab.com/products/matatalab-pro-set","https://www.matatalab.com/","support@matatalab.com","★★★★★","コードカードを並べるだけでプログラミングが学べる5〜9歳向けスクリーンフリー学習ロボット。世界70か国以上・1,000万人以上の子どもが使用。"],
  [1833,"子供・教育","LittleBits STEAM+ Inventor Kit","littleBits","Amazon US","https://littlebits.com/products/steam-plus-kit","https://littlebits.com/","support@littlebits.com","★★★★★","電子回路を磁石でつなぐだけで発明体験ができるSTEAM教育キット。アプリ連動・ステップバイステップガイドで子どもが自由に発明。"],
  [1834,"子供・教育","Osmo Genius Starter Kit for iPad","Osmo","Amazon US","https://www.playosmo.com/en/genius-starter-kit/","https://www.playosmo.com/","support@playosmo.com","★★★★★","iPadカメラを使い現実世界の操作をゲームに反映するAR教育ゲームキット。算数・英語・プログラミングを遊びながら学習。タブレット普及の日本向け。"],
  [1835,"子供・教育","Kano Computer Kit Touch","Kano","Amazon US","https://kano.me/store/products/kano-computer-kit-touch","https://kano.me/","support@kano.me","★★★★★","タッチスクリーン付きラズベリーパイベースのDIYコンピューターキット。組み立て・プログラミング・アート制作を一体で学べる革新的教育製品。"],
  [1836,"子供・教育","LeapFrog Magic Adventures Microscope","LeapFrog","Amazon US","https://www.leapfrog.com/en-us/products/magic-adventures-microscope","https://www.leapfrog.com/","support@leapfrog.com","★★★★☆","350倍の実体顕微鏡+200以上の科学事実を学べるARアプリを統合した子ども向けスマート顕微鏡。自然・科学への好奇心を育む。"],
  [1837,"子供・教育","Vtech Kidizoom Smartwatch DX3","VTech","Amazon US","https://www.vtech.com/en/kidizoom-smartwatch-dx3","https://www.vtech.com/","support@vtech.com","★★★★☆","子ども向けスマートウォッチ。カメラ・ゲーム・運動追跡・ストップウォッチ搭載。デジタルリテラシーを楽しく身につけられる。"],
  [1838,"子供・教育","Lovevery Play Gym & Kit","Lovevery","Lovevery公式","https://lovevery.com/","https://lovevery.com/","support@lovevery.com","★★★★★","発達心理学者監修の月齢別知育おもちゃ定期便。0〜3歳の脳発育を最大化するモンテッソーリ型玩具。子どもの発育に敏感な日本の親向け。"],
  [1839,"子供・教育","Moonlite Storybook Projector","Moonlite","Amazon US","https://moonlite.com/","https://moonlite.com/","info@moonlite.com","★★★★☆","スマートフォンに装着してリールを投影する絵本プロジェクター。天井・壁に絵本の物語をアニメーション投影。就寝前の読み聞かせを特別な体験に。"],
  [1840,"子供・教育","Elecrow CrowPi L Raspberry Pi Laptop","Elecrow","Amazon US","https://www.elecrow.com/crowpi-l.html","https://www.elecrow.com/","support@elecrow.com","★★★★★","15インチ画面・ラズベリーパイ搭載の子ども向けDIYプログラミング学習ラップトップ。Python・Scratch・Linux学習に最適。STEM教育の本格入門機。"],
  [1841,"子供・教育","LEGO Mindstorms Robot Inventor 51515","LEGO","Amazon US","https://www.lego.com/en-us/product/robot-inventor-51515","https://www.lego.com/","https://www.lego.com/en-us/service/help","★★★★★","5種以上のロボットをPythonでプログラミングできるLEGO教育ロボットキット。10〜14歳向けの本格STEMプログラミング体験。"],
  [1842,"子供・教育","Cozmo 2.0 Educational Robot","Digital Dream Labs","Amazon US","https://digitaldreamlabs.com/products/cozmo","https://digitaldreamlabs.com/","support@digitaldreamlabs.com","★★★★★","感情表現・顔認識・ゲーム・プログラミング学習機能を持つAIロボット。子どもと本物の感情的つながりを形成する世界唯一の教育ロボット。"],
  [1843,"子供・教育","Picoboard AI Coding Starter Kit","Picoboard","Amazon US","https://picoboard.io/","https://picoboard.io/","info@picoboard.io","★★★★☆","7歳以上向けのビジュアルブロックプログラミング+フィジカルコンピューティング統合学習キット。センサー・LED・モーターを実際に動かしながらプログラミングを学習。"],

  // ===== ファッション・アクセサリー (1844-1863) =====
  [1844,"ファッション・アクセサリー","Peuty Infinity Smart OLED Handbag","Peuty","Heyup公式","https://heyupnow.com/products/infinity-smart-color-changing-handbag","https://heyupnow.com/","info@peutydesign.com","★★★★★","CES 2026出展。フランス発・全粒本革製ハンドバッグにフレキシブルOLEDパネルを内蔵。アプリでリアルタイムデザイン変更可能。ファッション×テクノロジーの極致。"],
  [1845,"ファッション・アクセサリー","Nirva AI Jewelry Necklace & Bracelet","Nirva Labs","Nirva公式","https://www.nirva.life/","https://www.nirva.life/","hello@nirva.life","★★★★★","CES 2026出展。元Meta・Oculus開発チームが設計。一日を自動日記化・感情パターン追跡・社交マッピングができるAIジュエリー。見た目は高級アクセサリー。"],
  [1846,"ファッション・アクセサリー","Sensoria Fitness Smart Socks","Sensoria","Amazon US","https://www.sensoriafitness.com/smart-socks/","https://www.sensoriafitness.com/","info@sensoriafitness.com","★★★★★","テキスタイルセンサー内蔵の歩行分析スマートソックス。ランニングフォーム・歩行リズム・着地パターンをリアルタイム解析。スポーツ障害予防に革命的。"],
  [1847,"ファッション・アクセサリー","Ministry of Supply Velocity Suit","Ministry of Supply","Ministry of Supply公式","https://ministryofsupply.com/products/velocity-suit","https://ministryofsupply.com/","support@ministryofsupply.com","★★★★★","MITと共同開発した体温調節スマートスーツ。セルロース系繊維が体温に応じて自動的に換気性を調整。ビジネス着でも快適な体温管理。"],
  [1848,"ファッション・アクセサリー","Hexoskin Smart Shirt Biometric","Hexoskin","Amazon US","https://www.hexoskin.com/products/hexoskin-smart-shirt","https://www.hexoskin.com/","info@hexoskin.com","★★★★★","心拍・呼吸・活動量を継続的に計測する医療グレードスマートシャツ。NASAや軍用途でも採用実績。スポーツ・健康管理のプロ向け。"],
  [1849,"ファッション・アクセサリー","Epicore Biosystems Smart Sweat Patch","Epicore Biosystems","Epicore公式","https://www.epicorebiosystems.com/","https://www.epicorebiosystems.com/","info@epicorebiosystems.com","★★★★★","汗の成分（ナトリウム・塩化物・グルコース等）をリアルタイム分析するウェアラブルパッチ。スポーツ中の脱水・電解質バランスを科学的に管理。"],
  [1850,"ファッション・アクセサリー","Xenoma e-skin Smart Pajama","Xenoma","Xenoma公式","https://xenoma.com/","https://xenoma.com/","info@xenoma.com","★★★★★","日本発・体に貼り付いたような着心地で睡眠中の姿勢・呼吸・心拍を計測するスマートパジャマ。睡眠改善・医療モニタリングに革命的。"],
  [1851,"ファッション・アクセサリー","Vollebak Full Metal Jacket","Vollebak","Vollebak公式","https://www.vollebak.com/product/full-metal-jacket/","https://www.vollebak.com/","support@vollebak.com","★★★★★","グラフェン強化繊維・防弾ベスト並みの耐久性・防風防水性能を持つ超ハイテクアウタージャケット。極限環境の探検家・軍事向け素材を一般市場に。"],
  [1852,"ファッション・アクセサリー","Nadi X Smart Yoga Pants","Wearable X","Amazon US","https://www.wearablex.com/products/nadi-x-pant","https://www.wearablex.com/","hello@wearablex.com","★★★★★","AIコーチングで正確なヨガポーズを振動でリアルタイム誘導するスマートヨガパンツ。ヨガ人口が多い日本の女性向けに高い訴求力。"],
  [1853,"ファッション・アクセサリー","Solios Solar-Charged Minimalist Watch","Solios","Solios公式","https://solioswatches.com/","https://solioswatches.com/","hello@solioswatches.com","★★★★★","室内光で充電できるソーラーミニマリスト腕時計。電池交換不要・環境負荷ゼロ・スイス製ムーブメント搭載。サステナブル消費の日本市場向け。"],
  [1854,"ファッション・アクセサリー","Bellroy Lite Travel Wallet","Bellroy","Amazon US","https://bellroy.com/products/lite-travel-wallet","https://bellroy.com/","hello@bellroy.com","★★★★★","オーストラリア発・超スリム設計で海外旅行・パスポート・クレジットカードを一元管理できる機能的トラベルウォレット。日本人旅行者向け高需要。"],
  [1855,"ファッション・アクセサリー","Cotopaxi Allpa 35L Travel Pack","Cotopaxi","Amazon US","https://www.cotopaxi.com/products/allpa-35l-travel-pack","https://www.cotopaxi.com/","support@cotopaxi.com","★★★★★","世界1%の利益を貧困支援に寄付するB Corp認定ブランドのプレミアムトラベルバックパック。環境×機能美×社会貢献で日本の意識高い層に訴求。"],
  [1856,"ファッション・アクセサリー","Chrome Industries MXD Fathom Backpack","Chrome Industries","Amazon US","https://www.chromeindustries.com/products/mxd-fathom-backpack","https://www.chromeindustries.com/","support@chromeindustries.com","★★★★☆","自転車通勤者向けに設計された防水・耐久性特化型バックパック。シートベルトバックル・100%防水ファスナー搭載。サイクリスト向け製品として日本で高需要。"],
  [1857,"ファッション・アクセサリー","Aer Day Pack 3 X-Pac Backpack","Aer","Aer公式","https://www.aersf.com/products/day-pack-3-x-pac","https://www.aersf.com/","support@aersf.com","★★★★★","X-Pac素材採用の超軽量・耐水性ミニマリストデイパック。機能美と品質を重視する日本のビジネスパーソン向け高品質バッグ。"],
  [1858,"ファッション・アクセサリー","Tile Pro Bluetooth Item Tracker","Life360","Amazon US","https://www.tile.com/product/tile-pro","https://www.tile.com/","support@tile.com","★★★★★","Bluetooth信号で持ち物を検索できるスマートトラッカー。450m追跡範囲・水没防止・探したら音で知らせる。忘れ物が多い日本人向けに高需要。"],
  [1859,"ファッション・アクセサリー","Orbitkey Hybrid Laptop Sleeve 14","Orbitkey","Amazon US","https://www.orbitkey.com/products/hybrid-laptop-sleeve-14","https://www.orbitkey.com/","hello@orbitkey.com","★★★★★","キーオーガナイザー・ケーブル収納・カード入れを統合した多機能ラップトップスリーブ。テレワーク時代の携帯PCアクセサリーとして日本で高需要。"],
  [1860,"ファッション・アクセサリー","Moment Tectonic Chest Pack 8L","Moment","Amazon US","https://www.shopmoment.com/products/tectonic-chest-pack","https://www.shopmoment.com/","support@shopmoment.com","★★★★☆","カメラ・スマホ・アクセサリーを素早くアクセスできるチェストバッグ。写真家・旅行者・ストリートスナッパー向けに機能的設計。"],
  [1861,"ファッション・アクセサリー","NOMATIC Navigator 15L Commuter Pack","Nomatic","Amazon US","https://www.nomatic.com/products/navigator-15l-travel-pack","https://www.nomatic.com/","support@nomatic.com","★★★★☆","通勤・短期旅行に対応する機能重視型コンパクトバックパック。PCスリーブ・ウォータープルーフ・隠しポケット搭載。ビジネス通勤向け。"],
  [1862,"ファッション・アクセサリー","Courant Catch 3 Wireless Charger","Courant","Amazon US","https://courant.com/products/catch-3","https://courant.com/","info@courant.com","★★★★★","高級イタリアンレザーを使用した3台同時ワイヤレス充電パッド。インテリアとして映えるデザインでMagSafe対応。部屋のデスクを品良く整理整頓。"],
  [1863,"ファッション・アクセサリー","Ekster Parliament Slim Smart Wallet","Ekster","Amazon US","https://ekster.com/products/parliament-wallet","https://ekster.com/","support@ekster.com","★★★★★","ボタン一押しでカードが飛び出すスマートウォレット。AirTag対応・スリム設計・太陽光充電可能。スマートな持ち物管理で人気の高いオランダ発ブランド。"],

  // ===== クリーニング・収納・整理 (1864-1883) =====
  [1864,"クリーニング・収納・整理","ECOVACS DEEBOT X11 OMNICYCLONE","ECOVACS","Amazon US","https://www.ecovacs.com/us/deebot-robot-vacuum/deebot-x11-omnicyclone","https://www.ecovacs.com/","info@ecovacs.com","★★★★★","CES 2026 Innovation Award受賞。GaN急速充電・BLAST吸引・PureCyclone 2.0自動ゴミ捨て・AI音声アシストYIKO搭載の次世代ロボット掃除機。"],
  [1865,"クリーニング・収納・整理","Dreame Aqua10 Ultra Roller Robot Vacuum","Dreame","Amazon US","https://www.dreametech.com/products/dreame-aqua10-ultra","https://global.dreametech.com/","sales@dreame.tech","★★★★★","CES 2026 Innovation Award受賞。連続ローラーモップシステム搭載・床を洗いながら掃除するロボット掃除機。モップがけと掃除を同時にこなす革新製品。"],
  [1866,"クリーニング・収納・整理","Dreame H15 Pro Heat Wet Dry Vacuum","Dreame","Amazon US","https://www.dreametech.com/products/dreame-h15-pro-heat","https://global.dreametech.com/","sales@dreame.tech","★★★★★","CES 2026 Innovation Award受賞。熱水で粘着汚れを溶かしながら清掃するコードレスウェット・ドライ掃除機。高速気流で床面を即乾燥させる。"],
  [1867,"クリーニング・収納・整理","Roborock Saros Z70 Robot with Arm","Roborock","Amazon US","https://www.roborock.com/products/saros-z70","https://www.roborock.com/","support@roborock.com","★★★★★","CES 2026出展。ロボットアームで床の小物（靴下・コード等）を掴み除けてから掃除する世界初のロボット掃除機。散らかった部屋でも完全自律清掃。"],
  [1868,"クリーニング・収納・整理","Roborock Saros 20 Advanced Robot Vacuum","Roborock","Amazon US","https://www.roborock.com/products/saros-20","https://www.roborock.com/","support@roborock.com","★★★★★","AdaptiLift™ Chassis 3.0搭載・不整地・段差を乗り越えられる次世代ロボット掃除機。AI+LiDARで複雑な高級住宅でも完璧な自律清掃を実現。"],
  [1869,"クリーニング・収納・整理","LG Robot Vacuum with Built-In Station","LG Electronics","LG公式","https://www.lg.com/us/vacuums/robot-vacuums","https://www.lg.com/","https://www.lg.com/us/support/contact-us","★★★★★","CES 2026 Innovation Award受賞。ステーションを本体に内蔵したオールインワン設計のLGロボット掃除機。省スペース・完全ハンズフリーメンテナンス。"],
  [1870,"クリーニング・収納・整理","Narwal Freo X Ultra Robot Vacuum","Narwal","Amazon US","https://www.narwalhome.com/products/narwal-freo-x-ultra","https://www.narwalhome.com/","support@narwalhome.com","★★★★★","ゼロタングルブラシ・自動ゴミ捨て・自動洗浄・自動乾燥を完全自動化したロボット掃除機。完全ハンズフリーで頭に入れない日本の多忙な世帯向け。"],
  [1871,"クリーニング・収納・整理","iRobot Roomba Combo j9+ Robot","iRobot","Amazon US","https://store.irobot.com/roomba-combo/roomba-combo-j9-plus/4720845.html","https://www.irobot.com/","support@irobot.com","★★★★★","自動収納・自動洗浄付きのロボット掃除機＋モップがけ統合型。カーペットを検知してモップを自動収納する世界初機能搭載。"],
  [1872,"クリーニング・収納・整理","Lomi 2 Smart Food Waste Composter","Pela Earth","Amazon US","https://lomi.com/products/lomi-2","https://lomi.com/","support@lomi.com","★★★★★","生ゴミを3〜5時間で土に変えるスマートコンポスター。騒音・臭い・虫を完全シャットアウト。生ゴミ削減・SDGs意識の高い日本市場向けに革命的。"],
  [1873,"クリーニング・収納・整理","Simplehuman 58L Dual Compartment Sensor Can","Simplehuman","Amazon US","https://www.simplehuman.com/products/58-liter-voice-and-motion-sensor-can","https://www.simplehuman.com/","support@simplehuman.com","★★★★★","音声・モーションセンサーで自動開閉する2分割スマートゴミ箱。ゴミとリサイクルを自動分別。高品質・長寿命設計で日本市場の品質志向に合致。"],
  [1874,"クリーニング・収納・整理","IQAir HealthPro Plus Air Purifier","IQAir","IQAir公式","https://www.iqair.com/us/air-purifiers/healthpro-plus","https://www.iqair.com/","info@iqair.com","★★★★★","スイス発・世界の病院・医療施設が採用するHyperHEPA技術搭載最高性能空気清浄機。花粉症・PM2.5・ウイルス対策が重要な日本で最高水準の性能。"],
  [1875,"クリーニング・収納・整理","Coway Airmega 400S Smart Air Purifier","Coway","Amazon US","https://us.coway.com/products/airmega-400s","https://us.coway.com/","support@coway.com","★★★★★","Wi-Fi接続・リアルタイム空気質表示・自動強弱調整機能付き大空間対応スマート空気清浄機。花粉シーズン・梅雨の多湿時期に自動で対応する日本向け製品。"],
  [1876,"クリーニング・収納・整理","Blueair Blue Pure 311i+ Max","Blueair","Amazon US","https://www.blueair.com/us/air-purifiers/blue-pure-311i-plus-max/","https://www.blueair.com/","support@blueair.com","★★★★★","スウェーデン発・HEPASilent™超高性能静音空気清浄機。洗えるプレフィルターで経済的・省エネ。静音性が特に優れ在宅ワーク・睡眠中も使いやすい。"],
  [1877,"クリーニング・収納・整理","Tineco FLOOR ONE S9 Steam Vac","Tineco","Amazon US","https://us.tineco.com/products/floor-one-s9","https://us.tineco.com/","marketing@tineco.com","★★★★★","スチームモップ・ウェット・ドライ三機能一体型スマート床洗浄機。AIセンサーで汚れを検知し自動で洗浄液量を調整。清潔志向の日本市場に最適。"],
  [1878,"クリーニング・収納・整理","Mammotion SPINO S1 Pro Pool Cleaner","MAMMOTION","Amazon US","https://us.mammotion.com/products/spino-e1-robotic-pool-cleaner","https://us.mammotion.com/","distribution@mammotion.com","★★★★★","CES 2026出展。完全自律・自動ドック帰還機能付きロボットプールクリーナー。完全ハンズフリーで庭のプールを常時清潔に保てる。"],
  [1879,"クリーニング・収納・整理","Levoit Core 600S Smart Air Purifier","Levoit","Amazon US","https://levoit.com/products/levoit-core-600s-smart-true-hepa-air-purifier","https://levoit.com/","support@levoit.com","★★★★★","大型空間（63畳）対応・HEPAフィルター・スマートアプリ制御の高性能空気清浄機。コスパ最高水準でAmazon Best Seller多数獲得。"],
  [1880,"クリーニング・収納・整理","Dreo MC710S Smart Air Purifier","Dreo","Amazon US","https://www.dreosmart.com/products/dreo-mc710s","https://www.dreosmart.com/","support@dreosmart.com","★★★★☆","H13 True HEPAフィルター・VOC・CO2センサー搭載のスマート空気清浄機。アプリで空気質をリアルタイム監視。アレルギー・ペット対応の日本家庭向け。"],
  [1881,"クリーニング・収納・整理","iRobot Braava Jet m6 Robot Mop","iRobot","Amazon US","https://store.irobot.com/braava-jet-robot-mop/braava-jet-m6-robot-mop/braava-jet-m6.html","https://www.irobot.com/","support@irobot.com","★★★★★","精密噴水×振動清掃で硬いフロアを徹底的に清掃するロボットモップ。Roombaと連携し掃除完了後に自動でモップがけ開始する。日本のフローリング住宅に最適。"],
  [1882,"クリーニング・収納・整理","Raycop RHF-300 UV Futon Cleaner","Raycop","Amazon US","https://raycop.com/products/rhf-300","https://raycop.com/","support@raycop.com","★★★★★","紫外線+振動+強力吸引で布団・マットレスのダニ・アレルゲンを除去するスマート布団クリーナー。ダニ問題・花粉症・アレルギーが深刻な日本市場向けに最適。"],
  [1883,"クリーニング・収納・整理","Shark IQ Robot XL Self-Empty RV1001AE","Shark","Amazon US","https://www.sharkclean.com/shark-iq-robot-xl/","https://www.sharkclean.com/","support@sharkclean.com","★★★★★","自動ゴミ捨て機能付き・マッピング対応の高機能ロボット掃除機。ペット対応ブラシ・スマートホームアシスタント連動対応。コスパ優秀でAmazon人気No.1クラス。"],
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
console.log(`合計件数: ${1683 + added} 件`);
