/**
 * add_200_items_7.js — No.1884〜2083 (200件追加)
 * node add_200_items_7.js
 */

const XLSX = require('xlsx');
const path = require('path');

const INPUT_FILE  = path.join(__dirname, '海外便利グッズリスト_日本未上陸1905件_評価付.xlsx');
const OUTPUT_FILE = path.join(__dirname, '海外便利グッズリスト_日本未上陸2105件_評価付.xlsx');
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

  // ===== キッチン・調理器具 (1884-1903) =====
  [1884,"キッチン・調理器具","GE Profile Smart Refrigerator with Kitchen Assistant","GE Appliances","GE Appliances公式","https://www.geappliances.com/","https://www.geappliances.com/","info@geappliances.com","★★★★★","CES 2026出展。バーコードスキャン・AI食材認識・レシピ提案機能搭載の次世代スマート冷蔵庫。食品ロス削減と献立管理を一台で実現。"],
  [1885,"キッチン・調理器具","GoveeLife Smart Nugget Ice Maker Pro","Govee","Govee公式","https://www.govee.com/","https://www.govee.com/","marketing@govee.com","★★★★★","1日60ポンド製氷・動作音40dB・AI NoiseGuardテクノロジー搭載のスマート製氷機。CES 2026出展。静粛性と大容量を両立した日本向け製品。"],
  [1886,"キッチン・調理器具","Samsung Bespoke AI French Door Fridge AutoView","Samsung","Samsung公式","https://www.samsung.com/","https://www.samsung.com/","https://www.samsung.com/us/support/contact/","★★★★★","近づくと自動でガラス扉が照明する「AutoView」機能搭載ベスポーク冷蔵庫。CES 2026 Innovation Award受賞。フードショーケースとしての演出力が高い。"],
  [1887,"キッチン・調理器具","Samsung Bespoke AI Dishwasher 2026","Samsung","Samsung公式","https://www.samsung.com/","https://www.samsung.com/","https://www.samsung.com/us/support/contact/","★★★★☆","2026年型ベスポーク食洗機。従来比20%省エネ・AIが汚れ度に応じて洗浄プログラムを自動最適化。日本の省エネ志向に響く。"],
  [1888,"キッチン・調理器具","LG SIGNATURE 4-Door Fridge LCD Fresh Converter","LG Electronics","LG公式","https://www.lg.com/","https://www.lg.com/","https://www.lg.com/us/support/contact-us","★★★★★","6.8インチLCDディスプレイ・Fresh Converter+搭載の最上位冷蔵庫。CES 2026 Innovation Award受賞。食材の温湿度を自動管理。"],
  [1889,"キッチン・調理器具","Smeg Vitality System","Smeg","Smeg公式","https://www.smeg.com/","https://www.smeg.com/","info@smeg.it","★★★★★","急速冷却＋真空保存＋コンビスチームオーブンの3機能一体型システム。イタリアSmeg製で高いデザイン性と機能性を両立。食の質向上を求める富裕層向け。"],
  [1890,"キッチン・調理器具","GE Smart Ventilation Hood","GE Appliances","GE Appliances公式","https://www.geappliances.com/","https://www.geappliances.com/","info@geappliances.com","★★★★☆","NO2・空気質をリアルタイムモニタリングするスマートレンジフード。空気汚染センサー搭載で調理中の換気を自動制御。日本の住宅換気ニーズに合致。"],
  [1891,"キッチン・調理器具","Mellow Precision Sous Vide Cooker","Mellow","Mellow公式","https://www.cookmellow.com/","https://www.cookmellow.com/","hello@cookmellow.com","★★★★★","冷蔵・低温調理を一台で実現するスマート調理器。帰宅時間に合わせてアプリ操作で最適タイミングに調理開始。共働き世帯の夕食準備を革新。"],
  [1892,"キッチン・調理器具","Kuvings REVO830 Auto10 Juicer","Kuvings","Kuvings公式","https://kuvings.com/","https://kuvings.com/","info@kuvings.com","★★★★☆","Kickstarter 2025年トップ支援プロジェクト。自動連続搾汁・セルフクリーニング機能搭載の次世代低速ジューサー。健康志向の高い日本市場向け。"],
  [1893,"キッチン・調理器具","Crock-Pot 7Qt Express Crock Multi-Cooker","Crock-Pot","Amazon US","https://www.crock-pot.com/","https://www.crock-pot.com/","support@crock-pot.com","★★★★☆","圧力調理・スロークック・炊飯・蒸し調理など6機能を一台で実現するマルチクッカー。手軽な多機能調理器として日本市場でも需要あり。"],
  [1894,"キッチン・調理器具","Joule Turbo by ChefSteps","ChefSteps","ChefSteps公式","https://www.chefsteps.com/joule","https://www.chefsteps.com/","hello@chefsteps.com","★★★★★","業界最速クラスの低温調理器。ターボモードで通常の40%時間短縮。スマホアプリと連動しステップバイステップで料理をガイド。"],
  [1895,"キッチン・調理器具","Instant Pot Pro Plus Wi-Fi","Instant Brands","Amazon US","https://www.instantpot.com/","https://www.instantpot.com/","support@instantpot.com","★★★★☆","Wi-Fi接続・外出先からアプリ操作が可能な最上位マルチクッカー。世界的に人気のInstant Potが日本未上陸で国内ニーズ大。"],
  [1896,"キッチン・調理器具","Tovala Smart Steam Oven","Tovala","Tovala公式","https://www.tovala.com/","https://www.tovala.com/","hello@tovala.com","★★★★★","バーコードスキャンで自動調理設定するスマートスチームオーブン。ミールキット連携でプロ品質の食事が簡単に完成。"],
  [1897,"キッチン・調理器具","Multo Smart Cooking Robot","CookingPal","CookingPal公式","https://cookingpal.com/","https://cookingpal.com/","info@cookingpal.com","★★★★★","カメラ・センサー・AIを統合した次世代スマートクッキングロボット。切る・混ぜる・加熱の自動調理を実現。"],
  [1898,"キッチン・調理器具","Palate Smart Cutting Board","Palate","Kickstarter","https://getpalate.com/","https://getpalate.com/","hello@getpalate.com","★★★★☆","AIカメラ内蔵のスマートまな板。食材を自動認識しレシピを提案、カロリー計算も自動実行。健康管理に興味のある日本人向け。"],
  [1899,"キッチン・調理器具","Meater Block 4 Plus Wireless Thermometer","Meater","Meater公式","https://meater.com/","https://meater.com/","hello@meater.com","★★★★★","完全ワイヤレス4本セット肉用温度計。AIが肉の種類・厚みに応じて最適火入れをガイド。バーベキュー・調理愛好家向け。"],
  [1900,"キッチン・調理器具","AstroAI Smart Food Scale","AstroAI","Amazon US","https://astroai.com/","https://astroai.com/","support@astroai.com","★★★☆☆","栄養計算・カロリー記録アプリ連動のスマートキッチンスケール。糖質制限・カロリー管理に関心が高まる日本市場向け。"],
  [1901,"キッチン・調理器具","De'Longhi Rivelia La Specialista Arte Evo","De'Longhi","De'Longhi公式","https://www.delonghi.com/","https://www.delonghi.com/","https://www.delonghi.com/en/contact","★★★★★","デロンギ最高峰のエスプレッソマシン。AI自動グラインダー・水温精密制御・カフェメニュー15種搭載。コーヒー好きな日本人向け。"],
  [1902,"キッチン・調理器具","Zulay Kitchen Milk Frother Pro","Zulay Kitchen","Amazon US","https://zulaykitchen.com/","https://zulaykitchen.com/","support@zulaykitchen.com","★★★★☆","電動ミルクフォーマーの定番ブランド。複数の泡立てモードで本格カフェラテが自宅で実現。コーヒーカフェ文化が浸透する日本向け。"],
  [1903,"キッチン・調理器具","Typhur Dome Air Fryer","Typhur","Typhur公式","https://typhur.com/","https://typhur.com/","support@typhur.com","★★★★★","360度加熱・AIで食材を自動認識して最適温度・時間に設定するスマートエアフライヤー。時短調理と美味しさを両立。"],

  // ===== スマートホーム・インテリア・照明 (1904-1923) =====
  [1904,"スマートホーム・インテリア・照明","HEYMIRROR Smart Dressing Mirror","EONEOMS","EONEOMS公式","https://eoneoms.com/","https://eoneoms.com/","info@eoneoms.com","★★★★★","CES 2026 Innovation Award受賞のAIスマートドレッシングミラー。コーディネート提案・肌状態分析・天気連携で毎朝の準備を効率化。"],
  [1905,"スマートホーム・インテリア・照明","GE Matter-compatible Smart Shades","GE Lighting","GE Lighting公式","https://www.gelighting.com/","https://www.gelighting.com/","info@gelighting.com","★★★★☆","囁き音・充電式磁気バッテリー・ハブ不要のMatter対応スマートブラインド。スマートホーム標準規格Matter対応で将来性高い。"],
  [1906,"スマートホーム・インテリア・照明","Savant Smart Shades Matter over Thread","Savant Systems","Savant公式","https://www.savant.com/","https://www.savant.com/","info@savant.com","★★★★☆","Matter over Thread対応のバッテリー駆動スマートロールブラインド。工事不要で後付け設置が可能。賃貸住宅の多い日本に最適。"],
  [1907,"スマートホーム・インテリア・照明","Segway Master X Robotic Lawn Mower","Segway","Segway公式","https://www.segway.com/","https://www.segway.com/","info@segway.com","★★★★★","30インチアームで除草・果実収集・デブリ除去が可能なロボットアームモーア。CES 2026出展。庭管理の完全自動化を実現。"],
  [1908,"スマートホーム・インテリア・照明","VOVO Smart Toilet Neo TCB-090SA","VOVO","VOVO公式","https://vovousa.com/","https://vovousa.com/","support@vovousa.com","★★★★★","CES 2026 Innovation Award受賞。尿健康モニタリング機能搭載スマートトイレ。健康意識が高く温水洗浄便座文化のある日本で高い親和性。"],
  [1909,"スマートホーム・インテリア・照明","Govee Sky Ceiling Light Pro","Govee","Govee公式","https://www.govee.com/","https://www.govee.com/","marketing@govee.com","★★★★★","AIシーンモード・音楽同期・16万色対応のスマートシーリングライト。部屋全体を劇的に演出。スマートホーム入門商品として日本向け。"],
  [1910,"スマートホーム・インテリア・照明","Nanoleaf Lines Squared Smart Lights","Nanoleaf","Nanoleaf公式","https://nanoleaf.me/","https://nanoleaf.me/","support@nanoleaf.me","★★★★★","壁面に貼る幾何学スマートLEDパネル。音楽・映像連動・AI環境光モード搭載。インテリア照明として日本のZ世代に訴求。"],
  [1911,"スマートホーム・インテリア・照明","LIFX Color A60 Smart Bulb","LIFX","LIFX公式","https://www.lifx.com/","https://www.lifx.com/","support@lifx.com","★★★★☆","ハブ不要・1600万色対応・アプリ操作スマート電球。Matter対応で将来性高い。日本のスマート照明市場への参入製品として最適。"],
  [1912,"スマートホーム・インテリア・照明","Lockin V7 Max Smart Video Doorbell","Lockin","Lockin公式","https://www.lockin.com/","https://www.lockin.com/","info@lockin.com","★★★★★","AIフェイス認識・4K画質・双方向音声のスマートビデオドアベル。不在配達が多い日本の住宅事情に最適なセキュリティ製品。"],
  [1913,"スマートホーム・インテリア・照明","MAMMOTION LUBA 3 Robot Mower","MAMMOTION","MAMMOTION公式","https://mammotion.com/","https://mammotion.com/","support@mammotion.com","★★★★★","RTKナビゲーション・障害物自動回避のロボット芝刈り機。ワイヤー設置不要で庭の形状を自動マッピング。郊外住宅地向け製品。"],
  [1914,"スマートホーム・インテリア・照明","Eufy HomeVac S11 Infinity Cordless Vacuum","Eufy","Eufy公式","https://www.eufylife.com/","https://www.eufylife.com/","support@eufy.com","★★★★★","自動ゴミ収集ステーション連携のコードレス掃除機。充電しながら自動でゴミを収集。日本の小型住宅にも設置しやすいコンパクト設計。"],
  [1915,"スマートホーム・インテリア・照明","Tado Smart Thermostat X","Tado","Tado公式","https://www.tado.com/","https://www.tado.com/","support@tado.com","★★★★☆","Matter対応・位置情報連動・省エネAI搭載スマートサーモスタット。欧州シェアNo.1のスマート温度調節器。日本向け電力節約需要に合致。"],
  [1916,"スマートホーム・インテリア・照明","Philips Hue Gradient Signe Floor Lamp","Philips Hue","Philips Hue公式","https://www.philips-hue.com/","https://www.philips-hue.com/","support@philips-hue.com","★★★★★","画面内容に合わせて自動変色するアンビエント照明フロアランプ。映画鑑賞・ゲーム体験を劇的に向上。エンタメ需要の高い日本向け。"],
  [1917,"スマートホーム・インテリア・照明","Arlo Ultra 2 Spotlight Camera","Arlo","Arlo公式","https://www.arlo.com/","https://www.arlo.com/","support@arlo.com","★★★★★","4K画質・カラーナイトビジョン・AI人物検知のスマートセキュリティカメラ。空き巣対策需要が高まる日本で有望。"],
  [1918,"スマートホーム・インテリア・照明","Withings Body Scan Smart Scale","Withings","Withings公式","https://www.withings.com/","https://www.withings.com/","hello@withings.com","★★★★★","ECG・血管年齢・体組成・神経系コンディションを測定するスマート体重計。医療グレードの健康測定を家庭で実現。健康意識が高い日本向け。"],
  [1919,"スマートホーム・インテリア・照明","Amazon Echo Hub","Amazon","Amazon公式","https://www.amazon.com/","https://www.amazon.com/","https://www.amazon.com/contact-us","★★★★☆","Matter対応スマートホームコントロールパネル。全てのスマートデバイスを一画面で管理。Matter標準化でより重要になるハブデバイス。"],
  [1920,"スマートホーム・インテリア・照明","Brilliant Smart Home Control Panel","Brilliant","Brilliant公式","https://www.brilliant.tech/","https://www.brilliant.tech/","hello@brilliant.tech","★★★★★","壁スイッチ型タッチスクリーンのスマートホームコントロールセンター。Alexa内蔵で照明・温度・セキュリティを一元管理。"],
  [1921,"スマートホーム・インテリア・照明","Coway Airmega 400S Air Purifier","Coway","Coway公式","https://us.coway.com/","https://us.coway.com/","support@coway.com","★★★★★","HEPAフィルター・リアルタイム空気質モニター・スマホ連動の高性能空気清浄機。花粉症・PM2.5対策に敏感な日本市場に最適。"],
  [1922,"スマートホーム・インテリア・照明","Minut Point Smart Home Alarm","Minut","Minut公式","https://www.minut.com/","https://www.minut.com/","hello@minut.com","★★★★☆","騒音・温湿度・煙・動き・ガラス破砕音を検知するオールインワンスマートホームアラーム。月額不要で賃貸・民泊管理向けにも最適。"],
  [1923,"スマートホーム・インテリア・照明","Eve Energy Smart Plug Matter","Eve Systems","Eve公式","https://www.evehome.com/","https://www.evehome.com/","support@evehome.com","★★★☆☆","Matter対応でAndroid・iPhone両方から操作可能なスマートプラグ。電力使用量モニタリング機能搭載。スマートホーム入門商品として手軽。"],

  // ===== ウェアラブル・ヘルス・フィットネス (1924-1943) =====
  [1924,"ウェアラブル・ヘルス・フィットネス","Elemind Sleep Headband","Elemind Technologies","Elemind公式","https://elemindtech.com/","https://elemindtech.com/","info@elemindtech.com","★★★★★","MIT開発。脳波と同期した音響刺激で74%速い入眠を実現するスマートヘッドバンド。睡眠問題が深刻な日本人へのソリューション。"],
  [1925,"ウェアラブル・ヘルス・フィットネス","Shokz OpenFit Pro","Shokz","Shokz公式","https://shokz.com/","https://shokz.com/","support@shokz.com","★★★★★","CES 2026 Innovation Award受賞。開放型イヤホンでDolby Atmos・50時間バッテリー・ノイズキャンセリング機能を実現。ながら聴き需要の高い日本向け。"],
  [1926,"ウェアラブル・ヘルス・フィットネス","Abbott FreeStyle Libre + Libre Assist","Abbott","Abbott公式","https://www.freestyle.abbott/","https://www.abbott.com/","https://www.abbott.com/contactus.html","★★★★★","写真撮影だけで栄養・カロリー計算するAI連動の持続血糖モニター。糖尿病管理と一般健康管理の両方に対応。日本の糖尿病患者数増加背景に高需要。"],
  [1927,"ウェアラブル・ヘルス・フィットネス","Hypershell X Series Exoskeleton","Hypershell","Hypershell公式","https://hypershell.tech/","https://hypershell.tech/","info@hypershell.tech","★★★★★","歩行補助・疲労軽減・登山支援に対応するウェアラブルパワードスーツ。$899〜。高齢化社会の日本で歩行補助需要が高い。"],
  [1928,"ウェアラブル・ヘルス・フィットネス","Garmin Venu 4 Smartwatch","Garmin","Garmin公式","https://www.garmin.com/","https://www.garmin.com/","support@garmin.com","★★★★★","健康センサー・Body Battery・AI Morningコーチ搭載のプレミアムスマートウォッチ。ガーミンの最新フラッグシップ。健康管理需要の高い日本向け。"],
  [1929,"ウェアラブル・ヘルス・フィットネス","Bond Touch Haptic Bracelet","Bond Touch","Bond Touch公式","https://bondtouch.com/","https://bondtouch.com/","hello@bondtouch.com","★★★★☆","遠距離恋愛・家族間の触覚コミュニケーションを実現するスマートブレスレット。タップすると相手のブレスレットが振動。国際交流が増える日本でも需要あり。"],
  [1930,"ウェアラブル・ヘルス・フィットネス","Circular Ring 2 Smart Ring","Circular","Circular公式","https://circular.com/","https://circular.com/","support@circular.com","★★★★★","GPSトラッキング・ストレス測定・睡眠スコア分析搭載のスマートリング。指輪型ウェアラブルとして目立たずスタイリッシュ。日本でも需要急増中。"],
  [1931,"ウェアラブル・ヘルス・フィットネス","Withings ScanWatch Nova","Withings","Withings公式","https://www.withings.com/","https://www.withings.com/","hello@withings.com","★★★★★","ECG・血中酸素・睡眠無呼吸検知を持つ医療グレードのハイブリッドスマートウォッチ。アナログ時計デザインで普段使いしやすい。"],
  [1932,"ウェアラブル・ヘルス・フィットネス","Polar Vantage V3 Sports Watch","Polar","Polar公式","https://www.polar.com/","https://www.polar.com/","support@polar.com","★★★★★","マルチスポーツ対応・全身筋肉負荷測定・EEGを含む5センサー内蔵のプロ向けスポーツウォッチ。アスリート・トレーニング愛好家向け。"],
  [1933,"ウェアラブル・ヘルス・フィットネス","Coros Vertix 3 Adventure Watch","Coros","Coros公式","https://www.coros.com/","https://www.coros.com/","support@coros.com","★★★★☆","60日超バッテリー・二重バンド衛星GPS・標高気圧計搭載の極限アウトドア向けスマートウォッチ。登山・トレイルランの盛んな日本向け。"],
  [1934,"ウェアラブル・ヘルス・フィットネス","Amazfit Active 2 Smartwatch","Amazfit","Amazfit公式","https://www.amazfit.com/","https://www.amazfit.com/","support@amazfit.com","★★★★☆","AI健康コーチ・女性ホルモン周期管理・睡眠分析搭載のコスパ最高スマートウォッチ。高機能×低価格で日本市場に訴求。"],
  [1935,"ウェアラブル・ヘルス・フィットネス","Nurvv Run Smart Insoles","Nurvv","Nurvv公式","https://www.nurvv.com/","https://www.nurvv.com/","hello@nurvv.com","★★★★★","ランニングフォーム・足への荷重・ケガリスクをリアルタイム分析するスマートインソール。ランナー向けに怪我予防のコーチング機能搭載。"],
  [1936,"ウェアラブル・ヘルス・フィットネス","Muse S Gen 2 Meditation Headband","InteraXon","Muse公式","https://choosemuse.com/","https://choosemuse.com/","hello@choosemuse.com","★★★★★","脳波・心拍・呼吸を測定してリアルタイムフィードバックを提供する瞑想サポートヘッドバンド。マインドフルネス需要の高い日本向け。"],
  [1937,"ウェアラブル・ヘルス・フィットネス","Nix Hydration Biosensor","Nix Biosensors","Nix公式","https://nixbiosensors.com/","https://nixbiosensors.com/","support@nixbiosensors.com","★★★★☆","汗の成分をリアルタイム分析してパーソナライズ水分補給アドバイスを送るウェアラブルセンサー。スポーツ・熱中症対策として日本でも需要高。"],
  [1938,"ウェアラブル・ヘルス・フィットネス","Ultrahuman Ring AIR","Ultrahuman","Ultrahuman公式","https://www.ultrahuman.com/ring/","https://www.ultrahuman.com/","support@ultrahuman.com","★★★★★","月額料金不要のスマートリング。睡眠・心拍・HRVをトラッキングしAIが健康スコアを提供。Oura Ringの強力な競合として注目。"],
  [1939,"ウェアラブル・ヘルス・フィットネス","Biostrap Wrist Band EVO","Biostrap","Biostrap公式","https://biostrap.com/","https://biostrap.com/","support@biostrap.com","★★★★☆","医療グレードPPGセンサーで酸素飽和度・HRV・睡眠ステージを精密測定するウェアラブル。クリニカルグレードの健康トラッキング。"],
  [1940,"ウェアラブル・ヘルス・フィットネス","Sennheiser Momentum Sport True Wireless","Sennheiser","Sennheiser公式","https://www.sennheiser.com/","https://www.sennheiser.com/","support@sennheiser.com","★★★★★","心拍センサー内蔵・IP54防水・アダプティブANCのスポーツ向けTrue Wirelessイヤホン。ゾンビランとの連動機能搭載。スポーツ音楽好きな日本人向け。"],
  [1941,"ウェアラブル・ヘルス・フィットネス","Oura Ring Gen 4","Oura","Oura公式","https://ouraring.com/","https://ouraring.com/","support@ouraring.com","★★★★★","第4世代。より精度が向上した睡眠・心拍・体温・活動量トラッキングスマートリング。有名人・スポーツ選手も愛用するプレミアムブランド。"],
  [1942,"ウェアラブル・ヘルス・フィットネス","Withings U-Scan Urine Analyzer","Withings","Withings公式","https://www.withings.com/","https://www.withings.com/","hello@withings.com","★★★★★","トイレに設置するだけで尿検査・ビタミン・水分バランスをモニタリングするスマートデバイス。毎日の健康チェックを自動化。"],
  [1943,"ウェアラブル・ヘルス・フィットネス","Cionic Neural Sleeve 2","Cionic","Cionic公式","https://www.cionic.com/","https://www.cionic.com/","info@cionic.com","★★★★★","ニューロマスキュラー電気刺激で歩行を補助する機能的電気刺激スリーブ。脳卒中・MS患者の歩行機能回復に貢献。高齢化日本市場に需要大。"],

  // ===== アウトドア・スポーツ・旅行 (1944-1963) =====
  [1944,"アウトドア・スポーツ・旅行","Pongbot Aura AI Sports Robot","Pongbot Sports","Kickstarter","https://www.pongbotsports.com/","https://www.pongbotsports.com/","info@pongbotsports.com","★★★★★","AI搭載のテニス・ピックルボール・パデル対応スマートスポーツロボット。Kickstarter$1.6M調達。$499〜。一人練習革命。"],
  [1945,"アウトドア・スポーツ・旅行","Lumos Ultra Smart Bike Helmet","Lumos","Lumos公式","https://lumoshelmet.co/","https://lumoshelmet.co/","support@lumoshelmet.co","★★★★★","AI視認性システム・LED方向指示器・ブレーキランプ連動のコネクテッドスマートヘルメット。自転車通勤・アクティビティ需要の高い日本向け。"],
  [1946,"アウトドア・スポーツ・旅行","EcoFlow DELTA Pro 3 Portable Power Station","EcoFlow","EcoFlow公式","https://www.ecoflow.com/","https://www.ecoflow.com/","support@ecoflow.com","★★★★★","4000Wh大容量・双方向EV充電器互換・スマートホーム連携のポータブル電源。防災需要の高い日本市場でも注目されている製品。"],
  [1947,"アウトドア・スポーツ・旅行","Jackery Solar Generator 2000 Plus","Jackery","Jackery公式","https://www.jackery.com/","https://www.jackery.com/","support@jackery.com","★★★★★","3000Wh拡張可能・ソーラー充電対応のポータブル電源ステーション。キャンプ・防災需要の高い日本市場向けに高いポテンシャル。"],
  [1948,"アウトドア・スポーツ・旅行","DJI Avata 2 FPV Drone","DJI","DJI公式","https://www.dji.com/avata-2","https://www.dji.com/","support@dji.com","★★★★★","4K映像・FPVゴーグル付き・AIアクティブ追跡搭載のレース・映像撮影向けドローン。ドローン映像制作が普及する日本向け。"],
  [1949,"アウトドア・スポーツ・旅行","Insta360 X5 360 Camera","Insta360","Insta360公式","https://www.insta360.com/","https://www.insta360.com/","support@insta360.com","★★★★★","8K 360度撮影・AI自動編集・超広角アクションカム機能のアクションカメラ。旅行・スポーツ記録向けに日本でも人気拡大中。"],
  [1950,"アウトドア・スポーツ・旅行","Hydro Flask Insulated Tumbler 40oz","Hydro Flask","Amazon US","https://www.hydroflask.com/","https://www.hydroflask.com/","support@hydroflask.com","★★★★☆","24時間保冷・12時間保温の高機能ステンレスタンブラー。アウトドア・オフィス両用で使える定番ブランド。日本でのブランド認知度向上中。"],
  [1951,"アウトドア・スポーツ・旅行","YETI Hopper M30 Soft Cooler","YETI","Amazon US","https://www.yeti.com/","https://www.yeti.com/","support@yeti.com","★★★★★","防水・耐久性・保冷力No.1クラスのプレミアムソフトクーラーボックス。キャンプ・釣り・レジャー需要の高い日本のアウトドア市場向け。"],
  [1952,"アウトドア・スポーツ・旅行","Goal Zero Yeti 1500X Power Station","Goal Zero","Goal Zero公式","https://www.goalzero.com/","https://www.goalzero.com/","support@goalzero.com","★★★★☆","1516Whリチウム・ソーラー拡張対応・スマホ管理のポータブル電源。米国防災・キャンプ人気ブランド。日本の防災意識向上に合致。"],
  [1953,"アウトドア・スポーツ・旅行","Garmin inReach Mini 2 Satellite Communicator","Garmin","Garmin公式","https://www.garmin.com/inreach-mini-2","https://www.garmin.com/","support@garmin.com","★★★★★","衛星通信で圏外でもSOSメッセージ送信可能なサバイバルコミュニケーター。山岳遭難対策として登山者の多い日本で有望。"],
  [1954,"アウトドア・スポーツ・旅行","BioLite CampStove 2+ Wood Stove","BioLite","BioLite公式","https://www.bioliteenergy.com/","https://www.bioliteenergy.com/","support@bioliteenergy.com","★★★★☆","薪燃焼で発電・スマホ充電・電動ファン制御のキャンプストーブ。燃料不要のサバイバル向け製品。ソロキャンプブームの日本向け。"],
  [1955,"アウトドア・スポーツ・旅行","Cotopaxi Allpa 42L Travel Pack","Cotopaxi","Cotopaxi公式","https://www.cotopaxi.com/","https://www.cotopaxi.com/","support@cotopaxi.com","★★★★☆","航空機機内持ち込みサイズで機能的な旅行用バックパック。廃棄物素材使用の環境配慮型ブランド。エコ志向の日本向け旅行バッグ。"],
  [1956,"アウトドア・スポーツ・旅行","Osprey Farpoint 40 Travel Backpack","Osprey","Amazon US","https://www.osprey.com/","https://www.osprey.com/","support@osprey.com","★★★★★","米国バックパックブランドNo.1。軽量・パッカブル機内持ち込みサイズの旅行バッグ。旅行者・バックパッカー向け定番ブランド。"],
  [1957,"アウトドア・スポーツ・旅行","DJI Neo 2 Mini Drone","DJI","DJI公式","https://www.dji.com/neo","https://www.dji.com/","support@dji.com","★★★★★","250g以下・掌からの自律離陸・AIセルフィーモード搭載の超小型ドローン。旅行・Vlog撮影向けに日本市場でも高い需要が見込まれる。"],
  [1958,"アウトドア・スポーツ・旅行","Suunto Race S Multisport Watch","Suunto","Suunto公式","https://www.suunto.com/","https://www.suunto.com/","support@suunto.com","★★★★☆","AMOLED画面・精密GPS・トライアスロン対応のマルチスポーツウォッチ。鮮やかな画面と長バッテリーを両立。アスリート向け。"],
  [1959,"アウトドア・スポーツ・旅行","Anker Solix C1000 Portable Power Station","Anker","Anker公式","https://www.anker.com/","https://www.anker.com/","support@anker.com","★★★★★","1056Wh・2400W出力・LFPバッテリー搭載の高性能ポータブル電源。AnkerブランドはAmazon JapanでもNo.1シェアで日本展開しやすい。"],
  [1960,"アウトドア・スポーツ・旅行","Rokid Max AR Glasses","Rokid","Rokid公式","https://www.rokid.com/","https://www.rokid.com/","info@rokid.com","★★★★☆","50度視野角・215インチ仮想スクリーン・Spatial Audioのスマートグラス型モバイルシアター。通勤・旅行中のエンタメ用途に最適。"],
  [1961,"アウトドア・スポーツ・旅行","Sea Eagle 370 Pro Inflatable Kayak","Sea Eagle","Sea Eagle公式","https://www.seaeagle.com/","https://www.seaeagle.com/","support@seaeagle.com","★★★★☆","3人乗り・軽量折り畳み式インフレータブルカヤック。収納時はバッグに収まり車のトランクで運搬可能。海岸線が長い日本向け水上スポーツ製品。"],
  [1962,"アウトドア・スポーツ・旅行","Forclaz Travel 900 Ultralight Tent","Decathlon","Decathlon公式","https://www.decathlon.com/","https://www.decathlon.com/","https://www.decathlon.com/pages/contact-us","★★★★☆","950g超軽量・2人用・設営5分のトレッキング向けテント。コスパ最強のDecathlon製品として日本でも登山者・バックパッカーに響く。"],
  [1963,"アウトドア・スポーツ・旅行","Timeli Smart Trekking Pole","Timeli","Kickstarter","https://timeli.io/","https://timeli.io/","info@timeli.io","★★★★★","埋め込みセンサーでリアルタイム登山アシスタンス・疲労分析・ルートガイドを提供するスマートトレッキングポール。日本の登山人口向け。"],

  // ===== ペット用品 (1964-1983) =====
  [1964,"ペット用品","Fompet AI Pet Health Analyzer","Fompet","Kickstarter","https://fompet.com/","https://fompet.com/","info@fompet.com","★★★★★","CES 2026 Innovation Award受賞。15秒間のスキャンでペットの肥満度・健康リスクをAI分析するデバイス。動物病院費用を削減したい日本のペット飼育者向け。"],
  [1965,"ペット用品","WoofWolf Lux Robotic Dog Washing Machine","WoofWolf","WoofWolf公式","https://woofwooflux.com/","https://woofwooflux.com/","info@woofwooflux.com","★★★★★","AI制御で洗浄・乾燥まで20分で完了するロボット犬用洗濯機。グルーミング費用削減と愛犬のストレス軽減を両立。"],
  [1966,"ペット用品","Invoxia Minitailz Pet Tracker 2026","Invoxia","Invoxia公式","https://www.invoxia.com/","https://www.invoxia.com/","hello@invoxia.com","★★★★★","GPS+AI健康モニタリング（心拍・呼吸・活動）搭載のペットトラッカー最新版。圏外エリアもLTE対応でどこでも追跡可能。"],
  [1967,"ペット用品","Litter-Robot 4 Self-Cleaning Litter Box","Litter-Robot","Litter-Robot公式","https://www.litter-robot.com/","https://www.litter-robot.com/","support@litter-robot.com","★★★★★","自動回転・自動清掃・健康データモニタリング機能搭載のスマート自動猫トイレ。猫ブームの日本で圧倒的需要が見込まれる。"],
  [1968,"ペット用品","SATELLAI Collar Go GPS Cat Tracker","SATELLAI","SATELLAI公式","https://satellai.com/","https://satellai.com/","info@satellai.com","★★★★★","猫専用AI行動分析・GPSトラッキング・健康モニタリング搭載の首輪型デバイス。外出し猫の安全管理に敏感な日本の飼育者向け。"],
  [1969,"ペット用品","Furbo 360 Dog Camera","Furbo","Furbo公式","https://furbo.com/","https://furbo.com/","support@furbo.com","★★★★★","360度回転・AI吠え検知・自動おやつ投げ機能搭載のペットカメラ。日本でも人気のペット見守りカメラの最新モデル。"],
  [1970,"ペット用品","PetSafe Smart Feed Automatic Feeder","PetSafe","Amazon US","https://www.petsafe.net/","https://www.petsafe.net/","support@petsafe.net","★★★★☆","スマホ操作・スケジュール管理・食事量記録のスマート自動給餌器。留守番が多いペットの食事管理を自動化。共働き世帯向け。"],
  [1971,"ペット用品","Tractive GPS DOG 4 Tracker","Tractive","Tractive公式","https://tractive.com/","https://tractive.com/","support@tractive.com","★★★★★","リアルタイムGPS・活動量計・フェンスアラート搭載の犬用GPSトラッカー。欧州No.1シェアのペットGPSブランド。日本のペット迷子対策向け。"],
  [1972,"ペット用品","Petlibro Granary Automatic Pet Feeder","Petlibro","Amazon US","https://petlibro.com/","https://petlibro.com/","support@petlibro.com","★★★★☆","1080pカメラ内蔵・Wi-Fi連動・ログ記録付きスマート自動給餌器。リーズナブルな価格で高機能。コスパ重視の日本市場向け。"],
  [1973,"ペット用品","Wagz Freedom Smart Dog Collar","Wagz","Wagz公式","https://wagz.com/","https://wagz.com/","support@wagz.com","★★★★☆","GPS・AI音声トレーニング・仮想フェンス設定のスマード首輪。物理フェンスなしで安全なペット自由行動を実現。広い庭を持つ地方住宅向け。"],
  [1974,"ペット用品","SureFlap Microchip Cat Door Connect","SureFlap","SureFlap公式","https://www.surepetcare.com/","https://www.surepetcare.com/","support@surepetcare.com","★★★★★","マイクロチップ認証・カーフューモード・スマホ監視のスマートキャットドア。特定の猫だけ通過許可し安全を確保。"],
  [1975,"ペット用品","Petcube Bites 2 Lite Camera","Petcube","Petcube公式","https://petcube.com/","https://petcube.com/","hello@petcube.com","★★★★☆","リモートおやつ投げ・AI吠え検知・ペット家族との音声通話が可能なスマートペットカメラ。ペットの留守番対策として日本でも需要高。"],
  [1976,"ペット用品","Sureflap Pet Feeder Connect","SureFlap","SureFlap公式","https://www.surepetcare.com/","https://www.surepetcare.com/","support@surepetcare.com","★★★★★","マイクロチップで特定のペットのみ給餌するスマート自動給餌器。多頭飼いで食事管理が難しい場合に最適。"],
  [1977,"ペット用品","Cheerble Wicked Ball SE Cat Toy","Cheerble","Amazon US","https://cheerble.com/","https://cheerble.com/","support@cheerble.com","★★★★☆","AI制御で自律的に動き回る自動猫用ボール。3段階速度設定と自動休止機能搭載。共働き・一人暮らしのペット飼育者向け。"],
  [1978,"ペット用品","Motorola Halo+ Video Baby Monitor with Dog Camera","Motorola","Amazon US","https://www.motorola.com/","https://www.motorola.com/","support@motorola.com","★★★★☆","赤ちゃんとペットを同時モニタリングできる5インチスクリーン付きビデオモニター。育児とペット管理を一台で対応。"],
  [1979,"ペット用品","Paw5 Wooly Snuffle Mat","Paw5","Amazon US","https://paw5.com/","https://paw5.com/","hello@paw5.com","★★★★☆","犬の食事探索本能を刺激するノーズワークマット。ゆっくり食べさせることで早食い・肥満防止。ペットの精神的健康向けグッズ。"],
  [1980,"ペット用品","Garmob Automatic Dog Ball Launcher","Garmob","Amazon US","https://garmob.com/","https://garmob.com/","support@garmob.com","★★★★☆","室内外両対応・自動で2〜9m先にボールを投げる犬用自動ボールランチャー。留守番中の運動不足解消に最適。"],
  [1981,"ペット用品","AquaSense X Smart Water Fountain","AquaSense","Kickstarter","https://aquasense.pet/","https://aquasense.pet/","info@aquasense.pet","★★★★★","CES 2026 Innovation Award受賞。AIで水分摂取量をモニタリングするスマートペット給水器。腎臓病予防に重要な水分管理を自動化。"],
  [1982,"ペット用品","BioPet Vet Lab DNA & Health Test","BioPet","BioPet公式","https://www.biotype.com/","https://www.biotype.com/","info@biotype.com","★★★★☆","自宅で犬の犬種・遺伝病リスク・健康状態を調べるDNAテストキット。ペットの健康管理に積極的な日本の飼育者向け。"],
  [1983,"ペット用品","Whistle GO Explore GPS Pet Tracker","Whistle","Whistle公式","https://www.whistle.com/","https://www.whistle.com/","support@whistle.com","★★★★★","GPS+健康アラート+異常行動通知のペットトラッカー。ペットの位置と健康を一台で管理。米国最大シェアのペットトラッカーブランド。"],

  // ===== テクノロジー・ガジェット (1984-2003) =====
  [1984,"テクノロジー・ガジェット","XGIMI TITAN Noir 4K RGB Laser Projector","XGIMI","XGIMI公式","https://www.xgimi.com/","https://www.xgimi.com/","business@xgimi.com","★★★★★","Kickstarter$1,000万以上調達の4K RGBデュアルアイリスレーザープロジェクター。HDR10+・Dolby Vision対応。プレミアムシアター体験を家庭で実現。"],
  [1985,"テクノロジー・ガジェット","AWOL Vision Aetherion Max Triple Laser UST","AWOL Vision","AWOL Vision公式","https://awolvision.com/","https://awolvision.com/","support@awolvision.com","★★★★★","Kickstarter$1,470万調達の4K RGBトリプルレーザー超短焦点プロジェクター。壁から15cmで150インチ投影。ホームシアター革命製品。"],
  [1986,"テクノロジー・ガジェット","xTool WonderPress 5-in-1 Heat Press","xTool","xTool公式","https://www.xtool.com/","https://www.xtool.com/","support@xtool.com","★★★★★","Kickstarter$310万調達。熱転写・3Dサブリメーション・真空成形・DTFキュアリング・クラフトベーキング5機能搭載。DIY・ハンドメイド需要の高い日本向け。"],
  [1987,"テクノロジー・ガジェット","Revopoint POP 4 Pro 3D Scanner","Revopoint","Revopoint公式","https://www.revopoint3d.com/","https://www.revopoint3d.com/","support@revopoint3d.com","★★★★★","全素材対応・0.05mm精度の手持ち3Dスキャナー。Kickstarter出品で$579アーリーバード価格。3Dプリント・デザイン向け日本需要に合致。"],
  [1988,"テクノロジー・ガジェット","Samsung 130-inch Micro RGB TV","Samsung","Samsung公式","https://www.samsung.com/","https://www.samsung.com/","https://www.samsung.com/us/support/contact/","★★★★★","CES 2026 Best of Show受賞。130インチ・Micro RGBパネルの次世代最大テレビ。究極のホームシアター体験を提供するフラッグシップ製品。"],
  [1989,"テクノロジー・ガジェット","Mojie AR Glasses Pro","Mojie","Mojie公式","https://mojie.com/","https://mojie.com/","info@mojie.com","★★★★★","CES 2026出展のAR（拡張現実）スマートグラス。日常使い向け軽量デザインと高輝度ディスプレイを両立。AR市場の普及期製品として有望。"],
  [1990,"テクノロジー・ガジェット","Even Realities G2 Smart Glasses","Even Realities","Even Realities公式","https://www.evenrealities.com/","https://www.evenrealities.com/","support@evenrealities.com","★★★★★","通常の眼鏡と同等の重量・マイクロLED搭載のスマートグラス。ナビ・メッセージ・翻訳を視野に表示。日常使いARグラスとして革新的。"],
  [1991,"テクノロジー・ガジェット","XREAL Air 2 Ultra Mixed Reality Glasses","XREAL","XREAL公式","https://www.xreal.com/","https://www.xreal.com/","support@xreal.com","★★★★★","6DoF空間追跡・MR体験・3D空間マッピングのMixed Realityグラス。XREALは日本法人もあり日本市場へのアクセスが容易。"],
  [1992,"テクノロジー・ガジェット","Bambu Lab A1 mini 3D Printer Combo","Bambu Lab","Bambu Lab公式","https://bambulab.com/","https://bambulab.com/","support@bambulab.com","★★★★★","自動カラーチェンジ・AI補助・超高速印刷の家庭用3Dプリンター。Kickstarter系プロジェクトで世界的人気。DIY・ものづくり文化の日本向け。"],
  [1993,"テクノロジー・ガジェット","Boox Note Air3 C E-Ink Tablet","Onyx Boox","Boox公式","https://www.boox.com/","https://www.boox.com/","support@boox.com","★★★★★","カラーE-Ink・Android搭載のデジタルノートパッド。電子ペーパーの低眼精疲労で長時間読書・メモが可能。書籍文化の強い日本向け。"],
  [1994,"テクノロジー・ガジェット","Anker Prime Power Bank 27650mAh","Anker","Amazon US","https://www.anker.com/","https://www.anker.com/","support@anker.com","★★★★★","27650mAh・250W超高速充電・スマホ残量監視のスマートモバイルバッテリー。日本でも最大シェアのAnkerブランド最新フラッグシップ。"],
  [1995,"テクノロジー・ガジェット","Belkin BoostCharge Pro 4-in-1 Wireless Pad","Belkin","Amazon US","https://www.belkin.com/","https://www.belkin.com/","support@belkin.com","★★★★☆","Apple Watch・AirPods・iPhone・Android同時充電の4in1ワイヤレス充電パッド。デスク周りのケーブル整理に最適。"],
  [1996,"テクノロジー・ガジェット","Lenovo ThinkBook 16p Gen 4 AI Laptop","Lenovo","Lenovo公式","https://www.lenovo.com/","https://www.lenovo.com/","https://www.lenovo.com/us/en/contact/","★★★★★","NPU内蔵・AI機能統合・RTX 4060搭載のビジネス向けAIラップトップ。Copilot+PC対応で生産性大幅向上。テレワーク需要の日本向け。"],
  [1997,"テクノロジー・ガジェット","LG StanbyME 2 Wireless Portable TV","LG","LG公式","https://www.lg.com/","https://www.lg.com/","https://www.lg.com/us/support/contact-us","★★★★★","バッテリー内蔵・回転・傾き調整可能な27インチポータブルタッチスクリーンTV。どこでも設置できる場所を選ばないTV。日本の小型住宅向け。"],
  [1998,"テクノロジー・ガジェット","Logitech MX Master 3S Silent Mouse","Logitech","Amazon US","https://www.logitech.com/","https://www.logitech.com/","support@logitech.com","★★★★★","超静音クリック・電磁気スクロール・8000DPIのプロ向けワイヤレスマウス。在宅ワーク・オフィス向けに世界で評価が高い定番ブランド。"],
  [1999,"テクノロジー・ガジェット","Keychron Q6 Pro Mechanical Keyboard","Keychron","Keychron公式","https://www.keychron.com/","https://www.keychron.com/","support@keychron.com","★★★★★","アルミニウムボディ・Hot-Swap対応・カスタマイズ可能なプレミアム無線メカニカルキーボード。ゲーマー・プログラマー向けに日本でも人気上昇中。"],
  [2000,"テクノロジー・ガジェット","Elgato Prompter AI Teleprompter","Elgato","Amazon US","https://www.elgato.com/","https://www.elgato.com/","support@elgato.com","★★★★★","カメラ前に設置しAI字幕・原稿表示が可能なコンテンツクリエイター向けプロンプター。YouTube・動画配信ブームの日本向け製品。"],
  [2001,"テクノロジー・ガジェット","Plaud NotePin AI Voice Recorder","Plaud","Plaud公式","https://www.plaud.ai/","https://www.plaud.ai/","hello@plaud.ai","★★★★★","AI要約・翻訳・議事録自動生成が可能なウェアラブルAIボイスレコーダー。会議・インタビュー・勉強での記録・分析を自動化。ビジネス需要が高い。"],
  [2002,"テクノロジー・ガジェット","Bebird Note 5 Pro Ear Camera","Bebird","Amazon US","https://bebird.com/","https://bebird.com/","support@bebird.com","★★★★☆","1080Pカメラ付き耳洗浄ツール。スマホと接続してリアルタイムで耳の中を確認しながら安全に耳掃除。衛生意識の高い日本向け。"],
  [2003,"テクノロジー・ガジェット","Govee DreamView T2 TV Backlight","Govee","Govee公式","https://www.govee.com/","https://www.govee.com/","marketing@govee.com","★★★★★","AIカメラで画面色を解析して同期するアンビエントTV背面照明システム。映像体験を没入型に変える。コスパが高くホームシアター入門向け。"],

  // ===== 美容・スキンケア (2004-2023) =====
  [2004,"美容・スキンケア","L'Oréal Colorsonic Smart Dye Applicator","L'Oréal","L'Oréal公式","https://www.loreal.com/","https://www.loreal.com/","https://www.loreal.com/en/contact-us/","★★★★★","AIで髪色を均一に塗布するスマートヘアカラーアプリケーター。CES 2025出展の革新的セルフヘアカラーデバイス。美容意識が高い日本向け。"],
  [2005,"美容・スキンケア","Appotronics Laser Hair Cap","Appotronics","Appotronics公式","https://appotronics.com/","https://appotronics.com/","info@appotronics.com","★★★★★","医療グレードレーザーで育毛・発毛を促進するスマートヘアキャップ。CES 2026出展。男性型脱毛症に悩む日本男性市場に高需要。"],
  [2006,"美容・スキンケア","Omnilux Contour Face LED Mask","Omnilux","Omnilux公式","https://omniluxled.com/","https://omniluxled.com/","support@omniluxled.com","★★★★★","クリニカルグレードのLED光療法フェイスマスク。コラーゲン生成促進・シワ・たるみ改善。美容医療需要が高い日本市場向け。"],
  [2007,"美容・スキンケア","Foreo UFO 3 Smart Face Mask Device","Foreo","Foreo公式","https://www.foreo.com/","https://www.foreo.com/","support@foreo.com","★★★★★","T-Sonic振動・LED光・温冷ケアを組み合わせたスマートフェイスマスクデバイス。90秒でプロフェッショナルケアを実現。スキンケア先進国日本向け。"],
  [2008,"美容・スキンケア","Ziip Halo Nano Current Device","Ziip Beauty","Ziip公式","https://ziipbeauty.com/","https://ziipbeauty.com/","hello@ziipbeauty.com","★★★★★","ナノカレントと微弱電流でリフティング・輝き・抗老化効果を提供するスマート美容器。アプリ連動で200以上のトリートメント選択可能。"],
  [2009,"美容・スキンケア","NuFACE Trinity+ Facial Toning Device","NuFACE","NuFACE公式","https://www.mynuface.com/","https://www.mynuface.com/","support@mynuface.com","★★★★★","微弱電流で顔のリフティング・トーニングを行う美顔器。FDA認可取得済み。アンチエイジング美容への関心が高い日本女性向け。"],
  [2010,"美容・スキンケア","CurrentBody Skin LED Light Therapy Mask","CurrentBody","CurrentBody公式","https://www.currentbody.com/","https://www.currentbody.com/","hello@currentbody.com","★★★★★","臨床試験済みLEDフェイスマスク。赤色・近赤外線光でコラーゲン生成・肌トーン改善。セルフケア市場拡大中の日本向け。"],
  [2011,"美容・スキンケア","Dyson Supersonic r Hair Dryer","Dyson","Dyson公式","https://www.dyson.com/","https://www.dyson.com/","support@dyson.com","★★★★★","フォトニックセンサーで髪の状態を毎秒検知・AIが自動温度制御するDyson最新ヘアドライヤー。日本でもDysonブランドは絶大な人気。"],
  [2012,"美容・スキンケア","Tria Beauty 4X Laser Hair Removal","Tria Beauty","Tria公式","https://www.triabeauty.com/","https://www.triabeauty.com/","support@triabeauty.com","★★★★★","FDA認可・サロン品質の家庭用レーザー脱毛器。クリニック脱毛より低コストで永久脱毛を目指せる。脱毛需要が高い日本市場向け。"],
  [2013,"美容・スキンケア","Droplette Micro-Infuser Skincare Device","Droplette","Droplette公式","https://droplette.io/","https://droplette.io/","hello@droplette.io","★★★★★","マイクロドロップレット技術で成分を肌深部に届けるスキンケアデバイス。通常の100倍の浸透力を実現。スキンケア意識の高い日本人向け。"],
  [2014,"美容・スキンケア","Therabody TheraFace Mask","Therabody","Therabody公式","https://www.therabody.com/","https://www.therabody.com/","support@therabody.com","★★★★★","LED光療法・振動技術・高周波を組み合わせたTherabody製スマートフェイスマスク。スポーツ回復ブランドからの美容進出製品。"],
  [2015,"美容・スキンケア","MiXCELL AI Skin Analyzer","MiXCELL","MiXCELL公式","https://mixcell.ai/","https://mixcell.ai/","info@mixcell.ai","★★★★★","スマホカメラで肌状態をAI分析し個別スキンケアルーティンを提案するデバイス。パーソナライズ美容ニーズが高まる日本向け。"],
  [2016,"美容・スキンケア","Nira Precision Anti-Aging Laser","Nira","Nira公式","https://niraskin.com/","https://niraskin.com/","hello@niraskin.com","★★★★★","家庭用エントリーレベルのNd:YAGレーザーでシワ・コラーゲン生成を促進。クリニック施術と同等の波長を自宅で実現。コスト重視の日本市場向け。"],
  [2017,"美容・スキンケア","Reviive Lip Treatment Device","Reviive","Reviive公式","https://reviive.co/","https://reviive.co/","info@reviive.co","★★★★☆","赤色LED光療法・温熱効果でリップケアを行うスマートデバイス。乾燥・荒れ・色素沈着対策として日本の乾燥気候地域向け。"],
  [2018,"美容・スキンケア","Shark HyperAir Hair Dryer with IQ Brush","Shark Beauty","Shark公式","https://www.sharkclean.com/","https://www.sharkclean.com/","support@shark.com","★★★★★","IQブラシアタッチメント付き一体型ブロードライヤー。サロン仕上げを自宅で実現するDysonの強力競合製品。コスパで選ぶ日本人向け。"],
  [2019,"美容・スキンケア","Foreo Bear 2 Face Lifting Device","Foreo","Foreo公式","https://www.foreo.com/","https://www.foreo.com/","support@foreo.com","★★★★★","微弱電流・T-Sonic振動でフェイスラインを引き締めるスマート美顔器。アプリ連動で300以上のトリートメントプログラム提供。"],
  [2020,"美容・スキンケア","Skin Laundry Ionic Skin Cleansing System","Skin Laundry","Skin Laundry公式","https://skinlaundry.com/","https://skinlaundry.com/","info@skinlaundry.com","★★★★☆","イオン電流・音波振動・LEDを組み合わせた3合1スキンクレンジングシステム。深部洗浄・毛穴ケア・ブライトニング効果。"],
  [2021,"美容・スキンケア","PMD Pro Microdermabrasion Device","PMD Beauty","PMD公式","https://pmdbeauty.com/","https://pmdbeauty.com/","support@pmdbeauty.com","★★★★★","真空吸引・パテンテッドスピニングディスクで角質除去するマイクロダーマブレイジョンデバイス。クリニック品質を自宅で実現。"],
  [2022,"美容・スキンケア","Neutrogena Light Therapy Acne Spot Treatment","Neutrogena","Amazon US","https://www.neutrogena.com/","https://www.neutrogena.com/","https://www.neutrogena.com/pages/contact-us","★★★☆☆","青色LED光でニキビ菌を除去するスポット光線治療デバイス。FDA認可。薬剤不要の肌に優しいニキビ対策。日本のニキビケア需要向け。"],
  [2023,"美容・スキンケア","Therabody SmartGoggles Sleep & Relaxation","Therabody","Therabody公式","https://www.therabody.com/","https://www.therabody.com/","support@therabody.com","★★★★★","AIが心拍を検知してリアルタイムに温熱・振動を調整するスマートアイマスク。睡眠・リラクゼーション効果を科学的に実現。"],

  // ===== 子供・教育 (2024-2043) =====
  [2024,"子供・教育","Seedpace AI Story Robot","Seedpace","Seedpace公式","https://seedpace.com/","https://seedpace.com/","hello@seedpace.com","★★★★★","子供の発話・成長段階に合わせてリアルタイムでストーリーを生成するAIストーリーロボット。英語教育・創造力育成需要の日本向け。"],
  [2025,"子供・教育","Toniebox Smart Audio Player Gen 2","Tonies","Toniebox公式","https://us.tonies.com/","https://us.tonies.com/","support@tonies.com","★★★★★","フィギュアを置くだけで音楽・絵本・ポッドキャストが再生される子供向けスマートオーディオプレーヤー。スクリーンフリーで安全。幼児・保護者向け。"],
  [2026,"子供・教育","Sphero BOLT+ Programmable Robot","Sphero","Sphero公式","https://sphero.com/","https://sphero.com/","support@sphero.com","★★★★★","MATRIXLEDディスプレイ・プログラミング学習・ゲーム対応のスマートロボットボール。STEM教育向け製品として日本の教育市場で有望。"],
  [2027,"子供・教育","LEGO Mindstorms Alternative: Spike Prime","LEGO Education","LEGO Education公式","https://education.lego.com/","https://education.lego.com/","education@lego.com","★★★★★","Scratch・Python対応のLEGOロボティクス学習キット。小中学校でのプログラミング必修化を受け日本の教育市場でのニーズが高まっている。"],
  [2028,"子供・教育","KidSafe Kids Safe Wearable Watch","KidSafe","KidSafe公式","https://kidsafe.com/","https://kidsafe.com/","support@kidsafe.com","★★★★★","GPS追跡・SOS緊急ボタン・安全な通話機能を持つ子供向けスマートウォッチ。子供の安全意識が高い日本の親向け製品。"],
  [2029,"子供・教育","Moxie AI Social-Emotional Learning Robot","Embodied","Moxie公式","https://moxierobot.com/","https://moxierobot.com/","hello@moxierobot.com","★★★★★","子供の感情的知性・社会性・自信を育てるAIロボット。毎日の会話を通じてコミュニケーション能力向上をサポート。日本の教育親向け。"],
  [2030,"子供・教育","Learning Resources Coding Critters","Learning Resources","Amazon US","https://www.learningresources.com/","https://www.learningresources.com/","info@learningresources.com","★★★★☆","プログラミング概念を楽しく学べる動物型コーディングトイ。コードを入力して動物を動かす体験型STEM玩具。幼児・小学生向け。"],
  [2031,"子供・教育","Osmo Genius Starter Kit for iPad","Osmo","Osmo公式","https://www.playosmo.com/","https://www.playosmo.com/","support@playosmo.com","★★★★★","iPadカメラで現実の物体を認識しインタラクティブ学習を実現するゲームシステム。算数・国語・創造的思考力を遊びながら育成。"],
  [2032,"子供・教育","Matatalab Coding Robot Set","Matatalab","Matatalab公式","https://www.matatalab.com/","https://www.matatalab.com/","info@matatalab.com","★★★★★","スクリーンフリーでコーディングの基礎を学べる幼児向けロボットセット。プログラミング教育の低年齢化が進む日本市場向け。"],
  [2033,"子供・教育","KiwiCo Tinker Crate STEM Kit Subscription","KiwiCo","KiwiCo公式","https://www.kiwico.com/","https://www.kiwico.com/","hello@kiwico.com","★★★★★","月次配送の年齢別STEM工作キットサブスクリプション。手を動かして学ぶ体験型教育として日本でも理科離れ対策に有効。"],
  [2034,"子供・教育","Lovevery Play Kits for Children","Lovevery","Lovevery公式","https://lovevery.com/","https://lovevery.com/","hello@lovevery.com","★★★★★","発達段階に応じたモンテッソーリ式知育玩具サブスクリプション。生後0〜4歳向けに科学的根拠のある玩具を月次提供。"],
  [2035,"子供・教育","Artie 3000 Coding Robot","Educational Insights","Amazon US","https://www.educationalinsights.com/","https://www.educationalinsights.com/","support@educationalinsights.com","★★★★☆","お絵かきしながらプログラミングを学べるコーディングロボット。子供が描いたコードをロボットが実行して絵を描く。STEM教育向け。"],
  [2036,"子供・教育","PICO-8 Fantasy Console Education","Lexaloffle","PICO-8公式","https://www.lexaloffle.com/pico-8.php","https://www.lexaloffle.com/","hello@lexaloffle.com","★★★★☆","Lua言語でゲーム作成を学べるファンタジーコンソール。プログラミング・ゲームデザインを楽しく学ぶ。中高校生向けコーディング入門。"],
  [2037,"子供・教育","Wonder Workshop Dash Robot","Wonder Workshop","Amazon US","https://www.makewonder.com/","https://www.makewonder.com/","support@makewonder.com","★★★★★","Blockly・JavaScript・Pythonでプログラミング可能なK-5向けロボット。学校・塾のSTEM授業向けに米国で実績あり。日本教育市場向け。"],
  [2038,"子供・教育","IXL Learning Platform Subscription","IXL Learning","IXL公式","https://www.ixl.com/","https://www.ixl.com/","support@ixl.com","★★★★☆","AIで学習ギャップを特定しパーソナライズ問題を出題するK-12向けオンライン学習プラットフォーム。英語・算数を中心に全科目対応。"],
  [2039,"子供・教育","Piknik AI Language Learning for Kids","Piknik","Kickstarter","https://piknik.ai/","https://piknik.ai/","hello@piknik.ai","★★★★★","AIキャラクターとの会話で自然に英語・スペイン語などを習得できる子供向け言語学習アプリ。英語教育熱の高い日本の親向け。"],
  [2040,"子供・教育","Cozmo 2.0 AI Robot Toy","Digital Dream Labs","Amazon US","https://www.digitaldreamlabs.com/","https://www.digitaldreamlabs.com/","support@digitaldreamlabs.com","★★★★★","感情表現・ゲーム・プログラミング機能を持つAIロボットおもちゃ。Anki Cozmoの後継として子供向けAI教育玩具市場をリード。"],
  [2041,"子供・教育","MakeBlock mBot Neo Coding Robot","Makeblock","Makeblock公式","https://www.makeblock.com/","https://www.makeblock.com/","info@makeblock.com","★★★★★","Scratch・Python対応のSTEM学習用コーディングロボット。LEDマトリクス・センサー搭載でリアルな電子工作体験が可能。"],
  [2042,"子供・教育","Quizlet AI Study Assistant","Quizlet","Quizlet公式","https://quizlet.com/","https://quizlet.com/","https://quizlet.com/contact","★★★★☆","AIが弱点を分析してパーソナライズ学習計画を作成するデジタルフラッシュカード・学習プラットフォーム。受験・資格試験対策向け。"],
  [2043,"子供・教育","Headspace for Kids Mindfulness App","Headspace","Headspace公式","https://www.headspace.com/","https://www.headspace.com/","hello@headspace.com","★★★★☆","子供向け瞑想・マインドフルネス・睡眠ガイドアプリ。ADHD・不安・集中力向上のための習慣づくりに。子供のメンタルヘルス需要増加の日本向け。"],

  // ===== ファッション・アクセサリー (2044-2063) =====
  [2044,"ファッション・アクセサリー","Peuty Infinity Handbag Convertible","Peuty","Kickstarter","https://peuty.com/","https://peuty.com/","info@peuty.com","★★★★★","マグネット機構で形状を自在に変えるインフィニティハンドバッグ。1つのバッグが複数デザインに変身。ファッション感度の高い日本女性向け。"],
  [2045,"ファッション・アクセサリー","Nirva AI Jewelry Personalized Ring","Nirva","Nirva公式","https://nirva.ai/","https://nirva.ai/","info@nirva.ai","★★★★★","AIが個人の好み・肌色・スタイルを分析してカスタムデザインジュエリーを提案・製作するパーソナライズジュエリーブランド。"],
  [2046,"ファッション・アクセサリー","Sensoria Smart Fitness Socks","Sensoria Fitness","Sensoria公式","https://www.sensoriafitness.com/","https://www.sensoriafitness.com/","info@sensoriafitness.com","★★★★★","圧力センサー内蔵でランニングフォーム・歩行・足への負荷をリアルタイム分析するスマートソックス。ランナー・理学療法向け。"],
  [2047,"ファッション・アクセサリー","Tommy Hilfiger Adaptive Clothing Line","Tommy Hilfiger","Tommy Hilfiger公式","https://www.tommyhilfiger.com/","https://www.tommyhilfiger.com/","https://www.tommyhilfiger.com/en/customer-service","★★★★☆","磁気ボタン・ベルクロ・片手で着脱可能な設計の障がい者・高齢者向けアダプティブファッションライン。高齢化社会の日本に必要な製品。"],
  [2048,"ファッション・アクセサリー","Nomad Modern Leather Band for Apple Watch","Nomad","Amazon US","https://www.nomadgoods.com/","https://www.nomadgoods.com/","support@nomadgoods.com","★★★★☆","本革使用のプレミアムApple Watchバンド。Apple Watchの普及に伴い日本でも需要急増中のアクセサリー市場向け。"],
  [2049,"ファッション・アクセサリー","Bellroy Tech Kit Compact Organizer","Bellroy","Bellroy公式","https://bellroy.com/","https://bellroy.com/","support@bellroy.com","★★★★★","ケーブル・充電器・ガジェットを整理するコンパクトな旅行向けテックオーガナイザーポーチ。機能的かつスタイリッシュ。旅行好きな日本人向け。"],
  [2050,"ファッション・アクセサリー","Rothy's The Loafer Recycled Shoes","Rothy's","Rothy's公式","https://rothys.com/","https://rothys.com/","support@rothys.com","★★★★☆","ペットボトルリサイクル素材・機械洗濯可能・丈夫なサスティナブル靴。エコ志向が高まる日本市場でブランド認知度向上中。"],
  [2051,"ファッション・アクセサリー","Peak Design Travel Backpack 45L","Peak Design","Peak Design公式","https://www.peakdesign.com/","https://www.peakdesign.com/","support@peakdesign.com","★★★★★","モジュール式収納・魔法のポケット・機内持ち込み対応の旅行用バックパック。写真家・旅行者向けプレミアムブランド。"],
  [2052,"ファッション・アクセサリー","Brillen AI Sunglasses with Smart Lens","Intellilens","Kickstarter","https://intellilens.com/","https://intellilens.com/","info@intellilens.com","★★★★☆","AIで光量を自動調整するスマートサングラス。明暗変化に0.1秒以内で反応するポリマー液晶レンズ搭載。ドライブ・スポーツ向け。"],
  [2053,"ファッション・アクセサリー","Cariuma Oca Low Sustainable Sneakers","Cariuma","Cariuma公式","https://cariuma.com/","https://cariuma.com/","support@cariuma.com","★★★★☆","サトウキビ素材・天然ゴムソール・カーボンニュートラル認証のサスティナブルスニーカー。環境意識の高い日本の若者向けシューズ。"],
  [2054,"ファッション・アクセサリー","Native Union Stow Lite Travel Cable Kit","Native Union","Amazon US","https://www.nativeunion.com/","https://www.nativeunion.com/","support@nativeunion.com","★★★★☆","ケーブル・アダプター・充電器をスタイリッシュに整理するトラベルケーブルキット。シンプルなデザインと機能性で日本市場に訴求。"],
  [2055,"ファッション・アクセサリー","Tile Mate 2024 Bluetooth Tracker","Tile by Life360","Amazon US","https://www.tile.com/","https://www.tile.com/","support@tile.com","★★★★☆","スマホと連携してサイフ・カギ・バッグを追跡するBluetoothトラッカー。Apple AirTagとの比較で選択肢として日本でも需要あり。"],
  [2056,"ファッション・アクセサリー","Ekster Senate Slim RFID Wallet","Ekster","Ekster公式","https://ekster.com/","https://ekster.com/","support@ekster.com","★★★★★","ボタン一押しでカードがスライドして取り出せるRFIDブロッキング財布。Kickstarter人気製品。スキミング対策として日本でも有望。"],
  [2057,"ファッション・アクセサリー","Veja Campo Chrome Free Sneakers","Veja","Veja公式","https://www.veja-store.com/","https://www.veja-store.com/","contact@veja-store.com","★★★★☆","クロムフリーレザー・オーガニックコットン使用のエコフレンドリーフランス製スニーカー。セレブ愛用ブランドとして世界的人気。"],
  [2058,"ファッション・アクセサリー","Hermès Apple Watch Band Series","Hermès","Hermès公式","https://www.hermes.com/","https://www.hermes.com/","https://www.hermes.com/us/en/story/customer-service/","★★★★★","エルメス×Apple Watchのコラボレーションバンド。プレミアムファッションとテクノロジーの融合。高級ブランドに強い日本市場向け。"],
  [2059,"ファッション・アクセサリー","Dagne Dover Landon Carryall Tote","Dagne Dover","Dagne Dover公式","https://www.dagnedover.com/","https://www.dagnedover.com/","support@dagnedover.com","★★★★★","ネオプレン素材・完全防水・多機能インナーポケットのモダンキャリーオールバッグ。ワーキングウーマン向けにデザインされたプレミアムバッグ。"],
  [2060,"ファッション・アクセサリー","Ministry of Supply Performance Dress Shirt","Ministry of Supply","Ministry of Supply公式","https://ministryofsupply.com/","https://ministryofsupply.com/","support@ministryofsupply.com","★★★★☆","NASAフェイズチェンジ素材・4D伸縮・防シワのパフォーマンスドレスシャツ。ビジネス×コンフォートの革新的ウェア。テレワーク後のオフィス復帰向け。"],
  [2061,"ファッション・アクセサリー","Suicoke Moto-VHI Sandals","Suicoke","Suicoke公式","https://www.suicoke.com/","https://www.suicoke.com/","info@suicoke.com","★★★★☆","日本発の人気サンダルブランド。海外セレブ・ストリートブランドとのコラボで世界的知名度を獲得。日本ブランドの逆輸入パターン。"],
  [2062,"ファッション・アクセサリー","Moshi Altra Wrist Strap iPhone Case","Moshi","Moshi公式","https://www.moshi.com/","https://www.moshi.com/","support@moshi.com","★★★★☆","MagSafe対応・ストラップアタッチメント・落下保護のプレミアムiPhoneケース。台湾ブランドで日本でのブランド親和性が高い。"],
  [2063,"ファッション・アクセサリー","Cuyana System Tote Bag","Cuyana","Cuyana公式","https://www.cuyana.com/","https://www.cuyana.com/","support@cuyana.com","★★★★★","エコリネン・モジュール式インサート・カスタマイズ可能な総合機能トートバッグ。Less is more哲学で日本のシンプル美意識に合致。"],

  // ===== クリーニング・収納・整理 (2064-2083) =====
  [2064,"クリーニング・収納・整理","ECOVACS DEEBOT X11 Omni Robot Vacuum","ECOVACS","ECOVACS公式","https://www.ecovacs.com/","https://www.ecovacs.com/","support@ecovacs.com","★★★★★","自動モップ洗浄・乾燥・ゴミ収集の全自動ステーション付きロボット掃除機。AI障害物回避・マッピング精度が業界最高水準。"],
  [2065,"クリーニング・収納・整理","Dreame L20 Ultra Complete Robot Vacuum","Dreame Technology","Dreame公式","https://www.dreame-technology.com/","https://www.dreame-technology.com/","support@dreame-technology.com","★★★★★","延伸ロボットアームで家具下まで清掃・AI汚れ検知・ホットウォーター洗浄ステーション搭載の最上位ロボット掃除機。"],
  [2066,"クリーニング・収納・整理","Roborock Saros Z70 Robot Vacuum with Arm","Roborock","Roborock公式","https://www.roborock.com/","https://www.roborock.com/","support@roborock.com","★★★★★","CES 2025出展。物を掴める可動式ロボットアームを搭載した世界初のロボット掃除機。床掃除と同時に物の片付けを実現。"],
  [2067,"クリーニング・収納・整理","iRobot Roomba Combo j9+ Robot Vacuum Mop","iRobot","iRobot公式","https://www.irobot.com/","https://www.irobot.com/","support@irobot.com","★★★★★","吸引・水拭きコンボ・自動ゴミ収集・スマートホーム連携の最上位機種。ロボット掃除機パイオニアブランドの集大成。"],
  [2068,"クリーニング・収納・整理","Lomi 2 Smart Home Composter","Pela Earth","Lomi公式","https://lomi.com/","https://lomi.com/","hello@lomi.com","★★★★★","生ゴミを数時間で堆肥に変えるスマートコンポスター。臭いなし・静音・大容量の改良版。食品ロス削減意識の高い日本の都市住民向け。"],
  [2069,"クリーニング・収納・整理","Shark Detect Pro Robot Vacuum","Shark","Shark公式","https://www.sharkclean.com/","https://www.sharkclean.com/","support@shark.com","★★★★★","AI障害物検知・360度センサー・自動ゴミ収集ステーション搭載のプレミアムロボット掃除機。Roombaの強力競合。"],
  [2070,"クリーニング・収納・整理","Bissell SpinWave Cordless Hard Floor Spin Mop","Bissell","Amazon US","https://www.bissell.com/","https://www.bissell.com/","support@bissell.com","★★★★☆","コードレス・電動スピン回転モップ。フローリング・タイルの水拭き清掃を素早く実現。日本の住宅フローリング化に合わせた製品。"],
  [2071,"クリーニング・収納・整理","Miele Blizzard CX1 Bagless Vacuum","Miele","Miele公式","https://www.miele.com/","https://www.miele.com/","https://www.miele.com/en/com/contact.htm","★★★★★","ドイツMiele製サイクロン式掃除機。業界最高水準の吸引力・静音・耐久性20年設計。品質重視の日本市場向けプレミアム掃除機。"],
  [2072,"クリーニング・収納・整理","Method Laundry Detergent Plant-Based Pods","Method","Amazon US","https://methodproducts.com/","https://methodproducts.com/","hello@methodproducts.com","★★★★☆","植物由来成分・生分解性・濃縮タイプの洗濯洗剤ポッド。パッケージも100%再生素材。エコ洗剤需要の高まる日本向け。"],
  [2073,"クリーニング・収納・整理","OXO Good Grips Pop 2.0 Storage Containers","OXO","Amazon US","https://www.oxo.com/","https://www.oxo.com/","support@oxo.com","★★★★★","ワンタッチ密封・スタッキング可能・食洗機対応の食品保存コンテナセット。パントリー整理に最適。日本でも定評のある米国ブランド。"],
  [2074,"クリーニング・収納・整理","Dyson V15 Detect Absolute Vacuum","Dyson","Dyson公式","https://www.dyson.com/","https://www.dyson.com/","support@dyson.com","★★★★★","レーザーでホコリを可視化・AI微細粒子計測センサー搭載のDyson最上位コードレス掃除機。日本でDysonブランドは絶大な人気。"],
  [2075,"クリーニング・収納・整理","elfa Closet Swedish Organizer System","Container Store","Container Store公式","https://www.containerstore.com/","https://www.containerstore.com/","https://www.containerstore.com/s/contact","★★★★★","スウェーデン製カスタマイズ可能なモジュールクローゼットシステム。狭い住宅スペースを最大限活用できる北欧デザイン。小型住宅が多い日本向け。"],
  [2076,"クリーニング・収納・整理","iDesign Linus Multi-Use Drawer Organizer","iDesign","Amazon US","https://www.idesignliving.com/","https://www.idesignliving.com/","support@idesignliving.com","★★★★☆","透明・スタッキング可能・キッチン・浴室・デスクで使える多用途収納オーガナイザー。断捨離・整理整頓需要の高い日本向け定番製品。"],
  [2077,"クリーニング・収納・整理","Simplehuman Steel Frame Dish Rack","Simplehuman","Amazon US","https://www.simplehuman.com/","https://www.simplehuman.com/","support@simplehuman.com","★★★★★","ステンレス・折り畳み・水切りトレー付きのプレミアム水切りラック。機能美を追求したSimplehumanブランドが日本市場に最適。"],
  [2078,"クリーニング・収納・整理","Brabantia NewIcon Step Waste Bin","Brabantia","Amazon US","https://www.brabantia.com/","https://www.brabantia.com/","support@brabantia.com","★★★★☆","10年保証・ゆっくり閉まるソフトクローズ蓋・複数カラー展開のプレミアムペダルゴミ箱。デザイン重視の日本のキッチン向け。"],
  [2079,"クリーニング・収納・整理","Melitta Signature Auto Vacuum Self-Clean","Melitta","Melitta公式","https://www.melitta.com/","https://www.melitta.com/","https://www.melitta.com/en/service/contact/","★★★★☆","自動自己洗浄機能付きコーヒーグラインダー内蔵の掃除しやすいコーヒーメーカー。ドイツブランドの高品質家電として日本市場向け。"],
  [2080,"クリーニング・収納・整理","Yamazaki Tower Kitchen Organizer Set","Yamazaki Home","Amazon JP","https://www.yamazakihome.com/","https://www.yamazakihome.com/","support@yamazakihome.com","★★★★★","山崎実業のミニマルスチール収納シリーズ。海外でも人気急上昇中の日本ブランド。逆輸入パターンで海外向けに紹介できる珍しい事例。"],
  [2081,"クリーニング・収納・整理","Steamery Cirrus No.2 Clothes Steamer","Steamery","Steamery公式","https://steamery.com/","https://steamery.com/","hello@steamery.com","★★★★★","スウェーデン発の高級衣類スチーマー。シワ取り・消臭・除菌を2分で完了。ファッション意識が高く毎朝の服装を整える日本人向け。"],
  [2082,"クリーニング・収納・整理","Trashie Take Back Bag Textile Recycling","Trashie","Trashie公式","https://www.iamtrashie.com/","https://www.iamtrashie.com/","hello@iamtrashie.com","★★★★☆","古着・繊維製品を回収・リサイクルするエコテキスタイルリサイクルサービス。廃棄衣類問題に取り組みたい日本の消費者向け。"],
  [2083,"クリーニング・収納・整理","Molekule Air Pro Air Purifier","Molekule","Molekule公式","https://molekule.com/","https://molekule.com/","support@molekule.com","★★★★★","PECO技術でウイルス・VOC・アレルゲンを分子レベルで破壊する医療グレード空気清浄機。花粉症・アレルギー対策として日本向け最適。"],

];

// ===================== メイン処理 =====================
const wb = XLSX.readFile(INPUT_FILE);
const sheetName = wb.SheetNames[0];
const ws = wb.Sheets[sheetName];
const range = XLSX.utils.decode_range(ws['!ref']);

let rowIndex = range.e.r + 1; // 既存データの次の行

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

  // 製品URL
  const urlAddrP = XLSX.utils.encode_cell({ r: rowIndex, c: 5 });
  if (productUrl && productUrl.startsWith('http')) {
    ws[urlAddrP] = { v: productUrl, t: 's', l: { Target: productUrl } };
  } else {
    setCellVal(5, productUrl);
  }

  // メーカーHP
  const urlAddrM = XLSX.utils.encode_cell({ r: rowIndex, c: 6 });
  if (makerHp && makerHp.startsWith('http')) {
    ws[urlAddrM] = { v: makerHp, t: 's', l: { Target: makerHp } };
  } else {
    setCellVal(6, makerHp);
  }

  // メール or URL
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

// rangeを更新
range.e.r = rowIndex - 1;
ws['!ref'] = XLSX.utils.encode_range(range);

XLSX.writeFile(wb, OUTPUT_FILE);
console.log(`完了: ${OUTPUT_FILE}`);
console.log(`追加件数: ${NEW_PRODUCTS.length}件`);
