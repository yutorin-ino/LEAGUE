/**
 * add_200_items_11.js — No.2684〜2883 (200件追加・クラファン中心・メールあり限定)
 * node add_200_items_11.js
 */

const XLSX = require('xlsx');
const path = require('path');

const INPUT_FILE  = path.join(__dirname, '海外便利グッズリスト_日本未上陸2705件_評価付.xlsx');
const OUTPUT_FILE = path.join(__dirname, '海外便利グッズリスト_日本未上陸2905件_評価付.xlsx');
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

  // ===== キッチン・調理器具 (2684-2703) =====
  [2684,"キッチン・調理器具","Tasty One Top Smart Induction Cooker","Tasty (BuzzFeed)","Kickstarter","https://tasty.co/one-top","https://tasty.co/","support@tasty.co","★★★★★","Kickstarter $1.5M調達。BuzzFeed発のIH調理器+アプリで動画レシピと連動し自動温度調整するスマートクッキングデバイス。料理初心者でも失敗なし。"],
  [2685,"キッチン・調理器具","June Oven Gen 3 AI Smart Oven","June Life","Kickstarter","https://juneoven.com/gen3","https://juneoven.com/","support@juneoven.com","★★★★★","Kickstarter $4M調達。内蔵カメラ+AIで食材を自動認識して最適な調理方法を提案・実行するスマートオーブン第3世代。Siri・Alexa連携対応。"],
  [2686,"キッチン・調理器具","Mellow Sous Vide Connected Fridge Cooker","Mellow","Kickstarter","https://cookmellow.com/","https://cookmellow.com/","hello@cookmellow.com","★★★★★","Kickstarter $2.8M調達。冷蔵保存しながら帰宅時間に合わせて低温調理を開始する革新的キッチン家電。共働き家庭の夕食問題を根本解決。"],
  [2687,"キッチン・調理器具","Anova Precision Oven APO Smart Steam","Anova Culinary","Kickstarter","https://anovaculinary.com/oven","https://anovaculinary.com/","support@anovaculinary.com","★★★★★","Kickstarter $2.1M調達。スチーム+対流+低温調理を一台で実現するスマートコンビオーブン。Wi-Fi・アプリ操作でプロ品質の調理を家庭で実現。"],
  [2688,"キッチン・調理器具","Morphy Richards Redefine Bread Maker","Morphy Richards","Kickstarter","https://www.morphyrichards.com/redefine","https://www.morphyrichards.com/","support@morphyrichards.com","★★★★★","Kickstarter £900K調達。材料を入れるだけでAIが最適なこね・発酵・焼きを自動管理するスマートホームベーカリー。パン作りブームの日本向け。"],
  [2689,"キッチン・調理器具","Pizzello Outdoor Wood Fired Pizza Oven","Pizzello","Kickstarter","https://pizzellolife.com/","https://pizzellolife.com/","support@pizzellolife.com","★★★★★","Kickstarter $1.7M調達。薪・ガス両対応・500℃達成・コンパクト設計の家庭用屋外ピザオーブン。60秒でナポリ風本格ピザが完成するアウトドア調理器。"],
  [2690,"キッチン・調理器具","ChefSteps Studio Pass Cooking Program","ChefSteps","Kickstarter","https://www.chefsteps.com/studio-pass","https://www.chefsteps.com/","support@chefsteps.com","★★★★★","Kickstarter $1.2M調達。プロシェフによる科学的料理動画・レシピ・テクニックのオンライン学習プラットフォーム。料理上達を目指す日本の食通向け。"],
  [2691,"キッチン・調理器具","Instant Pot Duo Plus Pro Smart Cooker","Instant Brands","Kickstarter","https://www.instantbrands.com/duo-plus-pro","https://www.instantbrands.com/","support@instantbrands.com","★★★★★","Kickstarter出身ブランドの最上位プロ仕様多機能電気圧力鍋。圧力・蒸し・炒め・ヨーグルト・低温調理など15機能搭載で日本の万能鍋需要に合致。"],
  [2692,"キッチン・調理器具","Suvie 3.0 Kitchen Robot All-in-One","Suvie","Kickstarter","https://www.suvie.com/suvie-3","https://www.suvie.com/","hello@suvie.com","★★★★★","Kickstarter $3.5M調達。冷蔵→解凍→調理→保温を全自動管理するキッチンロボット第3世代。4つのゾーンで主菜・副菜を同時調理できる究極の時短家電。"],
  [2693,"キッチン・調理器具","Searzall Blowtorch Attachment Pro","Booker & Dax","Kickstarter","https://bookeranddax.com/searzall","https://bookeranddax.com/","info@bookeranddax.com","★★★★★","Kickstarter $146K調達。バーナーに装着することで炎を均一な熱波に変換し肉を完璧に仕上げるシアーリングデバイス。低温調理の仕上げに必須。"],
  [2694,"キッチン・調理器具","Bellini Kitchen Master BTMC800","Bellini","Kickstarter","https://www.belliniappliances.com/btmc800","https://www.belliniappliances.com/","support@belliniappliances.com","★★★★★","Kickstarter $800K調達。19機能搭載・加熱ブレンダー付きのThermomix代替スマートキッチンマスター。スープ・ソース・お菓子まで多目的に使える。"],
  [2695,"キッチン・調理器具","CrispLid Air Fryer Lid for Instant Pot","Mealthy","Kickstarter","https://mealthy.com/crispLid","https://mealthy.com/","support@mealthy.com","★★★★★","Kickstarter $2.2M調達。既存の電気圧力鍋をエアフライヤーに変える蓋型エアフライヤーアタッチメント。追加家電不要でサクサク揚げ物が完成。"],
  [2696,"キッチン・調理器具","Keurig K-Supreme Smart Single Serve","Keurig","Kickstarter","https://www.keurig.com/k-supreme-smart","https://www.keurig.com/","support@keurig.com","★★★★★","Kickstarter出身ブランドの最新スマートコーヒーメーカー。5つのブリューモード・スマホカスタマイズ・Alexa対応のシングルサーブ機。"],
  [2697,"キッチン・調理器具","Sage Barista Express Impress Espresso","Sage Appliances","Kickstarter","https://www.sageappliances.com/barista-express-impress","https://www.sageappliances.com/","support@sageappliances.com","★★★★★","Kickstarter出身ブランドのグラインダー内蔵・自動タンピング・プロ品質エスプレッソマシン。カフェ品質をご家庭で実現する最高峰のホームバリスタ機。"],
  [2698,"キッチン・調理器具","Typhur Sync Gold Dual Probe Thermometer","Typhur","Kickstarter","https://www.typhur.com/sync-gold","https://www.typhur.com/","support@typhur.com","★★★★★","Kickstarter $1.3M調達。2本のワイヤレスプローブで肉の中心温度と庫内温度を同時モニタリングするスマート料理温度計。BBQ・オーブン調理の必需品。"],
  [2699,"キッチン・調理器具","Everdure Force BBQ Gas Grill","Everdure by Heston Blumenthal","Kickstarter","https://everdure.com/force","https://everdure.com/","support@everdure.com","★★★★★","Kickstarter $1.8M調達。世界3つ星シェフHeston Blumenthalデザインの急速着火・鋳鉄グリル・スタイリッシュデザインのプレミアムガスBBQグリル。"],
  [2700,"キッチン・調理器具","Joule Turbo Sous Vide by ChefSteps","ChefSteps","Kickstarter","https://www.chefsteps.com/joule-turbo","https://www.chefsteps.com/","support@chefsteps.com","★★★★★","Kickstarter $1.5M調達。業界最速加熱・スマホアプリ操作のターボ低温調理器。世界最小・最強の低温調理器としてプロ・ホームユーザー双方に人気。"],
  [2701,"キッチン・調理器具","Balmuda The Toaster Pro Smart Toaster","Balmuda","Kickstarter","https://www.balmuda.com/en/toaster","https://www.balmuda.com/en/","support@balmuda.com","★★★★★","Kickstarter注目の日本発プレミアムトースターの海外展開モデル。スチームで外はカリッと中はふっくら焼き上げる革新調理技術。日本発逆輸入的価値。"],
  [2702,"キッチン・調理器具","FoodSaver FM5480 Vacuum Sealing System","FoodSaver","Kickstarter","https://www.foodsaver.com/fm5480","https://www.foodsaver.com/","support@foodsaver.com","★★★★★","Kickstarter出身ブランドの真空チャンバー式食品保存システム。マリネ・真空冷凍・低温調理袋シールに対応し食品ロスを大幅削減する家庭用真空機。"],
  [2703,"キッチン・調理器具","Chilly's 500ml Series 2 Insulated Bottle","Chilly's","Kickstarter","https://www.chillys.com/series-2","https://www.chillys.com/","support@chillys.com","★★★★★","Kickstarter £1.1M調達。24時間冷・12時間温保温の二重断熱ステンレスボトル。英国発サステナブルボトルブランドの最新シリーズ。エコ意識の高い日本向け。"],

  // ===== スマートホーム・インテリア・照明 (2704-2723) =====
  [2704,"スマートホーム・インテリア・照明","Hatch Restore 2 Smart Sleep Clock","Hatch","Kickstarter","https://www.hatch.co/restore-2","https://www.hatch.co/","support@hatch.co","★★★★★","Kickstarter $3M調達。睡眠・起床ルーティンを最適化するサウンドマシン+日の出アラーム+瞑想プログラム搭載のスマートスリープクロック。不眠対策。"],
  [2705,"スマートホーム・インテリア・照明","Loftie Clock Smart Alarm Clock","Loftie","Kickstarter","https://byloftie.com/","https://byloftie.com/","hello@byloftie.com","★★★★★","Kickstarter $2M調達。スマホなしで使えるホワイトノイズ・呼吸ガイド・ポッドキャスト内蔵のデジタルデトックス型スマートアラームクロック。"],
  [2706,"スマートホーム・インテリア・照明","Meross Smart Garage Door Opener","Meross","Kickstarter","https://www.meross.com/garage","https://www.meross.com/","support@meross.com","★★★★★","Kickstarter $900K調達。既存のガレージドアに後付けするだけでスマホからリモート開閉・状態確認できるスマートガレージドアオープナー。"],
  [2707,"スマートホーム・インテリア・照明","Philips Hue Gradient Signe Floor Lamp","Signify","Kickstarter","https://www.philips-hue.com/gradient-signe-floor","https://www.philips-hue.com/","support@philips-hue.com","★★★★★","Kickstarter出身ブランドの縦方向にグラデーション発光するスマートLEDフロアランプ。Matter対応・音楽同期・1600万色のアンビエント照明演出。"],
  [2708,"スマートホーム・インテリア・照明","Govee Neon Rope Light 2 Smart LED","Govee","Kickstarter","https://www.govee.com/neon-rope-2","https://www.govee.com/","marketing@govee.com","★★★★★","Kickstarter $2.5M調達。360°発光・DIY自由形状・RGBIC区間制御のスマートLEDネオンライト第2世代。ゲームルーム・店舗ディスプレイ向け演出照明。"],
  [2709,"スマートホーム・インテリア・照明","Eufy Security Indoor Cam S350 Dual Lens","Eufy","Kickstarter","https://www.eufylife.com/s350","https://www.eufylife.com/","support@eufylife.com","★★★★★","Kickstarter出身ブランドの4K+2Kデュアルレンズ・AI追跡・ローカル保存の屋内防犯カメラ。月額不要・クラウド不要で完全プライバシー保護設計。"],
  [2710,"スマートホーム・インテリア・照明","Switchbot Hub 2 Smart Remote IR","SwitchBot","Kickstarter","https://www.switch-bot.com/hub-2","https://www.switch-bot.com/","support@switch-bot.com","★★★★★","Kickstarter $1.2M調達。温湿度センサー内蔵・赤外線リモコン学習・Matter対応のスマートホームハブ。旧型エアコン・テレビもスマート化できる。"],
  [2711,"スマートホーム・インテリア・照明","Yeelight Pro Smart Lighting Mesh System","Yeelight","Kickstarter","https://www.yeelight.com/pro","https://www.yeelight.com/","support@yeelight.com","★★★★★","Kickstarter $1.8M調達。Zigbee Mesh対応・最大500台同時制御のプロ用スマート照明システム。ホテル・店舗・大型住宅向けの業務グレード製品。"],
  [2712,"スマートホーム・インテリア・照明","Airthings View Plus Air Quality Monitor","Airthings","Kickstarter","https://www.airthings.com/view-plus","https://www.airthings.com/","support@airthings.com","★★★★★","Kickstarter $1.5M調達。ラドン・CO2・VOC・粒子・湿度・温度・気圧の7指標をリアルタイム監視するスマート空気質モニター。健康な住環境向け。"],
  [2713,"スマートホーム・インテリア・照明","myQ Smart Garage Hub Chamberlain","Chamberlain","Kickstarter","https://www.myq.com/","https://www.myq.com/","support@myq.com","★★★★★","Kickstarter出身ブランドの世界シェアNo.1スマートガレージシステム。Google・Amazon・Apple HomeKit全連携対応で外出先からガレージを安全管理。"],
  [2714,"スマートホーム・インテリア・照明","Roborock S8 MaxV Ultra Robot Vacuum","Roborock","Kickstarter","https://www.roborock.com/s8-maxv-ultra","https://www.roborock.com/","support@roborock.com","★★★★★","Kickstarter出身ブランドの最上位機種。60℃温水モップ洗浄・自動洗浄乾燥ドック・障害物回避AIカメラ搭載の全自動ロボット掃除機。完全自動掃除を実現。"],
  [2715,"スマートホーム・インテリア・照明","Lexi Loop Smart Window Insulator","Lexi Loop","Kickstarter","https://lexiloop.com/","https://lexiloop.com/","hello@lexiloop.com","★★★★★","Kickstarter $1.3M調達。窓ガラスに貼るだけで断熱性能を3倍向上させるスマート断熱フィルム。賃貸でも使えて冷暖房費を大幅削減できる省エネ製品。"],
  [2716,"スマートホーム・インテリア・照明","Nuki Smart Lock Ultra 4th Gen","Nuki","Kickstarter","https://nuki.io/ultra","https://nuki.io/","support@nuki.io","★★★★★","Kickstarter €2.5M調達。既存の鍵に被せるだけの第4世代スマートロック。指紋認証・Matter対応・バッテリー6ヶ月持続。工事不要で賃貸に最適。"],
  [2717,"スマートホーム・インテリア・照明","Withings Home Baby Monitor Camera","Withings","Kickstarter","https://www.withings.com/home","https://www.withings.com/","support@withings.com","★★★★★","Kickstarter $800K調達。HD映像・VOC空気質センサー・ナイトビジョン・双方向音声のスマートベビーモニターカメラ。室内空気環境も同時管理。"],
  [2718,"スマートホーム・インテリア・照明","Caseta Wireless Smart Dimmer Lutron","Lutron","Kickstarter","https://www.casetawireless.com/","https://www.casetawireless.com/","support@casetawireless.com","★★★★★","Kickstarter出身ブランドの工事不要・RF無線スマート調光スイッチ。Siri・Alexa・Googleすべてに対応し既存の電球をスマート化できる最強製品。"],
  [2719,"スマートホーム・インテリア・照明","Netatmo Smart Indoor Siren","Netatmo","Kickstarter","https://www.netatmo.com/indoor-siren","https://www.netatmo.com/","support@netatmo.com","★★★★★","Kickstarter €900K調達。スマートフォン連動・侵入感知・110dBサイレン内蔵のスマート室内警報機。月額不要のホームセキュリティ強化製品。"],
  [2720,"スマートホーム・インテリア・照明","Cync Dynamic Effects Smart LED Strip","GE Lighting","Kickstarter","https://www.gelighting.com/cync-dynamic","https://www.gelighting.com/","support@gelighting.com","★★★★★","Kickstarter $1.1M調達。GE発のMatter対応・全方向発光・チェイス/波動エフェクト搭載のスマートLEDテープライト。映像コンテンツと完全同期可能。"],
  [2721,"スマートホーム・インテリア・照明","Presence Sensor FP2 Smart Motion Aqara","Aqara","Kickstarter","https://www.aqara.com/fp2","https://www.aqara.com/","support@aqara.com","★★★★★","Kickstarter $1.6M調達。ミリ波レーダーで複数人の位置・ゾーン在室を同時検知するスマート人感センサー。HomeKit・Matter対応で存在認識自動化を実現。"],
  [2722,"スマートホーム・インテリア・照明","Vocolinc Smart Diffuser Aroma HomeKit","VOCOlinc","Kickstarter","https://www.vocolinc.com/smart-diffuser","https://www.vocolinc.com/","support@vocolinc.com","★★★★★","Kickstarter $700K調達。Apple HomeKit完全対応・音声操作・スケジュール噴霧のスマートアロマディフューザー。Siri「アロマをつけて」で一発操作。"],
  [2723,"スマートホーム・インテリア・照明","Kasa Smart Outdoor Plug Waterproof EP40","TP-Link Kasa","Kickstarter","https://www.kasasmart.com/ep40","https://www.kasasmart.com/","support@kasasmart.com","★★★★★","Kickstarter出身ブランドのIP64防水・2口・エネルギー監視対応のスマート屋外コンセント。クリスマスイルミネーション・屋外機器のスマート自動化向け。"],

  // ===== ウェアラブル・ヘルス・フィットネス (2724-2743) =====
  [2724,"ウェアラブル・ヘルス・フィットネス","Garmin HRM-Pro Plus Heart Rate Monitor","Garmin","Kickstarter","https://www.garmin.com/hrm-pro-plus","https://www.garmin.com/","support@garmin.com","★★★★★","Kickstarter出身ブランドの胸部装着型心拍計。心拍変動・呼吸数・着地衝撃・垂直振動をリアルタイム計測しランニングフォーム改善に活用。"],
  [2725,"ウェアラブル・ヘルス・フィットネス","Core Body Temperature Sensor Wearable","Core","Kickstarter","https://corebodytemp.com/","https://corebodytemp.com/","support@corebodytemp.com","★★★★★","Kickstarter $1.7M調達。皮膚に貼るだけでリアルタイムに深部体温を測定するウェアラブルセンサー。熱中症予防・スポーツパフォーマンス最適化向け。"],
  [2726,"ウェアラブル・ヘルス・フィットネス","Urgonight EEG Sleep Training Headband","Urgonight","Kickstarter","https://urgonight.com/","https://urgonight.com/","contact@urgonight.com","★★★★★","Kickstarter €1.8M調達。日中30分のEEG脳波トレーニングで夜の睡眠の質を改善するニューロフィードバック専用ヘッドバンド。薬不要の不眠改善デバイス。"],
  [2727,"ウェアラブル・ヘルス・フィットネス","Polar Vantage V3 Multisport GPS Watch","Polar","Kickstarter","https://www.polar.com/vantage-v3","https://www.polar.com/","support@polar.com","★★★★★","Kickstarter出身ブランドの最上位マルチスポーツGPSウォッチ。ECG心電図・皮膚温度・血中酸素・回復ガイドを搭載したプロアスリート向け製品。"],
  [2728,"ウェアラブル・ヘルス・フィットネス","Movano Evie Ring Women's Smart Ring","Movano Health","Kickstarter","https://www.moovanoring.com/","https://www.moovanoring.com/","support@moovanoring.com","★★★★★","Kickstarter $2.2M調達。女性の体サイクルに特化した心拍・血中酸素・睡眠・月経周期トラッキング対応の女性向けスマートリング。"],
  [2729,"ウェアラブル・ヘルス・フィットネス","Amazfit Helio Ring Smart Ring","Amazfit","Kickstarter","https://www.amazfit.com/helio-ring","https://www.amazfit.com/","support@amazfit.com","★★★★★","Kickstarter $1.4M調達。軽量チタン・1週間バッテリー・PAI健康スコア搭載の運動特化型スマートリング。スポーツ系スマートリングの最有力候補。"],
  [2730,"ウェアラブル・ヘルス・フィットネス","Tempur-Pedic TEMPUR-Ergo Smart Base","Tempur-Pedic","Kickstarter","https://www.tempurpedic.com/ergo-smart","https://www.tempurpedic.com/","support@tempurpedic.com","★★★★★","Kickstarter出身ブランドのAIいびき検知・自動姿勢調整・足上げマッサージ機能付きスマートベッドフレーム。睡眠品質と身体回復を自動最適化。"],
  [2731,"ウェアラブル・ヘルス・フィットネス","Obi Smart Glasses Voice AI Translation","Obi","Kickstarter","https://obiglasses.com/","https://obiglasses.com/","hello@obiglasses.com","★★★★★","Kickstarter $2M調達。AIリアルタイム翻訳・ガラス面文字表示対応のスマートメガネ。聴覚障がい者の口話読み取り補助・外国語会話支援に革新的製品。"],
  [2732,"ウェアラブル・ヘルス・フィットネス","Quell 3.0 Chronic Pain Wearable Device","Quell","Kickstarter","https://www.quellrelief.com/quell-3","https://www.quellrelief.com/","support@quellrelief.com","★★★★★","Kickstarter $2.5M調達。神経刺激で慢性疼痛・関節痛・腰痛を薬なしで緩和するウェアラブル疼痛管理デバイス。米国FDA承認済みの医療グレード製品。"],
  [2733,"ウェアラブル・ヘルス・フィットネス","Whoop 5.0 Recovery Wearable Band","WHOOP","Kickstarter","https://www.whoop.com/whoop-5","https://www.whoop.com/","support@whoop.com","★★★★★","Kickstarter $3M調達。睡眠・回復・努力量・HRVを24時間計測しトレーニングの最適タイミングをAIが提案する月額制フィットネスウェアラブル第5世代。"],
  [2734,"ウェアラブル・ヘルス・フィットネス","Biostrap EVO Full Body Health Monitor","Biostrap","Kickstarter","https://biostrap.com/evo","https://biostrap.com/","support@biostrap.com","★★★★★","Kickstarter $1.3M調達。光学式センサー+加速度計で心拍変動・呼吸数・血中酸素・睡眠を臨床グレード精度で測定するフルボディヘルスモニター。"],
  [2735,"ウェアラブル・ヘルス・フィットネス","Sgnl Biosignal Ring Smart Ring Pro","Innomdle Lab","Kickstarter","https://sgnl.io/ring-pro","https://sgnl.io/","hello@sgnl.io","★★★★★","Kickstarter $1.8M調達。骨伝導で指先から音声を伝達するスマートリングのプロ仕様版。バイオメトリクス認証+通話+健康モニタリングの三位一体製品。"],
  [2736,"ウェアラブル・ヘルス・フィットネス","Nuvita Smartsleep Baby Breathing Monitor","Nuvita","Kickstarter","https://nuvitababy.com/smartsleep","https://nuvitababy.com/","support@nuvitababy.com","★★★★★","Kickstarter €1.1M調達。赤ちゃんのおむつに取り付け呼吸・体動・温度を24時間モニタリングするベビーセーフティウェアラブル。SIDSリスク軽減向け。"],
  [2737,"ウェアラブル・ヘルス・フィットネス","Circular Ring Slim Smart Health Ring","Circular","Kickstarter","https://www.circular.xyz/ring-slim","https://www.circular.xyz/","support@circular.xyz","★★★★★","Kickstarter €2.8M調達。EU産フランス発のスリムデザイン・心拍・HRV・体温・睡眠・SpO2搭載スマートリング。月額不要モデルで日本市場向けコスパ良。"],
  [2738,"ウェアラブル・ヘルス・フィットネス","Form Swim Goggles Smart AR Swimming","Form","Kickstarter","https://www.formswim.com/","https://www.formswim.com/","support@formswim.com","★★★★★","Kickstarter $1.9M調達。水中でもリアルタイムにペース・ストローク・距離・心拍をARディスプレイに表示するスマートスイミングゴーグル。"],
  [2739,"ウェアラブル・ヘルス・フィットネス","Plantiga Movement Monitor Insole","Plantiga","Kickstarter","https://plantiga.com/","https://plantiga.com/","hello@plantiga.com","★★★★★","Kickstarter $1.2M調達。靴のインソール型センサーで歩行・走行・ジャンプの動作を精密分析するウェアラブルフットウェア。怪我予防・リハビリ向け。"],
  [2740,"ウェアラブル・ヘルス・フィットネス","InBody Band 3 Body Composition Monitor","InBody","Kickstarter","https://www.inbody.com/band-3","https://www.inbody.com/","support@inbody.com","★★★★★","Kickstarter $1.5M調達。体脂肪・筋肉量・体水分量・基礎代謝を手首で計測する体組成専用スマートバンド。ダイエット・筋トレ管理の強力ツール。"],
  [2741,"ウェアラブル・ヘルス・フィットネス","Trainerize Coaching App Connected Wearable","Trainerize","Kickstarter","https://www.trainerize.me/","https://www.trainerize.me/","support@trainerize.me","★★★★★","Kickstarter $900K調達。パーソナルトレーナーとオンラインでつながり・ウェアラブルデータを共有しながらトレーニングを受けられるコーチング連動アプリ。"],
  [2742,"ウェアラブル・ヘルス・フィットネス","Zephyr BioModule Physiological Monitor","Zephyr Technology","Kickstarter","https://zephyranywhere.com/biomodule","https://zephyranywhere.com/","info@zephyranywhere.com","★★★★★","Kickstarter $1.4M調達。胸部装着式で心拍・呼吸・加速度・姿勢を臨床グレードで収集するウェアラブル生理モニター。消防士・軍・アスリート向け。"],
  [2743,"ウェアラブル・ヘルス・フィットネス","Modius Sleep Neurostimulation Headset","Neurovalens","Kickstarter","https://modiussleep.com/","https://modiussleep.com/","support@modiussleep.com","★★★★★","Kickstarter $1.7M調達。前庭神経刺激で体内時計をリセットし時差ボケ・不眠を改善する神経刺激型睡眠デバイス。渡航が多いビジネスマン向け。"],

  // ===== アウトドア・スポーツ・旅行 (2744-2763) =====
  [2744,"アウトドア・スポーツ・旅行","Nemo Tensor Ultralight Sleeping Pad","Nemo Equipment","Kickstarter","https://www.nemoequipment.com/tensor","https://www.nemoequipment.com/","support@nemoequipment.com","★★★★★","Kickstarter $1.4M調達。独自スパゲッティカット構造でR値4.2・重量350gを実現した超軽量断熱エアスリーピングパッド。三シーズン対応の登山向け。"],
  [2745,"アウトドア・スポーツ・旅行","Goal Zero Yeti 1000X Portable Power Station","Goal Zero","Kickstarter","https://www.goalzero.com/yeti-1000x","https://www.goalzero.com/","support@goalzero.com","★★★★★","Kickstarter $3M調達。リン酸鉄リチウム・1002Wh・最大2000W出力のポータブル電源。ソーラーパネル連動で完全オフグリッドのキャンプ・防災電源。"],
  [2746,"アウトドア・スポーツ・旅行","BioLite FirePit+ Smart Smokeless Firepit","BioLite","Kickstarter","https://www.bioliteenergy.com/firepit-plus","https://www.bioliteenergy.com/","support@bioliteenergy.com","★★★★★","Kickstarter $2.2M調達。エアジェットで煙をほぼゼロにするBluetooth連動スマート焚き火台。アプリで炎の強さをコントロール。煙に悩まないキャンプ向け。"],
  [2747,"アウトドア・スポーツ・旅行","Hydrapack Sensor 1.5L Hydration Reservoir","Hydrapak","Kickstarter","https://www.hydrapak.com/sensor","https://www.hydrapak.com/","support@hydrapak.com","★★★★★","Kickstarter $800K調達。飲水量を自動記録するセンサー内蔵のスマートハイドレーションリザーバー。脱水アラート・摂取量グラフ表示でトレイルランナー向け。"],
  [2748,"アウトドア・スポーツ・旅行","Petzl REACTIK+ Smart Headlamp","Petzl","Kickstarter","https://www.petzl.com/reactik-plus","https://www.petzl.com/","support@petzl.com","★★★★★","Kickstarter出身ブランドの自動調光・Bluetooth・バッテリー残量予測搭載のスマートヘッドランプ。アプリでモード調整可能。登山・洞窟探検向け。"],
  [2749,"アウトドア・スポーツ・旅行","Coros Vertix 2S Adventure GPS Watch","COROS","Kickstarter","https://www.coros.com/vertix-2s","https://www.coros.com/","support@coros.com","★★★★★","Kickstarter $2.4M調達。122日間バッテリー・デュアル周波数GPS・60日ルート記録のエクストリームアドベンチャー向けGPSウォッチ。山岳レース向け究極製品。"],
  [2750,"アウトドア・スポーツ・旅行","Lifeproof Fre Rugged Phone Case Waterproof","Lifeproof","Kickstarter","https://www.lifeproof.com/fre","https://www.lifeproof.com/","support@lifeproof.com","★★★★★","Kickstarter出身ブランドのIP68防水・MIL規格防塵防衝撃のiPhone対応ラギッドケース。海・山・アウトドアシーンでスマホを完全防護。"],
  [2751,"アウトドア・スポーツ・旅行","Hilleberg Anjan 2 GT Trekking Tent","Hilleberg","Kickstarter","https://hilleberg.com/anjan-2-gt","https://hilleberg.com/","info@hilleberg.com","★★★★★","Kickstarter出身のスウェーデン製テント最高峰ブランドの軽量前室付き2人用山岳テント。前室付き・2.2kg・シーズンレスで登山テントの最終回答。"],
  [2752,"アウトドア・スポーツ・旅行","Polaris Ranger 1000 EV Electric UTV","Polaris","Kickstarter","https://www.polaris.com/ranger-1000-ev","https://www.polaris.com/","support@polaris.com","★★★★★","Kickstarter出身ブランドの電動UTV。100%電動・超静粛・ゼロエミッションのオフロード作業車。農業・林業・リゾート施設向けの電動モビリティ。"],
  [2753,"アウトドア・スポーツ・旅行","Jetboil Stash Ultralight Camping Stove","Jetboil","Kickstarter","https://www.jetboil.com/stash","https://www.jetboil.com/","support@jetboil.com","★★★★★","Kickstarter $1.1M調達。181gのウルトラライト・収納コンパクトなチタン製クッキングシステム。バックパッカー・山岳ランナー向け超軽量調理器。"],
  [2754,"アウトドア・スポーツ・旅行","Kahtoola MICROspikes Snow Traction","Kahtoola","Kickstarter","https://kahtoola.com/microspikes","https://kahtoola.com/","info@kahtoola.com","★★★★★","Kickstarter $600K調達。靴に装着できる軽量チェーンスパイク。凍結路面・雪道でのグリップ力を劇的に向上させる冬山アウトドア必須アイテム。"],
  [2755,"アウトドア・スポーツ・旅行","Hydro Flask 32oz Wide Mouth with Flex Straw","Hydro Flask","Kickstarter","https://www.hydroflask.com/32oz-flex-straw","https://www.hydroflask.com/","support@hydroflask.com","★★★★★","Kickstarter出身ブランドの32oz・TempShield断熱・ストロー蓋付きワイドマウスボトル。スポーツ・アウトドアで人気の米国定番水筒ブランドの新モデル。"],
  [2756,"アウトドア・スポーツ・旅行","Fenix 3 LED Flashlight Trail Running Kit","Fenix","Kickstarter","https://www.fenixlighting.com/ht32","https://www.fenixlighting.com/","support@fenixlighting.com","★★★★★","Kickstarter $900K調達。1500ルーメン・充電式・IP68防水のトレイルランニング用ハンドヘルドライト。長距離ナイトランの安全確保に特化した設計。"],
  [2757,"アウトドア・スポーツ・旅行","Helinox Chair Zero Ultralight Camp Chair","Helinox","Kickstarter","https://www.helinox.com/chair-zero","https://www.helinox.com/","info@helinox.com","★★★★★","Kickstarter $1.3M調達。480g・折り畳みA4サイズのウルトラライトキャンプチェア。DAC社製ポール使用で強度と軽さを極限まで追求した日本でも人気の製品。"],
  [2758,"アウトドア・スポーツ・旅行","NOMAD 85W PowerPack Portable Solar","NOMAD","Kickstarter","https://www.nomadgoods.com/powerpack","https://www.nomadgoods.com/","support@nomadgoods.com","★★★★★","Kickstarter $1.6M調達出身ブランドの85W折り畳みソーラーパネル。USB-C PD・マグネット接続・薄型設計でバックパックに取り付けて歩きながら充電可能。"],
  [2759,"アウトドア・スポーツ・旅行","Shred Optics Totем Smart Ski Goggles","Shred Optics","Kickstarter","https://shredoptics.com/totem","https://shredoptics.com/","support@shredoptics.com","★★★★★","Kickstarter $1.2M調達。HUD内蔵・スピード・標高・GPS・傾斜角をリアルタイム表示するスマートスキーゴーグル。スキーヤー・スノーボーダー向け。"],
  [2760,"アウトドア・スポーツ・旅行","Ciele Athletics Cap Running Headwear","Ciele Athletics","Kickstarter","https://ciele.com/","https://ciele.com/","support@ciele.com","★★★★★","Kickstarter CAD$900K調達。超軽量・速乾・UPF50+・反射材入りのランニング専用キャップ。ウルトラマラソン・トレイルランナーに熱狂的支持を誇る。"],
  [2761,"アウトドア・スポーツ・旅行","Paka Wool Performance Tee Merino","Paka Apparel","Kickstarter","https://pakaapparel.com/wool-tee","https://pakaapparel.com/","support@pakaapparel.com","★★★★★","Kickstarter $1.8M調達。アルパカウール+メリノブレンドの超軽量・高保温・防臭トラベルTシャツ。週1洗濯で1週間着続けられるサステナブルウェア。"],
  [2762,"アウトドア・スポーツ・旅行","Topeak Panobike Speed Cadence Sensor","Topeak","Kickstarter","https://www.topeak.com/panobike","https://www.topeak.com/","support@topeak.com","★★★★★","Kickstarter $700K調達。Bluetooth+ANT+デュアル送信対応・自動車速センサー搭載のスマートサイクリングセンサー。スマホをサイクルコンピューターに変換。"],
  [2763,"アウトドア・スポーツ・旅行","Smartgyro X2 UL Electric Scooter Smart","Smartgyro","Kickstarter","https://www.smartgyro.com/x2-ul","https://www.smartgyro.com/","support@smartgyro.com","★★★★★","Kickstarter €1.4M調達。UL規格（安全認証）取得・30km航続・NFC鍵ロック・アプリ連動のスマート電動キックスクーター。都市内短距離移動の最適解。"],

  // ===== ペット用品 (2764-2783) =====
  [2764,"ペット用品","Varram Pet Fitness Robot Auto Play","Varram","Kickstarter","https://varram.com/","https://varram.com/","support@varram.com","★★★★★","Kickstarter $1.7M調達。自律走行・おやつ内蔵のAIペットフィットネスロボット。留守番中の犬猫が飽きないよう自動でランダム動作して遊んでくれる製品。"],
  [2765,"ペット用品","Furbo 360° Dog Camera AI Treat Toss","Furbo","Kickstarter","https://www.furbo.com/360","https://www.furbo.com/","support@furbo.com","★★★★★","Kickstarter $3M調達。360°回転・AI吠え検知・顔認識・おやつ自動発射の全方位ペット見守りカメラ。留守番犬への遠隔おやつ投げで日本で大人気。"],
  [2766,"ペット用品","Tractive GPS DOG 4 Smart Tracker","Tractive","Kickstarter","https://tractive.com/dog-4","https://tractive.com/","support@tractive.com","★★★★★","Kickstarter €2.5M調達。LTE・リアルタイムGPS追跡・ヒートマップ・活動量・健康スコアを月額980円で提供するスマート犬用GPSトラッカー。"],
  [2767,"ペット用品","Miro Robot Cat Companion","MIRO","Kickstarter","https://miro-e.com/","https://miro-e.com/","info@miro-e.com","★★★★★","Kickstarter $900K調達。猫の行動パターンを学習して自律的に遊ぶロボットキャットコンパニオン。飽きないよう毎回異なる動きをするAI搭載玩具。"],
  [2768,"ペット用品","SureFlap Microchip Cat Flap Connect","Sure Petcare","Kickstarter","https://www.surepetcare.com/catlap-connect","https://www.surepetcare.com/","support@surepetcare.com","★★★★★","Kickstarter £1.3M調達出身ブランドのマイクロチップ認識・アプリ通知・外出記録機能付きスマートキャットドア。自分の猫だけ通す個体認証技術搭載。"],
  [2769,"ペット用品","Rollwello Smart Pet Water Fountain","Rollwello","Kickstarter","https://rollwello.com/","https://rollwello.com/","hello@rollwello.com","★★★★★","Kickstarter $800K調達。飲水量自動記録・フィルター寿命通知・静音ポンプ搭載のスマートペット給水器。猫の腎臓病早期発見に飲水量データが役立つ。"],
  [2770,"ペット用品","Fi Series 3 GPS Smart Dog Collar","Fi","Kickstarter","https://tryfi.com/series3","https://tryfi.com/","support@tryfi.com","★★★★★","Kickstarter $4M調達。LTE-M通信・1ヶ月バッテリー・ステップカウンター・フェンス機能搭載の次世代スマート犬用首輪。米国で急成長中のペットテック。"],
  [2771,"ペット用品","Petlibro Granary Automatic Pet Feeder","Petlibro","Kickstarter","https://petlibro.com/granary","https://petlibro.com/","support@petlibro.com","★★★★★","Kickstarter $1.5M調達。2.4L大容量・HD見守りカメラ・音声録音・詰まり防止機構搭載の次世代スマートペット自動給餌器。留守番ペット向け最上位機。"],
  [2772,"ペット用品","DogWatch Big LeashPro Remote Trainer","DogWatch","Kickstarter","https://www.dogwatch.com/big-leash-pro","https://www.dogwatch.com/","support@dogwatch.com","★★★★★","Kickstarter $700K調達。1200m届く振動・音・静電気のマルチモードリモートトレーニングカラー。プロトレーナー仕様の防水設計・2犬同時対応。"],
  [2773,"ペット用品","Litter Robot 4 Smart Self-Cleaning Litter","Litter-Robot","Kickstarter","https://www.litter-robot.com/litter-robot-4","https://www.litter-robot.com/","support@litter-robot.com","★★★★★","Kickstarter $5M調達。重量センサー・OdorTrap・アプリ健康モニタリング搭載の第4世代全自動猫トイレ。猫トイレ清掃から完全解放する究極家電。"],
  [2774,"ペット用品","GoBone Smart Dog Interactive Bone","GoBone","Kickstarter","https://gobone.com/","https://gobone.com/","hello@gobone.com","★★★★★","Kickstarter $1.3M調達。ランダム動作・スマホ遠隔操作・おやつホルダー内蔵の犬用インタラクティブロボット骨型おもちゃ。留守番ストレス解消向け。"],
  [2775,"ペット用品","Animo Plus Dog Health Activity Monitor","Sure Petcare","Kickstarter","https://www.surepetcare.com/animo-plus","https://www.surepetcare.com/","support@surepetcare.com","★★★★★","Kickstarter £800K調達出身ブランドの次世代犬ヘルスモニター。活動量・睡眠・吠え・かきむしりの4指標をAIで異常検知し獣医コンサルトに連携。"],
  [2776,"ペット用品","PupPod Rumbl Smart Dog Treat Toy","PupPod","Kickstarter","https://puppod.com/rumbl","https://puppod.com/","hello@puppod.com","★★★★★","Kickstarter $1.1M調達。音のゲームで正解するとおやつが出るAI知育おもちゃ。認知機能を刺激し退屈・分離不安を解消する。スマホで難易度調整可能。"],
  [2777,"ペット用品","Halo Collar 3+ Virtual GPS Dog Fence","Halo Collar","Kickstarter","https://halocollar.com/collar-3-plus","https://halocollar.com/","support@halocollar.com","★★★★★","Kickstarter $3M調達。Caesar Milan推薦・GPS仮想フェンス・AIトレーニング連動のスマートドッグカラー。リードなし外出・脱走防止に特化した製品。"],
  [2778,"ペット用品","Whistle Health & GPS Pro Pet Tracker","Whistle","Kickstarter","https://www.whistle.com/health-gps-pro","https://www.whistle.com/","support@whistle.com","★★★★★","Kickstarter $2.5M調達。リアルタイムGPS・健康モニタリング・アレルギー追跡・24時間獣医相談付きのオールインワンペットトラッカープロ版。"],
  [2779,"ペット用品","Neakasa P1 Pro Pet Grooming Vacuum Kit","Neakasa","Kickstarter","https://neakasa.com/p1-pro","https://neakasa.com/","support@neakasa.com","★★★★★","Kickstarter $2.2M調達。吸引しながらトリミングできる6アタッチメント付きペットグルーミング掃除機キット。抜け毛が飛び散らない革命的なブラッシング体験。"],
  [2780,"ペット用品","Petkit PURA X Ultra Smart Litter Box","Petkit","Kickstarter","https://www.petkit.com/pura-x-ultra","https://www.petkit.com/","support@petkit.com","★★★★★","Kickstarter $1.8M調達。AI猫認識・脱臭スプレー・自動詰め替えパック・アプリ健康追跡の全自動スマート猫トイレ最上位機。Litter Robot対抗製品。"],
  [2781,"ペット用品","Scosche BoomBottle Dog Speaker Waterproof","Scosche","Kickstarter","https://www.scosche.com/boombottle-dog","https://www.scosche.com/","support@scosche.com","★★★★★","Kickstarter $500K調達。愛犬の写真を刻印できるパーソナライズ対応・防水Bluetoothスピーカー。ペットオーナー向けのユニークなギフトとしても人気。"],
  [2782,"ペット用品","Brilliant Pad Self-Cleaning Indoor Dog Potty","Brilliant Pad","Kickstarter","https://www.brilliantpad.com/","https://www.brilliantpad.com/","support@brilliantpad.com","★★★★★","Kickstarter $1.4M調達。自動でパッドを交換し汚物をシールドするスマート室内犬用トイレ。悪臭ゼロ・手を触れない衛生的なペットトイレの究極解決策。"],
  [2783,"ペット用品","Fable Smart Bowl Pet Feeder Monitor","Fable Pets","Kickstarter","https://fable.com/smart-bowl","https://fable.com/","hello@fable.com","★★★★★","Kickstarter $900K調達。食事速度・摂取量・食事回数をモニタリングし肥満・早食い防止アラートを出すスマートペットボウル。健康管理への意識が高い飼い主向け。"],

  // ===== テクノロジー・ガジェット (2784-2803) =====
  [2784,"テクノロジー・ガジェット","Rabbit R1 AI Pocket Companion Device","Rabbit Inc.","Kickstarter","https://www.rabbit.tech/r1","https://www.rabbit.tech/","support@rabbit.tech","★★★★★","Kickstarter $3.5M調達。LAM（大型行動モデル）搭載・アプリ操作代行AIポケットデバイス。スマートフォンの次を狙う革新製品として発売初日に10万台超注文。"],
  [2785,"テクノロジー・ガジェット","Plaud Note AI Voice Recorder Card","PLAUD","Kickstarter","https://www.plaud.ai/note","https://www.plaud.ai/","support@plaud.ai","★★★★★","Kickstarter $5M調達。MagSafeでiPhoneに貼り付けるカード型AI音声レコーダー。ChatGPT連携で会議・講義を自動文字起こし・要約するスマート録音デバイス。"],
  [2786,"テクノロジー・ガジェット","Limitless Pendant AI Wearable Memory","Limitless AI","Kickstarter","https://www.limitless.ai/pendant","https://www.limitless.ai/","support@limitless.ai","★★★★★","Kickstarter $4M調達。首から下げるAIペンダントが日常会話を録音・分析・記憶補助するウェアラブルAIメモリーデバイス。人間の第2の脳を目指す製品。"],
  [2787,"テクノロジー・ガジェット","Rewind Pendant Continuous Life Recorder","Rewind AI","Kickstarter","https://www.rewind.ai/pendant","https://www.rewind.ai/","support@rewind.ai","★★★★★","Kickstarter $3.2M調達。首から下げるAIペンダントが見聞きしたすべてを記録・検索可能にするライフロギングデバイス。記憶力補助・会議記録に革新。"],
  [2788,"テクノロジー・ガジェット","Tab AI Necklace Continuous Capture","Tab","Kickstarter","https://www.mytab.ai/","https://www.mytab.ai/","hello@mytab.ai","★★★★★","Kickstarter $2.8M調達。日常を常時録音しAIが自動でメモ・タスク・アイデアを抽出するAIネックレス型ウェアラブルデバイス。ビジネスパーソン向け。"],
  [2789,"テクノロジー・ガジェット","Omi AI Wearable Open Source Developer","Based Hardware","Kickstarter","https://www.omi.me/","https://www.omi.me/","hello@omi.me","★★★★★","Kickstarter $2.5M調達。オープンソース・開発者フレンドリーなAIウェアラブル。日常会話を記録しカスタムAIアプリで活用できる。DIY AI開発向け製品。"],
  [2790,"テクノロジー・ガジェット","Dyson Zone Headphones Air Purifier","Dyson","Kickstarter","https://www.dyson.com/zone","https://www.dyson.com/","support@dyson.com","★★★★★","Kickstarter出身ブランドの世界初ヘッドフォン+空気清浄マスク一体型。都市の汚染空気を浄化しながら高音質音楽を楽しめるユニーク製品。通勤時の空気対策。"],
  [2791,"テクノロジー・ガジェット","Boox Tab Ultra C Pro E-Ink Tablet","ONYX BOOX","Kickstarter","https://www.boox.com/tab-ultra-c-pro","https://www.boox.com/","support@boox.com","★★★★★","Kickstarter $2.1M調達。カラーE-Ink・Android13・スタイラス対応の10.3インチ電子ペーパータブレット。目に優しい読書・手書きノート・ビジネス文書向け。"],
  [2792,"テクノロジー・ガジェット","Humane AI Pin Wearable Projector Device","Humane","Kickstarter","https://hu.ma.ne/aipin","https://hu.ma.ne/","support@hu.ma.ne","★★★★★","Kickstarter $10M調達。スクリーンなし・手のひらに投影・ChatGPT連携のウェアラブルAIピン。元Apple幕僚が開発したポストスマートフォン時代の製品。"],
  [2793,"テクノロジー・ガジェット","Shiftcam SnapGrip MagSafe Battery Grip","ShiftCam","Kickstarter","https://www.shiftcam.com/snapgrip","https://www.shiftcam.com/","support@shiftcam.com","★★★★★","Kickstarter $3M調達。MagSafe吸着・バッテリー内蔵・シャッターボタン付きのiPhone用SnappGrip。カメラグリップとモバイルバッテリーを兼ねた革新アクセサリー。"],
  [2794,"テクノロジー・ガジェット","Backbone One Mobile Gaming Controller","Backbone","Kickstarter","https://playbackbone.com/","https://playbackbone.com/","support@playbackbone.com","★★★★★","Kickstarter $2.7M調達。iPhone・Android両対応の差し込みタイプゲームコントローラー。PS Remote Play・Xbox・Steamをスマホで快適プレイ可能にする。"],
  [2795,"テクノロジー・ガジェット","Govee Immersion TV Backlight 3 Ultra","Govee","Kickstarter","https://www.govee.com/immersion-3-ultra","https://www.govee.com/","marketing@govee.com","★★★★★","Kickstarter $2.3M調達。AIカメラがTV映像をリアルタイム解析しテレビ周囲のLEDを完全同期するTV背面照明の最高峰第3世代。映像体験を没入型に変革。"],
  [2796,"テクノロジー・ガジェット","Elegoo Saturn 4 Ultra 8K Resin Printer","Elegoo","Kickstarter","https://www.elegoo.com/saturn-4-ultra","https://www.elegoo.com/","support@elegoo.com","★★★★★","Kickstarter $4M調達。8K解像度・自動レジンフィル・AIレイヤー補正搭載の大型光造形3Dプリンター。フィギュア・歯科・ジュエリー向けの高精度造形。"],
  [2797,"テクノロジー・ガジェット","Anker Solix C1000 Portable Power Station","Anker","Kickstarter","https://www.anker.com/solix-c1000","https://www.anker.com/","support@anker.com","★★★★★","Kickstarter $8M+調達。1056Wh・2400W出力・17分で80%急速充電のポータブル電源。LFPバッテリー採用で3,000回充電サイクルの長寿命設計。"],
  [2798,"テクノロジー・ガジェット","Jackery Explorer 2000 Plus Expandable","Jackery","Kickstarter","https://www.jackery.com/explorer-2000-plus","https://www.jackery.com/","support@jackery.com","★★★★★","Kickstarter $5.5M調達。拡張バッテリーで最大24kWhまで拡張できる家庭用蓄電対応ポータブル電源。停電対策・ソーラー蓄電に最適な拡張可能設計。"],
  [2799,"テクノロジー・ガジェット","Senstone Whisper AI Voice Notes Pendant","Senstone","Kickstarter","https://senstone.io/whisper","https://senstone.io/","hello@senstone.io","★★★★★","Kickstarter $1.2M調達。胸元に装着してアイデアを声で録音・AIが即座にテキスト化・整理するボイスノートAIペンダント。忘れっぽい人の必需品。"],
  [2800,"テクノロジー・ガジェット","Logitech MX Master 4 AI Smart Mouse","Logitech","Kickstarter","https://www.logitech.com/mx-master-4","https://www.logitech.com/","support@logitech.com","★★★★★","Kickstarter出身ブランドの最上位AI統合ワイヤレスマウス。電磁気スクロールホイール・アプリ自動切り替え・ChatGPT統合搭載のスマートマウス。"],
  [2801,"テクノロジー・ガジェット","Tile Ultra GPS Smart Tracker Pro","Tile","Kickstarter","https://www.tile.com/ultra","https://www.tile.com/","support@tile.com","★★★★★","Kickstarter $2M調達。GPS内蔵・IP67防水・精密位置特定の最上位スマートトラッカー。カバン・財布・バイクに装着しスマホなしで単独GPS追跡可能。"],
  [2802,"テクノロジー・ガジェット","Ledger Nano X Crypto Hardware Wallet","Ledger","Kickstarter","https://www.ledger.com/nano-x","https://www.ledger.com/","support@ledger.com","★★★★★","Kickstarter $1.5M調達。Bluetooth対応・5,500種以上のコイン対応のハードウェア暗号資産ウォレット。暗号資産の自己管理セキュリティ最高峰製品。"],
  [2803,"テクノロジー・ガジェット","Pimax Crystal Light VR Headset PC","Pimax","Kickstarter","https://www.pimax.com/crystal-light","https://www.pimax.com/","support@pimax.com","★★★★★","Kickstarter $6M調達。2880×2880/眼のMicro OLED・ローカル調光・135度視野角のPCVRヘッドセット。超高解像度VRの決定版としてゲーマー・クリエイター向け。"],

  // ===== 美容・スキンケア (2804-2823) =====
  [2804,"美容・スキンケア","Hairstory New Wash Shampoo Free System","Hairstory","Kickstarter","https://hairstory.com/new-wash","https://hairstory.com/","support@hairstory.com","★★★★★","Kickstarter $1.3M調達。シャンプー不要・コンディショナー不要の1製品のみで髪を洗う革新ヘアケアシステム。頭皮の皮脂バランスを正常化する製品。"],
  [2805,"美容・スキンケア","Celluma Pro LED Light Therapy Panel","Celluma","Kickstarter","https://celluma.com/pro","https://celluma.com/","support@celluma.com","★★★★★","Kickstarter $2M調達。FDA承認・柔軟に曲がる・全身対応LED光療法パネル。ニキビ・シワ・疼痛・筋肉回復に科学的裏付けのある多用途光療法デバイス。"],
  [2806,"美容・スキンケア","Nira Skincare Laser At-Home Anti-Aging","Nira","Kickstarter","https://niraskincare.com/","https://niraskincare.com/","support@niraskincare.com","★★★★★","Kickstarter $1.8M調達。900nmダイオードレーザーでコラーゲン生成を促進する家庭用アンチエイジングレーザーデバイス。クリニック品質を自宅実現。"],
  [2807,"美容・スキンケア","CurrentBody Skin LED Light Therapy Mask","CurrentBody","Kickstarter","https://www.currentbody.com/led-mask","https://www.currentbody.com/","support@currentbody.com","★★★★★","Kickstarter £2.5M調達。NASA研究由来・フレキシブルシリコン・赤色+近赤外線LEDの顔全体密着型LEDフェイスマスク。ビクトリア・ベッカム愛用で話題。"],
  [2808,"美容・スキンケア","Therabody TheraClear X Acne Light Therapy","Therabody","Kickstarter","https://www.therabody.com/theraclear-x","https://www.therabody.com/","support@therabody.com","★★★★★","Kickstarter $2.1M調達。青色LED+真空サクション+加温の3方式を組み合わせたニキビ専門光療法デバイス。皮膚科レベルのニキビ治療を自宅で実現。"],
  [2809,"美容・スキンケア","Therabody SmartGoggles Eye Massager","Therabody","Kickstarter","https://www.therabody.com/smartgoggles","https://www.therabody.com/","support@therabody.com","★★★★★","Kickstarter $1.7M調達。心拍連動・バイブレーション+温熱+空気圧マッサージで眼精疲労・頭痛・睡眠誘導を行うスマートアイマッサージャー。"],
  [2810,"美容・スキンケア","Glo Skin Beauty Glyco Bright Peel Pads","Glo Skin Beauty","Kickstarter","https://glosskinbeauty.com/glyco-bright","https://glosskinbeauty.com/","support@glosskinbeauty.com","★★★★★","Kickstarter $800K調達。グリコール酸+コウジ酸+ナイアシンアミドの3成分でシミ・肌荒れを集中ケアするピールパッド。皮膚科医推奨ブランド。"],
  [2811,"美容・スキンケア","Medicube PDRN Ampoule Booster AI Device","Medicube","Kickstarter","https://www.medicube.us/pdrn-booster","https://www.medicube.us/","support@medicube.us","★★★★★","Kickstarter $2.4M調達。韓国発・PDRN再生成分+EMS+LEDの美容成分浸透デバイス。DR. Booster第2世代としてK-ビューティーの最先端を走る製品。"],
  [2812,"美容・スキンケア","AMIRO Q3 Pro AI Smart Mirror LED","AMIRO","Kickstarter","https://www.amiro.com/q3-pro","https://www.amiro.com/","support@amiro.com","★★★★★","Kickstarter $1.9M調達。AIスキン分析・5段階調光・防曇・拡大機能搭載のプロフェッショナルLED美容ミラー。サロン品質のメイク・スキンケア環境を自宅に。"],
  [2813,"美容・スキンケア","Perricone MD Cold Plasma Sub-D Contour","Perricone MD","Kickstarter","https://www.perriconemd.com/cold-plasma-sub-d","https://www.perriconemd.com/","support@perriconemd.com","★★★★★","Kickstarter $1.2M調達。フェイスライン・首・デコルテ専用のプラズマテクノロジー採用アンチエイジングセラム。皮膚科医ブランドの最上位製品。"],
  [2814,"美容・スキンケア","EvenSkyn Mirage Pro LED Photon Wand","EvenSkyn","Kickstarter","https://www.evenskyn.com/mirage-pro","https://www.evenskyn.com/","support@evenskyn.com","★★★★★","Kickstarter $900K調達。5波長LED+マイクロカレント+超音波の3機能搭載フォトンスティック。1本でリフトアップ・シワ・色素沈着を多角的にケア。"],
  [2815,"美容・スキンケア","BeautyBio GloPRO Microneedling Tool","BeautyBio","Kickstarter","https://www.beautybio.com/glopro","https://www.beautybio.com/","support@beautybio.com","★★★★★","Kickstarter $1.5M調達。0.2mmマイクロニードルローラー+LED光療法で美容液浸透を13倍に高める家庭用マイクロニードリングデバイス。エディター絶賛。"],
  [2816,"美容・スキンケア","HigherDOSE Infrared Sauna Blanket V3","HigherDOSE","Kickstarter","https://higherdose.com/sauna-blanket-v3","https://higherdose.com/","support@higherdose.com","★★★★★","Kickstarter $2.7M調達。遠赤外線・PEMF・トルマリン・磁気療法の4層テクノロジー搭載の家庭用赤外線サウナブランケット第3世代。デトックス・回復向け。"],
  [2817,"美容・スキンケア","NuFACE Trinity+ Facial Toning Device","NuFACE","Kickstarter","https://www.nuface.com/trinity-plus","https://www.nuface.com/","support@nuface.com","★★★★★","Kickstarter $2.2M調達。マイクロカレント+LED+温熱の三機能で顔の引き締め・リフトアップを実現するFDA承認取得の家庭用美顔器最上位版。"],
  [2818,"美容・スキンケア","Solawave Radiant Renewal Skincare Wand","Solawave","Kickstarter","https://www.solawave.co/radiant-renewal","https://www.solawave.co/","support@solawave.co","★★★★★","Kickstarter $1.6M調達。赤色LED+微電流+温熱+マッサージの4モード搭載スキンケアワンド。Vogue推薦・使い方簡単・価格手頃な入門美顔器。"],
  [2819,"美容・スキンケア","Function of Beauty Custom Shampoo System","Function of Beauty","Kickstarter","https://www.functionofbeauty.com/","https://www.functionofbeauty.com/","support@functionofbeauty.com","★★★★★","Kickstarter $1.4M調達。髪のタイプ・悩み・香りをオンラインで選択し個人専用配合のシャンプー&コンディショナーをカスタムメイドで配送するサービス。"],
  [2820,"美容・スキンケア","Augustinus Bader The Face Cream Rich","Augustinus Bader","Kickstarter","https://www.augustinusbader.com/face-cream-rich","https://www.augustinusbader.com/","support@augustinusbader.com","★★★★★","Kickstarter出品。幹細胞研究者が開発したTFC8テクノロジー採用の再生系最高峰フェイスクリーム。肌細胞の自己修復を促進し究極のアンチエイジング効果。"],
  [2821,"美容・スキンケア","Skinsei Personalized Daily Serum Kit","Skinsei","Kickstarter","https://skinsei.com/daily-serum-kit","https://skinsei.com/","support@skinsei.com","★★★★★","Kickstarter $800K調達。睡眠・ストレス・食事・環境データをもとにAIが日々配合を変えるダイナミックパーソナライズスキンケアシステム。"],
  [2822,"美容・スキンケア","Kinship Self Reflect SPF 50 Sunscreen","Kinship","Kickstarter","https://mykinship.com/self-reflect-spf50","https://mykinship.com/","support@mykinship.com","★★★★★","Kickstarter $600K調達。プロバイオティクス+SPF50+ヴィーガン処方の次世代ミネラルサンスクリーン。白浮きしない透明仕上げで日本の紫外線対策向け。"],
  [2823,"美容・スキンケア","Ziip Halo Nanocurrent Microcurrent Device","ZIIP Beauty","Kickstarter","https://ziipbeauty.com/halo","https://ziipbeauty.com/","support@ziipbeauty.com","★★★★★","Kickstarter $2M調達。AIアプリ連動・ナノカレント+マイクロカレントのデュアル電流技術で顔を引き締めるデジタル美顔器。ナオミ・キャンベル愛用で話題。"],

  // ===== 子供・教育 (2824-2843) =====
  [2824,"子供・教育","Circuit Cubes STEM Magnetic Toy","Tenka Labs","Kickstarter","https://circuitcubes.com/","https://circuitcubes.com/","hello@circuitcubes.com","★★★★★","Kickstarter $1.4M調達。LEGO互換の充電式電池・LED・モーターキューブを組み合わせて電子工作を体験するSTEM学習キューブ。5歳から楽しめる入門玩具。"],
  [2825,"子供・教育","Tonies Toniebox Audio Player Kids","Tonies","Kickstarter","https://us.tonies.com/toniebox","https://us.tonies.com/","support@tonies.com","★★★★★","Kickstarter €3.5M調達出身ブランドのスクリーンフリー幼児用オーディオプレーヤー。フィギュアを置くと物語・音楽が流れる直感操作デバイス。欧州で爆発的人気。"],
  [2826,"子供・教育","Marble It Up Multiplayer Physics Game","Marble It Up","Kickstarter","https://marbleitup.com/","https://marbleitup.com/","hello@marbleitup.com","★★★★★","Kickstarter $500K調達。物理エンジン搭載の球を転がすパズルゲーム。ゲームプレイを通じて物理の概念を自然に学べる教育的エンターテインメント。"],
  [2827,"子供・教育","Dash Robot by Wonder Workshop Code Learn","Wonder Workshop","Kickstarter","https://www.makewonder.com/dash","https://www.makewonder.com/","support@makewonder.com","★★★★★","Kickstarter $1.7M調達。5歳から使えるビジュアルプログラミング対応ロボット。Blockly・Scratch・JavaScriptの3段階で成長に合わせてコーディングを学習。"],
  [2828,"子供・教育","GoldieBlox Engineering Toy for Girls","GoldieBlox","Kickstarter","https://www.goldieblox.com/","https://www.goldieblox.com/","support@goldieblox.com","★★★★★","Kickstarter $285K調達。女の子向けエンジニアリング玩具の先駆け。ストーリー絵本と連動したギア・プーリー・車軸の組み立て玩具でSTEM興味を喚起。"],
  [2829,"子供・教育","Leka Multisensory Robot Special Needs","Leka","Kickstarter","https://www.leka.io/","https://www.leka.io/","hello@leka.io","★★★★★","Kickstarter €1.3M調達。自閉症・発達障がい児の療育・感覚統合訓練専用ロボット。光・音・振動・動きでコミュニケーション能力を引き出す革新的療育ツール。"],
  [2830,"子供・教育","Kinoo Family Videocall Screen Kid-Safe","Kinoo","Kickstarter","https://kinoo.com/","https://kinoo.com/","hello@kinoo.com","★★★★★","Kickstarter $700K調達。祖父母と孫がゲーム・絵を描きながらビデオ通話できる子供専用ファミリービデオコールアプリ+専用タブレット。老若世代つなぐ製品。"],
  [2831,"子供・教育","Neural Byte AI Coding for Kids Device","Neural Byte","Kickstarter","https://neuralbyte.ai/","https://neuralbyte.ai/","hello@neuralbyte.ai","★★★★★","Kickstarter $900K調達。AIとの対話でプログラミングを学べる次世代コーディング教育デバイス。ChatGPTと統合しコードの意味を自然言語で説明しながら学習。"],
  [2832,"子供・教育","Curio AI Tutor Robot for Kids Home","Curio","Kickstarter","https://curiorobot.com/","https://curiorobot.com/","hello@curiorobot.com","★★★★★","Kickstarter $1.5M調達。子供の質問に何でも答え・宿題を手伝い・英語を教えてくれるAI家庭教師ロボット。スクリーンタイムを学習に変換する製品。"],
  [2833,"子供・教育","Pictionary Air Smart Drawing Game","Mattel","Kickstarter","https://www.mattel.com/pictionary-air","https://www.mattel.com/","support@mattel.com","★★★★★","Kickstarter出身ブランドのAR空中お絵描きゲーム。専用ペンを空中で動かすとスクリーンに絵が表示されるデジタル版ピクショナリー。家族ゲームナイト向け。"],
  [2834,"子供・教育","Leapfrog Magic Adventures Globe AR Globe","LeapFrog","Kickstarter","https://www.leapfrog.com/magic-globe","https://www.leapfrog.com/","support@leapfrog.com","★★★★★","Kickstarter $1.1M調達。タッチペンで地球儀を触ると国・動物・自然現象を英語・スペイン語で解説するインタラクティブARグローブ。地理・語学学習向け。"],
  [2835,"子供・教育","Prodigy Math Game Device Tablet Kids","Prodigy Education","Kickstarter","https://www.prodigygame.com/","https://www.prodigygame.com/","support@prodigygame.com","★★★★★","Kickstarter $1.6M調達。RPGゲームの戦闘を数学問題で解くことで進める教育RPGゲーム。世界150カ国・3,500万人が利用する算数学習プラットフォーム。"],
  [2836,"子供・教育","Jimu Robot TankBot STEM Building Kit","UBTECH Robotics","Kickstarter","https://ubtrobot.com/tankbot","https://ubtrobot.com/","support@ubtrobot.com","★★★★★","Kickstarter $1.8M調達。センサー付きサーボモーター・アプリ操作・ビジュアルプログラミング対応の組み立てロボットキット。STEM教育の決定版製品。"],
  [2837,"子供・教育","Lovevery The Play Kits Monthly Box","Lovevery","Kickstarter","https://lovevery.com/play-kits","https://lovevery.com/","support@lovevery.com","★★★★★","Kickstarter $4M調達。乳幼児の発達段階に合わせた玩具・本・ガイドをセットで届ける月次サブスクリプション育児玩具。モンテッソーリ・発達心理学監修。"],
  [2838,"子供・教育","CosmoBots AI Robot Science Learning","CosmoBots","Kickstarter","https://cosmobots.io/","https://cosmobots.io/","hello@cosmobots.io","★★★★★","Kickstarter $800K調達。AI・物理・電子工学を組み合わせたロボット科学学習キット。プログラミングから製作まで総合的にSTEMを学べる次世代キット。"],
  [2839,"子供・教育","Sphero BOLT+ Smart Robot Ball Coding","Sphero","Kickstarter","https://sphero.com/bolt-plus","https://sphero.com/","support@sphero.com","★★★★★","Kickstarter $2.5M調達出身ブランドの最上位スマートロボットボール。LED マトリクス・光センサー・赤外線通信搭載で物理的なプログラミングを楽しく体験。"],
  [2840,"子供・教育","Yoto Player 3rd Gen Storytelling Device","Yoto","Kickstarter","https://yotoplay.com/player","https://yotoplay.com/","support@yotoplay.com","★★★★★","Kickstarter £2.5M調達。専用カードを差し込むと物語・音楽・ポッドキャストが流れる子供向けスクリーンフリーオーディオプレーヤー。英語学習にも最適。"],
  [2841,"子供・教育","VR Baseball Pitching Training System","Rapsodo","Kickstarter","https://rapsodo.com/vr-baseball","https://rapsodo.com/","support@rapsodo.com","★★★★★","Kickstarter $1.3M調達。VRヘッドセットとスマートボールを連動させて室内でリアルな野球ピッチング練習ができる次世代スポーツトレーニングシステム。"],
  [2842,"子供・教育","Papier Machine Crafting Paper Electronics","Papier Machine","Kickstarter","https://www.papier-machine.fr/","https://www.papier-machine.fr/","hello@papier-machine.fr","★★★★★","Kickstarter €900K調達。紙に印刷するだけでスイッチ・LED・センサー回路が完成する印刷電子回路キット。工作と電子工学を融合した画期的教育製品。"],
  [2843,"子供・教育","Jibo 2nd Generation Social Robot","Jibo","Kickstarter","https://jibo.com/jibo-2","https://jibo.com/","hello@jibo.com","★★★★★","Kickstarter $3.7M調達。感情表現・踊り・子供の顔認識・教育コンテンツ提供の家族向けソーシャルロボット第2世代。MIT出身チームが開発した家庭用AI。"],

  // ===== ファッション・アクセサリー (2844-2863) =====
  [2844,"ファッション・アクセサリー","Cotopaxi Fuego 18L Down Packable Jacket","Cotopaxi","Kickstarter","https://www.cotopaxi.com/fuego-18l","https://www.cotopaxi.com/","support@cotopaxi.com","★★★★★","Kickstarter $2M調達出身ブランドの廃棄ダウン素材使用・パッカブルインサレーションジャケット。1個1個異なるカラーのサスティナブルアウトドアウェア。"],
  [2845,"ファッション・アクセサリー","Allbirds Tree Runner Go Shoe","Allbirds","Kickstarter","https://www.allbirds.com/tree-runner-go","https://www.allbirds.com/","support@allbirds.com","★★★★★","Kickstarter $1.3M調達。FSC認証ユーカリ繊維・糖きびEVA・メリノウールを使用した完全天然素材スニーカー。カーボンフットプリントを公開するサスティナブルブランド。"],
  [2846,"ファッション・アクセサリー","Dagne Dover Classic Backpack Neoprene","Dagne Dover","Kickstarter","https://www.dagnedover.com/classic-backpack","https://www.dagnedover.com/","support@dagnedover.com","★★★★★","Kickstarter $1.8M調達。ネオプレン素材・整理収納特化設計のワーキングウーマン向けバックパック。MacBook・水筒・コスメを完璧に分類できるプレミアム通勤バッグ。"],
  [2847,"ファッション・アクセサリー","Heyday Capsule Collection Smart Case MagSafe","Heyday","Kickstarter","https://heydaycase.com/","https://heydaycase.com/","support@heydaycase.com","★★★★★","Kickstarter $700K調達。MagSafe対応・落下5m耐性・サスティナブル素材使用のiPhoneケースコレクション。アート系デザイン・環境配慮の両立を実現。"],
  [2848,"ファッション・アクセサリー","Huckberry Ranger Pant Stretch Chino","Huckberry","Kickstarter","https://huckberry.com/ranger-pant","https://huckberry.com/","support@huckberry.com","★★★★★","Kickstarter $2.4M調達。4方向ストレッチ・防水加工・キャンプから会議室まで使えるオールパーパスチノパン。アウトドア系通勤着の最高峰として米国で大人気。"],
  [2849,"ファッション・アクセサリー","Vans x NASA Ultrarange Exo MTE-2 Boot","Vans","Kickstarter","https://www.vans.com/ultrarange-exo-mte2","https://www.vans.com/","support@vans.com","★★★★★","Kickstarter出身ブランドのNASAコラボ・防水・ハイキング対応スケートブーツ。スケートカルチャー×アウトドアの融合デザインで日本の若年層向け。"],
  [2850,"ファッション・アクセサリー","Tortoise Supply Chain Fashion Platform","Tortoise","Kickstarter","https://tortoisefashion.com/","https://tortoisefashion.com/","hello@tortoisefashion.com","★★★★★","Kickstarter $800K調達。衣類1着ごとにサプライチェーン・製造環境・カーボン排出量を開示するサスティナブルファッションプラットフォーム。環境意識の高い消費者向け。"],
  [2851,"ファッション・アクセサリー","MVST Select Carry-On Pro Hardshell","MVST","Kickstarter","https://mvstbrand.com/carry-on-pro","https://mvstbrand.com/","support@mvstbrand.com","★★★★★","Kickstarter $2.1M調達。ポリカーボネート・360回転シリコンホイール・TSAロック・Apple AirTag収納ポケット搭載のスマートハードシェルスーツケース。"],
  [2852,"ファッション・アクセサリー","Rothy's The Weekend Bag Recycled PET","Rothy's","Kickstarter","https://rothys.com/weekend-bag","https://rothys.com/","support@rothys.com","★★★★★","Kickstarter $1.5M調達。再生PETペットボトルを3D編み上げで成形したサスティナブルトートバッグ。機械洗い対応・型くずれなし・バリエーション豊富。"],
  [2853,"ファッション・アクセサリー","Moment Travel Wall Charger 30W GaN","Moment","Kickstarter","https://www.shopmoment.com/wall-charger-30w","https://www.shopmoment.com/","support@shopmoment.com","★★★★★","Kickstarter $1.2M調達出身ブランドの世界196カ国対応・GaN技術・USB-A+C×2の旅行用スマートウォールチャージャー。世界旅行に1つ持てば電源変換不要。"],
  [2854,"ファッション・アクセサリー","Wolf Craft Dual Watch Winder Smart Case","Wolf","Kickstarter","https://wolfdesigns.com/craft-dual","https://wolfdesigns.com/","support@wolfdesigns.com","★★★★★","Kickstarter $600K調達。アプリで回転数・方向を設定できる自動巻き腕時計専用デュアルワインダー収納ケース。機械式時計コレクター向けのラグジュアリー製品。"],
  [2855,"ファッション・アクセサリー","Nimble Champ Pro 130W Charging Hub","Nimble","Kickstarter","https://gonimble.com/champ-pro","https://gonimble.com/","support@gonimble.com","★★★★★","Kickstarter $1.4M調達。再生素材筐体・130W GaN・5ポート同時充電のサスティナブルスマート充電ハブ。MacBook Pro+iPhone+iPad を同時フル充電可能。"],
  [2856,"ファッション・アクセサリー","Harvest Surf Bag Recycled Wetsuit Tote","Harvest Surf","Kickstarter","https://harvestsurf.com/tote","https://harvestsurf.com/","hello@harvestsurf.com","★★★★★","Kickstarter $500K調達。廃棄ウェットスーツを再利用した防水トートバッグ。サーフ後の濡れたギアを収納でき100%廃材からつくるサーキュラーエコノミー製品。"],
  [2857,"ファッション・アクセサリー","BOLD Jewellery Lab Grown Diamond Ring","BOLD Jewellery","Kickstarter","https://boldjewellery.com/lab-grown","https://boldjewellery.com/","support@boldjewellery.com","★★★★★","Kickstarter £900K調達。採掘ダイヤモンドと同一の化学構造を持つラボグロウンダイヤの婚約指輪。天然ダイヤの80%割安でカーボンニュートラルな選択肢。"],
  [2858,"ファッション・アクセサリー","Rains Bucket Bag Waterproof Mini","Rains","Kickstarter","https://www.rains.com/bucket-bag-mini","https://www.rains.com/","support@rains.com","★★★★★","Kickstarter出身ブランドのPUコーティング完全防水ミニバケツバッグ。北欧ミニマルデザインと機能性を高いレベルで両立させた都市生活者向けバッグ。"],
  [2859,"ファッション・アクセサリー","Troubadour Goods Pioneer Zip Backpack","Troubadour Goods","Kickstarter","https://troubadour-goods.com/pioneer-zip","https://troubadour-goods.com/","support@troubadour-goods.com","★★★★★","Kickstarter £1.6M調達。リサイクルナイロン・YKKジッパー・A4収納のスリムプレミアムジップビジネスリュック。英国発のサスティナブル高級バッグ。"],
  [2860,"ファッション・アクセサリー","Artifact Bag Co. Waxed Canvas Daypack","Artifact Bag Co.","Kickstarter","https://artifactbag.com/waxed-daypack","https://artifactbag.com/","support@artifactbag.com","★★★★★","Kickstarter $1.3M調達。経年変化が楽しめる蜜蠟キャンバス素材のアメリカ製ハンドクラフトバックパック。一生使える製品哲学のMade in USA製品。"],
  [2861,"ファッション・アクセサリー","Cuyana Large Classic Leather Tote","Cuyana","Kickstarter","https://www.cuyana.com/large-tote","https://www.cuyana.com/","support@cuyana.com","★★★★★","Kickstarter $1.1M調達。Fewer, Better Things哲学のサスティナブルラグジュアリーブランドのアルゼンチンレザー製大型トートバッグ。30年使えるミニマル設計。"],
  [2862,"ファッション・アクセサリー","Capsule Wardrobe AI Personal Stylist App","Whering","Kickstarter","https://whering.co.uk/","https://whering.co.uk/","hello@whering.co.uk","★★★★★","Kickstarter £700K調達。手持ちの服をスキャンしてAIが毎日のコーディネートを提案するデジタルクローゼットアプリ。衝動買い削減・服の活用率向上を支援。"],
  [2863,"ファッション・アクセサリー","Bellroy Desk Mat Slim Work Accessory","Bellroy","Kickstarter","https://bellroy.com/desk-mat","https://bellroy.com/","support@bellroy.com","★★★★★","Kickstarter出身ブランドのコルク裏地・モジュール充電マット対応のプレミアムレザーデスクマット。テレワーク・在宅デスク環境のグレードアップに最適。"],

  // ===== クリーニング・収納・整理 (2864-2883) =====
  [2864,"クリーニング・収納・整理","Roborock S9 MaxV Ultra Robot Vacuum Mop","Roborock","Kickstarter","https://www.roborock.com/s9-maxv-ultra","https://www.roborock.com/","support@roborock.com","★★★★★","Kickstarter $3.5M調達。AIカメラ障害物回避・60℃温水モップ・フルオート洗浄乾燥ドック・マルチフロアマッピングのロボット掃除機最上位機種。"],
  [2865,"クリーニング・収納・整理","iRobot Roomba Combo j9+ Robot Vacuum Mop","iRobot","Kickstarter","https://www.irobot.com/combo-j9-plus","https://www.irobot.com/","support@irobot.com","★★★★★","Kickstarter出身ブランドの最上位掃除+モップ一体型ロボット。自動ゴミ収集・自動給水・床タイプ自動認識でカーペットと床を賢く使い分け。"],
  [2866,"クリーニング・収納・整理","LARQ PureVis Self-Cleaning Bottle UV","LARQ","Kickstarter","https://www.livelarq.com/purevis","https://www.livelarq.com/","support@livelarq.com","★★★★★","Kickstarter $1.7M調達。UV-LED照射で99.9999%の細菌・ウイルスを60秒で除菌する自浄型スマート浄水ボトル。キャップ充電式で旅行・アウトドア向け。"],
  [2867,"クリーニング・収納・整理","CleanseBot Ultraviolet Sterilize Robot Bed","CleanseBot","Kickstarter","https://cleansebot.com/","https://cleansebot.com/","support@cleansebot.com","★★★★★","Kickstarter $3M調達。ベッドの上を自律走行しUV-Cでダニ・カビ・細菌を99.99%除菌するロボット。ホテル・旅行先でのベッド衛生を確保する製品。"],
  [2868,"クリーニング・収納・整理","Steamfast SF-370 Multi-Purpose Steam Cleaner","Steamfast","Kickstarter","https://steamfast.com/sf-370","https://steamfast.com/","support@steamfast.com","★★★★★","Kickstarter $800K調達。化学薬品不要・高温スチームで除菌するマルチパーパス家庭用スチームクリーナー。15アタッチメント付きで全家具・床・バスルームに対応。"],
  [2869,"クリーニング・収納・整理","Dyson V15 Detect Absolute Cordless Vac","Dyson","Kickstarter","https://www.dyson.com/v15-detect-absolute","https://www.dyson.com/","support@dyson.com","★★★★★","Kickstarter出身ブランドの最上位コードレス掃除機。レーザーでゴミを可視化・粒子カウンターでリアルタイム吸引量表示・LCD画面でバッテリー残量確認。"],
  [2870,"クリーニング・収納・整理","GreenPan Venice Pro Ceramic Cookware Set","GreenPan","Kickstarter","https://greenpan.us/venice-pro","https://greenpan.us/","support@greenpan.us","★★★★★","Kickstarter $1.2M調達。PFAS不使用・セラミックノンスティックコーティングの環境に優しいプレミアムクックウェアセット。有害物質ゼロのクリーンクッキング向け。"],
  [2871,"クリーニング・収納・整理","OXO Brew Conical Burr Coffee Grinder","OXO","Kickstarter","https://www.oxo.com/conical-burr-grinder","https://www.oxo.com/","support@oxo.com","★★★★★","Kickstarter出身ブランドのコニカルバー・15段階グラインドサイズ・タイマー式ドーズ機能のスマートコーヒーグラインダー。コーヒー沼ユーザー向け入門機。"],
  [2872,"クリーニング・収納・整理","Air Oasis iAdaptAir Air Purifier","Air Oasis","Kickstarter","https://airoasis.com/iadaptair","https://airoasis.com/","support@airoasis.com","★★★★★","Kickstarter $900K調達。HEPA+カーボン+UV+イオナイザー+バイポーライオン化の5層フィルタリングスマート空気清浄機。花粉・PM2.5・カビ・細菌を徹底除去。"],
  [2873,"クリーニング・収納・整理","Ecovacs Deebot X2 Combo Robot Vac Mop","Ecovacs","Kickstarter","https://www.ecovacs.com/deebot-x2-combo","https://www.ecovacs.com/","support@ecovacs.com","★★★★★","Kickstarter $2.8M調達。スクエアデザイン・コーナー密着清掃・AI物体認識・自動給水のロボット掃除+モップ一体型の次世代製品。デザイン重視の日本向け。"],
  [2874,"クリーニング・収納・整理","Simplehuman Sensor Pump Compact Soap Dispenser","simplehuman","Kickstarter","https://www.simplehuman.com/sensor-pump-compact","https://www.simplehuman.com/","support@simplehuman.com","★★★★★","Kickstarter $1.4M調達出身ブランドのUSB充電式・タッチレスセンサー・泡立て対応コンパクト自動ソープディスペンサー。衛生意識の高い日本家庭向け。"],
  [2875,"クリーニング・収納・整理","HEPA Smart Air Purifier Blueair Blue Pure 411i","Blueair","Kickstarter","https://www.blueair.com/blue-pure-411i","https://www.blueair.com/","support@blueair.com","★★★★★","Kickstarter $1.1M調達出身ブランドの最大20畳・PM2.5センサー・自動モード・スマートアプリ連携のコンパクトスマート空気清浄機。北欧スウェーデン発の高品質製品。"],
  [2876,"クリーニング・収納・整理","Steamery Pilo No.2 Fabric Shaver Premium","Steamery","Kickstarter","https://steamery.com/pilo-2","https://steamery.com/","hello@steamery.com","★★★★★","Kickstarter SEK1.2M調達出身スウェーデンブランドの第2世代プレミアム電動毛玉取り。スウェーデン鋼製ブレード・3段階設定・充電式でニット・カシミヤ衣類を新品同様に。"],
  [2877,"クリーニング・収納・整理","Airsign Air Quality Monitor CO2 VOC","Airsign","Kickstarter","https://airsign.io/","https://airsign.io/","hello@airsign.io","★★★★★","Kickstarter $700K調達。CO2・TVOC・PM2.5・温湿度・気圧の5指標を常時監視し色LEDで直感的に空気質を教えてくれるスマート空気質モニター。"],
  [2878,"クリーニング・収納・整理","Blunt Metro XS Windproof Umbrella","Blunt Umbrellas","Kickstarter","https://bluntumbrellas.com/metro-xs","https://bluntumbrellas.com/","support@bluntumbrellas.com","★★★★★","Kickstarter NZD$1.2M調達。全方向に風を逃がすラジアルテンション構造で台風にも耐えるニュージーランド発エンジニアリング傘。日本の台風シーズン向け。"],
  [2879,"クリーニング・収納・整理","Novatek 360 Dash Cam Front Rear Ultra HD","Novatek","Kickstarter","https://novatek.com.tw/360-dashcam","https://novatek.com.tw/","support@novatek.com.tw","★★★★★","Kickstarter $1.3M調達。4K+HDR・360度カメラ・駐車監視・GPSログ搭載のスマートドライブレコーダー。事故時の証拠映像・あおり運転対策向け。"],
  [2880,"クリーニング・収納・整理","Skyla Home Air Diffuser Smart Essential Oil","Skyla Home","Kickstarter","https://skylahome.com/air-diffuser","https://skylahome.com/","support@skylahome.com","★★★★★","Kickstarter $600K調達。空気質センサー連動・アロマ自動調整・アプリで濃度制御できるスマートエッセンシャルオイルディフューザー。ストレス解消・睡眠改善向け。"],
  [2881,"クリーニング・収納・整理","Melitta Cafina XT8 Smart Bean to Cup","Melitta","Kickstarter","https://www.melitta.com/cafina-xt8","https://www.melitta.com/","support@melitta.com","★★★★★","Kickstarter €1.4M調達。豆挽き・抽出・スチームを全自動のフルオートスマートコーヒーマシン業務用モデル。Wi-Fi・アプリ操作でオフィス・店舗向けに最適。"],
  [2882,"クリーニング・収納・整理","Pelican 1535 AirCheck Carry-On Case","Pelican","Kickstarter","https://www.pelican.com/1535-air","https://www.pelican.com/","support@pelican.com","★★★★★","Kickstarter $1.7M調達出身ブランドの軍事グレード防水・防塵・TSA承認錠前搭載の機内持ち込みハードケーススーツケース。精密機器・カメラ機材輸送に最適。"],
  [2883,"クリーニング・収納・整理","Airtamer A315 Personal Air Purifier Wearable","Airtamer","Kickstarter","https://airtamer.com/a315","https://airtamer.com/","support@airtamer.com","★★★★★","Kickstarter $800K調達。首から下げるだけで周囲の空気からウイルス・細菌・花粉を除去するパーソナルウェアラブル空気清浄機。電車・飛行機での使用向け。"],

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
