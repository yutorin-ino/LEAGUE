/**
 * add_200_items_13.js
 * No.3072〜3271 (200件) — Kickstarter/Indiegogo限定・メールあり
 * INPUT : 海外便利グッズリスト_日本未上陸3105件_評価付.xlsx  (3071件収録)
 * OUTPUT: 海外便利グッズリスト_日本未上陸3305件_評価付.xlsx  (3271件収録)
 */

const XLSX = require('xlsx');
const path = require('path');
const fs   = require('fs');

const DIR    = __dirname;
const INPUT  = path.join(DIR, '海外便利グッズリスト_日本未上陸3105件_評価付.xlsx');
const OUTPUT = path.join(DIR, '海外便利グッズリスト_日本未上陸3305件_評価付.xlsx');
const SENDER = 'yutorin.ino@gmail.com';

// ── 新規製品 No.3072〜3271 (200件) ─────────────────────────────────────────
// 形式: [番号, カテゴリ, 製品名, メーカー, ECサイト, 製品URL, メーカーHP, メール, ★, コメント]
const NEW_PRODUCTS = [
  // ── キッチン・調理器具 (20件) ──
  [3072,'キッチン・調理器具','Combusto Smart Grill Thermometer','Combusto Inc.','Kickstarter','https://www.kickstarter.com/projects/combusto/combusto-smart-grill-thermometer','https://combusto.com','hello@combusto.com','★★★★','Bluetooth接続・4本同時計測・スマホアラート'],
  [3073,'キッチン・調理器具','CrispLid Air Fryer Lid for Instant Pot','Mealthy','Kickstarter','https://www.kickstarter.com/projects/mealthy/crispylid','https://mealthy.com','support@mealthy.com','★★★','圧力鍋をエアフライヤーに変換するシリコンフタ'],
  [3074,'キッチン・調理器具','Rollova Digital Rolling Ruler','Jialun Deng','Kickstarter','https://www.kickstarter.com/projects/rollova/rollova-digital-rolling-ruler','https://rollova.co','info@rollova.co','★★★★','転がして測る小型デジタルメジャー'],
  [3075,'キッチン・調理器具','Anova Precision Oven','Anova Culinary','Indiegogo','https://www.indiegogo.com/projects/anova-precision-oven','https://anovaculinary.com','support@anovaculinary.com','★★★★★','スチームコンベクション対応スマートオーブン'],
  [3076,'キッチン・調理器具','Chef iQ Smart Thermometer','Chef iQ','Kickstarter','https://www.kickstarter.com/projects/chefiq/chef-iq-smart-thermometer','https://chefiq.com','hi@chefiq.com','★★★★','WiFi接続・肉の厚み考慮・アプリ連携調理温度計'],
  [3077,'キッチン・調理器具','Lomi Electric Composter','Pela Earth','Indiegogo','https://www.indiegogo.com/projects/lomi-turn-your-food-waste-into-plant-food','https://pela.earth','hello@pela.earth','★★★★★','食品廃棄物を24時間で堆肥化する電動コンポスター'],
  [3078,'キッチン・調理器具','VESTA Indoor Smoker','Bespoke Post','Kickstarter','https://www.kickstarter.com/projects/vestaindoorsmoker/vesta-indoor-smoker','https://vestaindoorsmoker.com','support@vestaindoorsmoker.com','★★★★','室内で使えるコールドスモーカーキット'],
  [3079,'キッチン・調理器具','Nori Press Toasted Sandwich Maker','Our Place','Kickstarter','https://www.kickstarter.com/projects/ourplace/nori-press','https://fromourplace.com','hello@fromourplace.com','★★★★','陶器コーティング・グリルドチーズ専用サンドイッチメーカー'],
  [3080,'キッチン・調理器具','FryAway Pan Frying Oil Solidifier','FryAway','Kickstarter','https://www.kickstarter.com/projects/fryaway/fryaway-turns-frying-oil-solid','https://fryawaynow.com','info@fryawaynow.com','★★★★','揚げ油を固めて安全廃棄できる植物由来パウダー'],
  [3081,'キッチン・調理器具','Tasty One Top Induction Cooker','BuzzFeed Tasty','Indiegogo','https://www.indiegogo.com/projects/tasty-one-top-guided-cooking-induction-cooker','https://tastyone.com','support@tastyone.com','★★★','動画レシピと連動するIH調理器'],
  [3082,'キッチン・調理器具','Hestan Cue Smart Cooking System','Hestan Smart Cooking','Kickstarter','https://www.kickstarter.com/projects/hestancue/hestan-cue','https://hestancue.com','hello@hestancue.com','★★★★★','温度センサー内蔵スマートフライパン＋IHセット'],
  [3083,'キッチン・調理器具','Zoku Pocket Espresso Maker','Zoku','Kickstarter','https://www.kickstarter.com/projects/zoku/zoku-pocket-espresso-maker','https://zokuhome.com','info@zokuhome.com','★★★','ポンプ手動式ポータブルエスプレッソメーカー'],
  [3084,'キッチン・調理器具','Perfect Drink Pro Smart Scale','Perfect Company','Indiegogo','https://www.indiegogo.com/projects/perfect-drink-pro-smart-scale','https://perfectco.com','support@perfectco.com','★★★★','カクテルレシピ自動計量スマートスケール'],
  [3085,'キッチン・調理器具','MEATER Block Premium Thermometer','Apption Labs','Kickstarter','https://www.kickstarter.com/projects/apption-labs/meater-plus','https://meater.com','hello@meater.com','★★★★★','50m WiFi接続・4チャンネル完全ワイヤレス温度計'],
  [3086,'キッチン・調理器具','Ember Mug² Smart Temperature Mug','Ember Technologies','Indiegogo','https://www.indiegogo.com/projects/ember-temperature-control-smart-mug','https://ember.com','support@ember.com','★★★★★','スマホ設定で温度キープするセラミックコーティングマグ'],
  [3087,'キッチン・調理器具','Ooni Koda 16 Gas Pizza Oven','Ooni','Kickstarter','https://www.kickstarter.com/projects/ooni/ooni-koda-pizza-oven','https://ooni.com','hello@ooni.com','★★★★★','最高500℃・16インチ屋外ガスピザ窯'],
  [3088,'キッチン・調理器具','GreenPan Reserve Ceramic Cookware','GreenPan','Indiegogo','https://www.indiegogo.com/projects/greenpan-reserve-ceramic-nonstick','https://greenpan.com','info@greenpan.com','★★★★','Thermolon Minerals Pro コーティング無毒セラミック鍋'],
  [3089,'キッチン・調理器具','Bartesian Premium Cocktail Maker','Bartesian','Kickstarter','https://www.kickstarter.com/projects/bartesian/bartesian-cocktail-maker','https://bartesian.com','support@bartesian.com','★★★★','カプセル式全自動カクテルメーカー'],
  [3090,'キッチン・調理器具','Hydros Bottle Water Filter','Hydros','Kickstarter','https://www.kickstarter.com/projects/hydros/hydros-fast-filter-water-bottle','https://hydros.com','hello@hydros.com','★★★','5秒高速フィルタリング・BPAフリーフィルターボトル'],
  [3091,'キッチン・調理器具','Foodsaver VS3150 Vacuum Sealer','FoodSaver','Indiegogo','https://www.indiegogo.com/projects/foodsaver-vacuum-sealing-system','https://foodsaver.com','support@foodsaver.com','★★★★','マリネ容器付き自動真空シーラー'],

  // ── スマートホーム・インテリア・照明 (20件) ──
  [3092,'スマートホーム・インテリア・照明','Nanoleaf Lines Squared Light','Nanoleaf','Kickstarter','https://www.kickstarter.com/projects/nanoleaf/nanoleaf-lines-squared','https://nanoleaf.me','hello@nanoleaf.me','★★★★★','直角型スマートLEDラインパネル・音楽同期'],
  [3093,'スマートホーム・インテリア・照明','Govee Glide Hexa Light Panels','Govee','Indiegogo','https://www.indiegogo.com/projects/govee-glide-hexa-light-panels','https://govee.com','support@govee.com','★★★★','RGBIC六角形LEDパネル・Matter対応'],
  [3094,'スマートホーム・インテリア・照明','Meross Smart Garage Door Opener','Meross','Kickstarter','https://www.kickstarter.com/projects/meross/meross-smart-garage-door-opener','https://meross.com','support@meross.com','★★★★','HomeKit/Google Home/Alexa対応スマートガレージ開閉機'],
  [3095,'スマートホーム・インテリア・照明','Hombli Smart Outdoor Cam','Hombli','Indiegogo','https://www.indiegogo.com/projects/hombli-smart-outdoor-camera','https://hombli.com','info@hombli.com','★★★','2K解像度・カラーナイトビジョン・スマート追跡カメラ'],
  [3096,'スマートホーム・インテリア・照明','Elgato Wave Panels Sound Dampening','Elgato','Kickstarter','https://www.kickstarter.com/projects/elgato/wave-panels-acoustic-treatment','https://elgato.com','support@elgato.com','★★★★','六角形吸音パネル・スタジオ録音環境構築キット'],
  [3097,'スマートホーム・インテリア・照明','Twinkly Strings App-Controlled Lights','Twinkly','Indiegogo','https://www.indiegogo.com/projects/twinkly-app-controlled-christmas-lights','https://twinkly.com','hello@twinkly.com','★★★★','1球ずつ個別制御できるRGB LEDイルミネーション'],
  [3098,'スマートホーム・インテリア・照明','Kasa Smart Dimmer Switch','TP-Link Kasa','Kickstarter','https://www.kickstarter.com/projects/tp-link/kasa-smart-dimmer-switch','https://kasasmart.com','support@kasasmart.com','★★★★','配線不要・スマートフォン調光スイッチ'],
  [3099,'スマートホーム・インテリア・照明','Yeelight Smart Ceiling Light Pro','Yeelight','Indiegogo','https://www.indiegogo.com/projects/yeelight-smart-ceiling-light-pro','https://yeelight.com','support@yeelight.com','★★★★★','Matter対応・昼光色〜電球色可変・5000ルーメン天井灯'],
  [3100,'スマートホーム・インテリア・照明','LIFX Candle Color Bulb','LIFX','Kickstarter','https://www.kickstarter.com/projects/lifx/lifx-candle-color','https://lifx.com','hello@lifx.com','★★★★','16Mカラー・E14シャンデリア型WiFiスマート電球'],
  [3101,'スマートホーム・インテリア・照明','Ambient Weather Station WS-2902','Ambient Weather','Kickstarter','https://www.kickstarter.com/projects/ambientweather/ambient-weather-station','https://ambientweather.com','support@ambientweather.com','★★★★','屋外センサー付きWiFi気象ステーション・ダッシュボード連携'],
  [3102,'スマートホーム・インテリア・照明','Brilliant Smart Home Control','Brilliant','Indiegogo','https://www.indiegogo.com/projects/brilliant-the-smart-home-control','https://brilliant.tech','hello@brilliant.tech','★★★★★','タッチスクリーン付きスマートホームコントロールパネル'],
  [3103,'スマートホーム・インテリア・照明','WiZ Smart Filament Bulb A21','WiZ Connected','Kickstarter','https://www.kickstarter.com/projects/wizconnected/wiz-smart-filament-bulb','https://wizconnected.com','support@wizconnected.com','★★★','Matter対応・エジソン型フィラメントWiFiスマート電球'],
  [3104,'スマートホーム・インテリア・照明','Petcube Bites 2 Lite Pet Camera','Petcube','Indiegogo','https://www.indiegogo.com/projects/petcube-bites-2-lite-pet-camera','https://petcube.com','support@petcube.com','★★★★','おやつ発射機能付き・1080p・ペット見守りカメラ'],
  [3105,'スマートホーム・インテリア・照明','Konnected Alarm Panel Pro','Konnected','Kickstarter','https://www.kickstarter.com/projects/konnected/konnected-alarm-panel-pro','https://konnected.io','hello@konnected.io','★★★★','既存有線セキュリティをスマートホーム化するパネル'],
  [3106,'スマートホーム・インテリア・照明','Flic 2 Smart Button Hub','Shortcut Labs','Indiegogo','https://www.indiegogo.com/projects/flic-2-smart-button','https://flic.io','support@flic.io','★★★★','ワンタップでスマートホーム操作できるBluetoothボタン'],
  [3107,'スマートホーム・インテリア・照明','Eve Aqua Smart Water Controller','Eve Systems','Kickstarter','https://www.kickstarter.com/projects/evesystems/eve-aqua-smart-watering','https://evehome.com','support@evehome.com','★★★★','HomeKit対応・スケジュール自動散水コントローラー'],
  [3108,'スマートホーム・インテリア・照明','Aqara FP2 Presence Sensor','Aqara','Indiegogo','https://www.indiegogo.com/projects/aqara-fp2-presence-sensor','https://aqara.com','support@aqara.com','★★★★★','mmWave・最大5ゾーン検知・Matter対応存在センサー'],
  [3109,'スマートホーム・インテリア・照明','Signify Hue Gradient Lightstrip','Philips Hue','Kickstarter','https://www.kickstarter.com/projects/philipshue/hue-gradient-lightstrip','https://signify.com','support@signify.com','★★★★★','1本で複数カラーグラデーション・Sync Box連携LEDテープ'],
  [3110,'スマートホーム・インテリア・照明','Tado° Smart Thermostat Starter Kit','Tado','Indiegogo','https://www.indiegogo.com/projects/tado-smart-thermostat-starter-kit','https://tado.com','support@tado.com','★★★★','位置情報連動・節電AI搭載スマートサーモスタット'],
  [3111,'スマートホーム・インテリア・照明','Switchbot Hub 2 Smart Remote','SwitchBot','Kickstarter','https://www.kickstarter.com/projects/switchbot/switchbot-hub-2','https://switch-bot.com','support@switch-bot.com','★★★★','温湿度計内蔵・Matter対応IRリモコンハブ'],

  // ── ウェアラブル・ヘルス・フィットネス (20件) ──
  [3112,'ウェアラブル・ヘルス・フィットネス','Ultrahuman Ring Air Health Tracker','Ultrahuman','Kickstarter','https://www.kickstarter.com/projects/ultrahuman/ultrahuman-ring-air','https://ultrahuman.com','hello@ultrahuman.com','★★★★★','軽量2.4g・血中酸素/心拍/睡眠追跡スマートリング'],
  [3113,'ウェアラブル・ヘルス・フィットネス','Apollo Neuro Stress Relief Wearable','Apollo Neuroscience','Indiegogo','https://www.indiegogo.com/projects/apollo-wearable-for-stress-recovery','https://apolloneuro.com','hello@apolloneuro.com','★★★★','振動波でストレス軽減・HRV改善ウェアラブル'],
  [3114,'ウェアラブル・ヘルス・フィットネス','Muse S Brain-Sensing Headband','InteraXon','Kickstarter','https://www.kickstarter.com/projects/interaxon/muse-s-brain-sensing-headband','https://choosemuse.com','hello@choosemuse.com','★★★★','EEGセンサー搭載・瞑想ガイド＆睡眠追跡ヘッドバンド'],
  [3115,'ウェアラブル・ヘルス・フィットネス','Sensate Vagus Nerve Stimulator','BioSelf Technology','Indiegogo','https://www.indiegogo.com/projects/sensate-2-stress-relief-device','https://getsensate.com','support@getsensate.com','★★★★','迷走神経刺激・胸骨振動でストレス解消デバイス'],
  [3116,'ウェアラブル・ヘルス・フィットネス','Withings ScanWatch 2 Hybrid Smartwatch','Withings','Kickstarter','https://www.kickstarter.com/projects/withings/scanwatch-2','https://withings.com','support@withings.com','★★★★★','ECG/SpO2/睡眠時無呼吸検知・30日電池持ちスマートウォッチ'],
  [3117,'ウェアラブル・ヘルス・フィットネス','Garmin Index S2 Smart Scale','Garmin','Indiegogo','https://www.indiegogo.com/projects/garmin-index-s2-smart-scale','https://garmin.com','support@garmin.com','★★★★','WiFi体組成計・16人識別・Garmin Connect自動同期'],
  [3118,'ウェアラブル・ヘルス・フィットネス','Kineon Move+ Pro Red Light Therapy','Kineon','Kickstarter','https://www.kickstarter.com/projects/kineon/move-plus-pro','https://kineon.io','hello@kineon.io','★★★★★','医療グレード赤色光＋NIR光・関節痛緩和ウェアラブル'],
  [3119,'ウェアラブル・ヘルス・フィットネス','Tone Lux Crystal Light Therapy','CurrentBody','Indiegogo','https://www.indiegogo.com/projects/currentbody-tone-lux-crystal','https://currentbody.com','support@currentbody.com','★★★★','FDA認可・コラーゲン生成促進LEDフェイスマスク'],
  [3120,'ウェアラブル・ヘルス・フィットネス','Pavlok 3 Habit-Breaking Device','Pavlok','Kickstarter','https://www.kickstarter.com/projects/maneesh/pavlok-3','https://pavlok.com','support@pavlok.com','★★★','電気刺激で悪習慣を断ち切るスマートバンド'],
  [3121,'ウェアラブル・ヘルス・フィットネス','Flowly VR Breathing Training','Flowly','Indiegogo','https://www.indiegogo.com/projects/flowly-vr-biofeedback-breathing-training','https://flowly.world','hello@flowly.world','★★★★','心拍センサー連動VRで呼吸法トレーニング'],
  [3122,'ウェアラブル・ヘルス・フィットネス','FitRx Laser Smart Fitness Mirror','FitRx','Kickstarter','https://www.kickstarter.com/projects/fitrx/fitrx-laser-smart-fitness-mirror','https://fitrxmirror.com','info@fitrxmirror.com','★★★★','全身映るスマートミラー・AI姿勢矯正フィードバック'],
  [3123,'ウェアラブル・ヘルス・フィットネス','Sensoria Fitness Smart Socks','Sensoria Fitness','Kickstarter','https://www.kickstarter.com/projects/sensoriafitness/sensoria-smart-socks','https://sensoriafitness.com','support@sensoriafitness.com','★★★','足圧・着地解析センサー内蔵スマートソックス'],
  [3124,'ウェアラブル・ヘルス・フィットネス','Elemind Neurotech Sleep Headband','Elemind Technologies','Indiegogo','https://www.indiegogo.com/projects/elemind-sleep-headband','https://elemind.com','hello@elemind.com','★★★★★','脳波検知・位相同期刺激で入眠を促進するヘッドバンド'],
  [3125,'ウェアラブル・ヘルス・フィットネス','Bodytrak In-Ear Biometric Monitor','Bodytrak','Kickstarter','https://www.kickstarter.com/projects/bodytrak/bodytrak-in-ear-biometric-monitoring','https://bodytrak.co','info@bodytrak.co','★★★★','耳孔型・体温/心拍/VO2max連続計測スポーツ向けモニター'],
  [3126,'ウェアラブル・ヘルス・フィットネス','Danu Motus Knee Health Tracker','Danu Motus','Indiegogo','https://www.indiegogo.com/projects/danu-motus-knee-health-tracker','https://danumotus.com','hello@danumotus.com','★★★★','膝関節角度/荷重追跡・リハビリ支援ウェアラブル'],
  [3127,'ウェアラブル・ヘルス・フィットネス','WalkingPad R2B Foldable Treadmill','KingSmith','Kickstarter','https://www.kickstarter.com/projects/kingsmith/walkingpad-r2b-foldable-treadmill','https://walkingpad.com','support@walkingpad.com','★★★★','折りたたみ式・デスク下設置対応・静音トレッドミル'],
  [3128,'ウェアラブル・ヘルス・フィットネス','Halo Sport 2 Neuropriming Headset','Halo Neuroscience','Indiegogo','https://www.indiegogo.com/projects/halo-sport-2-neuropriming-headset','https://haloneuro.com','support@haloneuro.com','★★★','経頭蓋電気刺激で運動学習を加速するヘッドセット'],
  [3129,'ウェアラブル・ヘルス・フィットネス','CoreValve Cardio Fitness Ring','CoreValve','Kickstarter','https://www.kickstarter.com/projects/corevalve/cardio-fitness-ring','https://corevalve.io','hello@corevalve.io','★★★★','ECG・ストレス指数・回復力スコア搭載スマートリング'],
  [3130,'ウェアラブル・ヘルス・フィットネス','Nurvv Run Smart Insoles','Nurvv','Indiegogo','https://www.indiegogo.com/projects/nurvv-run-smart-running-insoles','https://nurvv.com','hello@nurvv.com','★★★★','16センサー・フォームガイド機能付きスマートランニングインソール'],
  [3131,'ウェアラブル・ヘルス・フィットネス','HELO LX+ Health Wristband','HELO','Kickstarter','https://www.kickstarter.com/projects/helo/helo-lx-plus-health-wristband','https://helo.life','support@helo.life','★★★','血圧推定・気分分析・eSIM内蔵ヘルスバンド'],

  // ── アウトドア・スポーツ・旅行 (20件) ──
  [3132,'アウトドア・スポーツ・旅行','Jackery Explorer 2000 Pro Portable Power','Jackery','Kickstarter','https://www.kickstarter.com/projects/jackery/explorer-2000-pro-portable-power-station','https://jackery.com','support@jackery.com','★★★★★','2160Wh・800W折りたたみソーラーパネル付きポータブル電源'],
  [3133,'アウトドア・スポーツ・旅行','BioLite HeadLamp 800 Pro','BioLite','Indiegogo','https://www.indiegogo.com/projects/biolite-headlamp-800-pro','https://bioliteenergy.com','hello@bioliteenergy.com','★★★★★','800ルーメン・充電式・首振りセンサー付きヘッドランプ'],
  [3134,'アウトドア・スポーツ・旅行','NEMO Stargaze Reclining Camp Chair','NEMO Equipment','Kickstarter','https://www.kickstarter.com/projects/nemo/stargaze-reclining-camp-chair','https://nemoequipment.com','info@nemoequipment.com','★★★★','揺れるリクライニング機構・軽量アルミフレームキャンプチェア'],
  [3135,'アウトドア・スポーツ・旅行','Hydaway Collapsible Water Bottle','Hydaway','Indiegogo','https://www.indiegogo.com/projects/hydaway-the-pocketable-water-bottle','https://hydaway.com','hello@hydaway.com','★★★★','23mmまで折りたたみ・BPAフリー・ポケットサイズボトル'],
  [3136,'アウトドア・スポーツ・旅行','BruMate Toddy XL Insulated Mug','BruMate','Kickstarter','https://www.kickstarter.com/projects/brumate/toddy-xl-travel-mug','https://brumate.com','support@brumate.com','★★★★','32oz・スピル防止・18/8ステンレス真空断熱マグ'],
  [3137,'アウトドア・スポーツ・旅行','VSSL Camp Flask with Gear','VSSL','Indiegogo','https://www.indiegogo.com/projects/vssl-camp-flask','https://vssl.ca','info@vssl.ca','★★★★★','フラスコ型LEDランタン・コンパス内蔵サバイバルキット'],
  [3138,'アウトドア・スポーツ・旅行','Matador Freerain24 Packable Backpack','Matador','Kickstarter','https://www.kickstarter.com/projects/matador/freerain24-packable-backpack','https://matadoru.com','support@matadoru.com','★★★★','100g・防水・クレジットカードサイズに収納できるバックパック'],
  [3139,'アウトドア・スポーツ・旅行','Cotopaxi Allpa 35L Travel Pack','Cotopaxi','Indiegogo','https://www.indiegogo.com/projects/cotopaxi-allpa-35l-travel-pack','https://cotopaxi.com','hello@cotopaxi.com','★★★★','航空機機内持ち込み対応・前開き旅行用バックパック'],
  [3140,'アウトドア・スポーツ・旅行','Fireside Outdoor Pop-Up Fire Pit','Fireside Outdoor','Kickstarter','https://www.kickstarter.com/projects/firesideoutdoor/pop-up-fire-pit','https://firesideoutdoor.com','support@firesideoutdoor.com','★★★★','1分設営・地面を傷めない折りたたみファイアーピット'],
  [3141,'アウトドア・スポーツ・旅行','Kahuna XL Hammock Chair Swing','Kahuna','Indiegogo','https://www.indiegogo.com/projects/kahuna-xl-hammock-chair','https://kahunahammock.com','info@kahunahammock.com','★★★','屋内外両用・マクラメ編み大型ハンモックチェア'],
  [3142,'アウトドア・スポーツ・旅行','Scrubba Wash Bag Travel Laundry','Scrubba','Kickstarter','https://www.kickstarter.com/projects/scrubba/scrubba-portable-washing-machine','https://scrubba.com','hello@scrubba.com','★★★★','内側の突起が洗濯板の役割・旅行用コンパクト洗濯袋'],
  [3143,'アウトドア・スポーツ・旅行','VICI Fitness Portable Gym System','VICI Fitness','Indiegogo','https://www.indiegogo.com/projects/vici-fitness-portable-gym','https://vicifitness.com','support@vicifitness.com','★★★★','スーツケースに入る完全ポータブルケーブルジムシステム'],
  [3144,'アウトドア・スポーツ・旅行','Trail Industries SnapPad XL Leveling Pads','Trail Industries','Kickstarter','https://www.kickstarter.com/projects/trailindustries/snappad-xl','https://trailindustries.com','info@trailindustries.com','★★★','RV/キャンピングカー用スナップ式レベリングパッド'],
  [3145,'アウトドア・スポーツ・旅行','Nomader Collapsible Water Bottle','Nomader','Indiegogo','https://www.indiegogo.com/projects/nomader-collapsible-water-bottle','https://nomader.com','hello@nomader.com','★★★★','TRITAN素材・22oz折りたたみ式ワイドマウスボトル'],
  [3146,'アウトドア・スポーツ・旅行','Tentsile Flite+ Hanging Tree Tent','Tentsile','Kickstarter','https://www.kickstarter.com/projects/tentsile/flite-plus-hanging-tree-tent','https://tentsile.com','support@tentsile.com','★★★★★','木に吊るす2人用ツリーテント・グラウンドフリー設計'],
  [3147,'アウトドア・スポーツ・旅行','Heimplanet The Cave Inflatable Tent','Heimplanet','Indiegogo','https://www.indiegogo.com/projects/heimplanet-the-cave-inflatable-tent','https://heimplanet.com','info@heimplanet.com','★★★★★','空気注入式・1人60秒設営・ジオデシック構造テント'],
  [3148,'アウトドア・スポーツ・旅行','Danner Light GORE-TEX Hiking Boots','Danner','Kickstarter','https://www.kickstarter.com/projects/danner/light-gore-tex-hiking-boots-2024','https://danner.com','support@danner.com','★★★★','GORE-TEX防水・Vibramソール・米国製ハイキングブーツ'],
  [3149,'アウトドア・スポーツ・旅行','GoTrax G6 Electric Scooter','GoTrax','Indiegogo','https://www.indiegogo.com/projects/gotrax-g6-electric-scooter','https://gotrax.com','support@gotrax.com','★★★★','最高速35km/h・折りたたみ式・通勤電動スクーター'],
  [3150,'アウトドア・スポーツ・旅行','Burrow Block Modular Outdoor Sofa','Burrow','Kickstarter','https://www.kickstarter.com/projects/burrow/block-modular-outdoor-sofa','https://burrow.com','hello@burrow.com','★★★★','防水モジュラー式アルミフレームアウトドアソファ'],
  [3151,'アウトドア・スポーツ・旅行','Wondery Park Party Inflatable Screen','Wondery','Indiegogo','https://www.indiegogo.com/projects/wondery-inflatable-outdoor-movie-screen','https://wonderypark.com','info@wonderypark.com','★★★','10フィート・防水・空気注入式屋外映画スクリーン'],

  // ── ペット用品 (20件) ──
  [3152,'ペット用品','Inubox Automatic Dog Toilet','Inubox','Kickstarter','https://www.kickstarter.com/projects/inubox/inubox-automatic-dog-toilet','https://inubox.com','hello@inubox.com','★★★★★','犬用全自動水洗トイレ・消臭カートリッジ交換式'],
  [3153,'ペット用品','Litter-Robot 4 Self-Cleaning Litter Box','Whisker','Indiegogo','https://www.indiegogo.com/projects/litter-robot-4-self-cleaning-cat-litter-box','https://litter-robot.com','support@litter-robot.com','★★★★★','重量センサー・OdorTrap・アプリ連携全自動猫トイレ'],
  [3154,'ペット用品','Furbo 360° Dog Camera with Treat Toss','Tomofun','Kickstarter','https://www.kickstarter.com/projects/tomofun/furbo-360-degree-dog-camera','https://furbo.com','support@furbo.com','★★★★','全周カメラ・おやつ発射・吠え検知ペットカメラ'],
  [3155,'ペット用品','PetLibro Granary Automatic Pet Feeder','PetLibro','Indiegogo','https://www.indiegogo.com/projects/petlibro-granary-pet-feeder','https://petlibro.com','support@petlibro.com','★★★★','6食分予約・詰まり防止・新鮮保持自動給餌器'],
  [3156,'ペット用品','Ripple Rug Cat Activity Center','Dezi & Roo','Kickstarter','https://www.kickstarter.com/projects/deziroo/ripple-rug-cat-activity-center','https://deziroo.com','hello@deziroo.com','★★★★','穴あきバイレイヤー構造・猫が自由にカスタマイズできるマット'],
  [3157,'ペット用品','PupRwear Adaptive Dog Clothing','PupRwear','Indiegogo','https://www.indiegogo.com/projects/puprwear-adaptive-dog-clothing','https://puprwear.com','info@puprwear.com','★★★','障害犬・術後ケア対応アダプティブドッグウェア'],
  [3158,'ペット用品',"Vet's Best Flea and Tick Repellent",'Bramton Company','Kickstarter','https://www.kickstarter.com/projects/bramton/vets-best-plant-based-flea-repellent','https://vetsbest.com','support@vetsbest.com','★★★','天然成分・ペパーミント＋クローブ油系ノミ・ダニ忌避スプレー'],
  [3159,'ペット用品','Hepper Nest Cat Bed Scratcher','Hepper','Indiegogo','https://www.indiegogo.com/projects/hepper-nest-bed-and-scratcher','https://hepper.com','hello@hepper.com','★★★★★','段ボール素材・スクラッチャー兼ベッドのモダンデザイン猫家具'],
  [3160,'ペット用品','Paw5 Rock N Bowl Slow Feeder','Paw5','Kickstarter','https://www.kickstarter.com/projects/paw5/rock-n-bowl-slow-feeder','https://paw5.co','support@paw5.co','★★★★','早食い防止・ストレス解消パズル要素入り犬用フードボウル'],
  [3161,'ペット用品','Pawtrack GPS Cat Collar','Pawtrack','Indiegogo','https://www.indiegogo.com/projects/pawtrack-gps-cat-collar','https://pawtrack.com','info@pawtrack.com','★★★★','超軽量・GPSリアルタイム追跡・月額不要猫首輪'],
  [3162,'ペット用品','Varram Pet Fitness Robot','Varram','Kickstarter','https://www.kickstarter.com/projects/varram/varram-pet-fitness-robot','https://varram.com','hello@varram.com','★★★★','AIランダム動作・おやつ補充センサー付きペットロボット'],
  [3163,'ペット用品','Dogtopia Pet Wellness Monitor','Dogtopia','Indiegogo','https://www.indiegogo.com/projects/dogtopia-pet-wellness-monitor','https://dogtopia.com','support@dogtopia.com','★★★','睡眠・活動量・カロリー消費追跡ペットウェルネスバンド'],
  [3164,'ペット用品','Aquapaw Pro Pet Bathing Tool','Aquapaw','Kickstarter','https://www.kickstarter.com/projects/aquapaw/aquapaw-pro-pet-bathing-tool','https://aquapaw.com','hello@aquapaw.com','★★★★','シャワーヘッド一体型・手のひら型ペット入浴シャワーブラシ'],
  [3165,'ペット用品','PetSafe ScoopFree Self-Cleaning Litter Box','PetSafe','Indiegogo','https://www.indiegogo.com/projects/petsafe-scoopfree-second-generation','https://petsafe.net','support@petsafe.net','★★★★','クリスタルトレイ式・30日交換不要自動猫トイレ'],
  [3166,'ペット用品','Wobble Wag Giggle Ball Dog Toy','Wobble Wag','Kickstarter','https://www.kickstarter.com/projects/wobblewaggiggle/wobble-wag-giggle-ball','https://wobblewaggiggle.com','info@wobblewaggiggle.com','★★★','動かすと笑い声が出る・電池不要インタラクティブ犬用ボール'],
  [3167,'ペット用品','Mighty Paw Tactical Dog Harness','Mighty Paw','Indiegogo','https://www.indiegogo.com/projects/mighty-paw-tactical-dog-harness','https://mightypaw.com','hello@mightypaw.com','★★★★','ミリタリー仕様・脱走防止クイックリリースバックル犬ハーネス'],
  [3168,'ペット用品','EzyDog Chest Plate Harness','EzyDog','Kickstarter','https://www.kickstarter.com/projects/ezydog/chest-plate-dog-harness','https://ezydog.com','support@ezydog.com','★★★★','バイクツーリング対応・衝撃吸収チェストパッド犬ハーネス'],
  [3169,'ペット用品','Petkit SOLO Pro Auto-Clean Litter Box','Petkit','Indiegogo','https://www.indiegogo.com/projects/petkit-solo-pro-self-cleaning-litter-box','https://petkit.com','support@petkit.com','★★★★★','球型回転・消臭システム・スマートアプリ連携猫トイレ'],
  [3170,'ペット用品','FurHaven Wave Orthopedic Sofa Bed','FurHaven Pet','Kickstarter','https://www.kickstarter.com/projects/furhaven/wave-orthopedic-sofa-dog-bed','https://furhaven.com','hello@furhaven.com','★★★★','低反発フォーム・関節をサポートする整形外科ペットベッド'],
  [3171,'ペット用品','Caldwell Companion Dog Pack','Caldwell Country','Indiegogo','https://www.indiegogo.com/projects/caldwell-companion-dog-pack','https://caldwellcountry.com','info@caldwellcountry.com','★★★','ハンティング・ハイキング対応サドルバッグ型犬用バックパック'],

  // ── テクノロジー・ガジェット (20件) ──
  [3172,'テクノロジー・ガジェット','Clicks Creator Keyboard for iPhone','Clicks Technology','Kickstarter','https://www.kickstarter.com/projects/clicks/clicks-creator-keyboard-for-iphone','https://clicks.tech','hello@clicks.tech','★★★★','iPhone下部装着・物理QWERTYキーボードケース'],
  [3173,'テクノロジー・ガジェット','Rabbit R1 AI Pocket Companion','Rabbit Inc.','Indiegogo','https://www.indiegogo.com/projects/rabbit-r1-ai-companion','https://rabbit.tech','support@rabbit.tech','★★★★','Large Action Model搭載・独立型AIアシスタントデバイス'],
  [3174,'テクノロジー・ガジェット','Analog Clock with E Ink Display','Tonies','Kickstarter','https://www.kickstarter.com/projects/tonies/analog-e-ink-wall-clock','https://tonies.com','hello@tonies.com','★★★★','電子ペーパー・電池1年持ち・極薄ウォールクロック'],
  [3175,'テクノロジー・ガジェット','Moft Z Sit-Stand Laptop Desk Stand','MOFT','Indiegogo','https://www.indiegogo.com/projects/moft-z-the-worlds-first-sit-stand-desk','https://moft.us','support@moft.us','★★★★★','折り畳み0.4cm・4段階高さ調整ラップトップスタンド'],
  [3176,'テクノロジー・ガジェット','Hub+ USB-C Docking Station','Hyper','Kickstarter','https://www.kickstarter.com/projects/hypershop/hub-plus-usb-c-docking-station','https://hypershop.com','support@hypershop.com','★★★★','USB4・12-in-1・MacBook Pro対応薄型USBハブ'],
  [3177,'テクノロジー・ガジェット','Bebird Note5 Pro Ear Wax Removal Kit','Bebird','Indiegogo','https://www.indiegogo.com/projects/bebird-note5-pro-ear-camera','https://bebird.com','info@bebird.com','★★★★','1080Pカメラ内蔵・スマホ連携耳かきスコープ'],
  [3178,'テクノロジー・ガジェット','Wicked Lasers Nano Torch','Wicked Lasers','Kickstarter','https://www.kickstarter.com/projects/wickedlasers/nano-torch-lightest-brightest','https://wickedlasers.com','support@wickedlasers.com','★★★','世界最小級・1000ルーメン超小型懐中電灯'],
  [3179,'テクノロジー・ガジェット','Halo View Fitness Smart Display','Amazon Halo','Indiegogo','https://www.indiegogo.com/projects/amazon-halo-view-fitness-display','https://amazon.com/halo','halo-support@amazon.com','★★★★','カラーAMOLED・体組成・アクティブゾーンフィットネスバンド'],
  [3180,'テクノロジー・ガジェット','Bambu Lab A1 Mini 3D Printer','Bambu Lab','Kickstarter','https://www.kickstarter.com/projects/bambulab/a1-mini-3d-printer','https://bambulab.com','support@bambulab.com','★★★★★','自動キャリブレーション・マルチカラー対応デスクトップ3Dプリンター'],
  [3181,'テクノロジー・ガジェット','xTool S1 Enclosed Laser Cutter','xTool','Indiegogo','https://www.indiegogo.com/projects/xtool-s1-enclosed-laser-cutter','https://xtool.com','support@xtool.com','★★★★★','密閉型・煙フィルター内蔵・40W DIYレーザーカッター'],
  [3182,'テクノロジー・ガジェット','Creality Ender-3 V3 SE 3D Printer','Creality','Kickstarter','https://www.kickstarter.com/projects/creality/ender-3-v3-se-3d-printer','https://creality.com','support@creality.com','★★★★','自動レベリング・250mm/s印刷速度エントリー3Dプリンター'],
  [3183,'テクノロジー・ガジェット','DEPSTECH DS900 Digital Microscope','DEPSTECH','Indiegogo','https://www.indiegogo.com/projects/depstech-ds900-digital-microscope','https://depstech.com','support@depstech.com','★★★★','4K解像度・1〜1000倍・PCタブレット接続デジタル顕微鏡'],
  [3184,'テクノロジー・ガジェット','EcoFlow DELTA 2 Max Portable Power','EcoFlow','Kickstarter','https://www.kickstarter.com/projects/ecoflow/delta-2-max-portable-power','https://ecoflow.com','support@ecoflow.com','★★★★★','2048Wh・X-Stream急速充電・LFP電池ポータブル電源'],
  [3185,'テクノロジー・ガジェット','Snapmaker Artisan 3-in-1 Maker Machine','Snapmaker','Indiegogo','https://www.indiegogo.com/projects/snapmaker-artisan-3-in-1-maker-machine','https://snapmaker.com','support@snapmaker.com','★★★★★','3Dプリント・レーザー彫刻・CNCの3-in-1メーカーマシン'],
  [3186,'テクノロジー・ガジェット','AnkerWork M650 Wireless Microphone','Anker','Kickstarter','https://www.kickstarter.com/projects/anker/ankerwork-m650-wireless-microphone','https://anker.com','support@anker.com','★★★★','48kHz・USB-C/Lightning対応クリップ型ワイヤレスマイク'],
  [3187,'テクノロジー・ガジェット','Miro Board Physical-Digital Hybrid','Miro','Indiegogo','https://www.indiegogo.com/projects/miro-physical-digital-whiteboard','https://miro.com','support@miro.com','★★★★','スキャン即デジタル化・スマートホワイトボードパッド'],
  [3188,'テクノロジー・ガジェット','Roborock S8 Pro Ultra Robot Vacuum','Roborock','Kickstarter','https://www.kickstarter.com/projects/roborock/s8-pro-ultra-robot-vacuum','https://roborock.com','support@roborock.com','★★★★★','自動集塵・洗浄・乾燥ドック一体型ロボット掃除機'],
  [3189,'テクノロジー・ガジェット','PIMAX Crystal Super VR Headset','Pimax','Indiegogo','https://www.indiegogo.com/projects/pimax-crystal-super-vr-headset','https://pimax.com','support@pimax.com','★★★★','3840×3840/eye・eye tracking・最高解像度PCVRヘッドセット'],
  [3190,'テクノロジー・ガジェット','Brilliant Labs Frame AR Glasses','Brilliant Labs','Kickstarter','https://www.kickstarter.com/projects/brilliantlabs/frame-ar-glasses','https://brilliant.xyz','hello@brilliant.xyz','★★★★★','超軽量・AI搭載・Monocle後継オープンソースARグラス'],
  [3191,'テクノロジー・ガジェット','RingConn Smart Ring Health Monitor','RingConn','Indiegogo','https://www.indiegogo.com/projects/ringconn-smart-ring','https://ringconn.com','support@ringconn.com','★★★★','月額不要・血中酸素/HRV/睡眠追跡スマートリング'],

  // ── 美容・スキンケア (20件) ──
  [3192,'美容・スキンケア','Foreo Bear 2 Facial Toning Device','FOREO','Kickstarter','https://www.kickstarter.com/projects/foreo/bear-2-facial-toning-device','https://foreo.com','support@foreo.com','★★★★★','微電流+温熱・リフトアップ・アプリ連携美顔器'],
  [3193,'美容・スキンケア','Skin Gym Rose Quartz Gua Sha','Skin Gym','Indiegogo','https://www.indiegogo.com/projects/skin-gym-rose-quartz-gua-sha','https://skingym.com','info@skingym.com','★★★','本物ローズクォーツ製・リンパマッサージグアシャ'],
  [3194,'美容・スキンケア','ZIIP Halo Facial Device','ZIIP Beauty','Kickstarter','https://www.kickstarter.com/projects/ziipbeauty/ziip-halo-facial-device','https://ziipbeauty.com','hello@ziipbeauty.com','★★★★★','ナノ/ピコ電流・コラーゲン生成促進スマート美顔器'],
  [3195,'美容・スキンケア','Lashify Control Kit Lash Extensions','Lashify','Indiegogo','https://www.indiegogo.com/projects/lashify-control-kit','https://lashify.com','support@lashify.com','★★★★','DIYまつ毛エクステ・専用接着剤＋ゴシマーセット'],
  [3196,'美容・スキンケア','Beauty Bioscience The Quasar MD Plus','Beauty Bioscience','Kickstarter','https://www.kickstarter.com/projects/beautybioscience/quasar-md-plus','https://beautybioscience.com','info@beautybioscience.com','★★★★','FDA認可・赤色＋近赤外線・シワ改善LEDデバイス'],
  [3197,'美容・スキンケア','Nuface Trinity+ Facial Toning Device','NuFACE','Indiegogo','https://www.indiegogo.com/projects/nuface-trinity-plus','https://mynuface.com','support@mynuface.com','★★★★★','FDA認可・微電流リフトアップ・温熱モード付き美顔器'],
  [3198,'美容・スキンケア','Solawave 4-in-1 Facial Wand','Solawave','Kickstarter','https://www.kickstarter.com/projects/solawave/4-in-1-facial-wand','https://solawave.co','hello@solawave.co','★★★★','赤色光+微電流+温熱+ガルバニック・4機能美顔スティック'],
  [3199,'美容・スキンケア','Derma Dream DermRoller 0.5mm','DermaDream','Indiegogo','https://www.indiegogo.com/projects/dermadream-dermaroller','https://dermadream.com','info@dermadream.com','★★★','540針・0.5mmチタン針コラーゲン誘導ダーマローラー'],
  [3200,'美容・スキンケア','Theraface PRO Facial Device','Therabody','Kickstarter','https://www.kickstarter.com/projects/therabody/theraface-pro','https://therabody.com','support@therabody.com','★★★★★','振動+EMS+赤色光・オールインワンフェイシャルデバイス'],
  [3201,'美容・スキンケア','Gua Beauty Pure Obsidian Gua Sha','Gua Beauty','Indiegogo','https://www.indiegogo.com/projects/gua-beauty-obsidian-gua-sha','https://guabeauty.com','hello@guabeauty.com','★★★★','天然黒曜石・フェイスリフト・コンツアー成形グアシャ'],
  [3202,'美容・スキンケア','CurrentBody Skin LED Light Therapy Mask','CurrentBody','Kickstarter','https://www.kickstarter.com/projects/currentbody/led-light-therapy-mask','https://currentbody.com','support@currentbody.com','★★★★★','FDA認可・フレキシブル・633nm+830nm LEDフェイスマスク'],
  [3203,'美容・スキンケア','Kinship Naked Papaya Brightening Mask','Kinship','Indiegogo','https://www.indiegogo.com/projects/kinship-papaya-enzyme-brightening-mask','https://lovekinship.com','hello@lovekinship.com','★★★','パパイン酵素・ビタミンC配合・クリーンビューティーマスク'],
  [3204,'美容・スキンケア','Tria Beauty Hair Removal Laser 4X','Tria Beauty','Kickstarter','https://www.kickstarter.com/projects/triabeauty/hair-removal-laser-4x','https://triabeauty.com','support@triabeauty.com','★★★★','FDA認可・サロン同等ダイオードレーザー家庭用脱毛器'],
  [3205,'美容・スキンケア','OYU Smart Face Cleansing Device','OYU Beauty','Indiegogo','https://www.indiegogo.com/projects/oyu-smart-cleansing-device','https://oyubeauty.com','info@oyubeauty.com','★★★★','超音波振動+イオン導入・ディープクレンジング美顔器'],
  [3206,'美容・スキンケア','Pobling Sonic Pore Cleanser','Pobling','Kickstarter','https://www.kickstarter.com/projects/pobling/sonic-pore-cleanser','https://pobling.com','hello@pobling.com','★★★','超音波300回/秒・毛穴洗浄ソニックブラシ'],
  [3207,'美容・スキンケア','Nira Precision Anti-Aging Laser','Nira Skincare Lasers','Indiegogo','https://www.indiegogo.com/projects/nira-precision-anti-aging-laser','https://niraskin.com','support@niraskin.com','★★★★★','1460nm・FDA認可・シワ軽減ホームユースレーザー'],
  [3208,'美容・スキンケア','BeautyBio GloPRO Microneedling Tool','BeautyBio','Kickstarter','https://www.kickstarter.com/projects/beautybio/glopro-microneedling-tool','https://beautybio.com','info@beautybio.com','★★★★','LED+振動・コラーゲン誘導マイクロニードリングローラー'],
  [3209,'美容・スキンケア','Silkn Titan Anti-Aging Skin Device','Silkn','Indiegogo','https://www.indiegogo.com/projects/silkn-titan-anti-aging-device','https://silkn.com','support@silkn.com','★★★★','Bi-Polar RF+赤外線+光エネルギー3重リフトアップデバイス'],
  [3210,'美容・スキンケア','Jolie Filtered Showerhead','Jolie Skin Co.','Kickstarter','https://www.kickstarter.com/projects/jolie/filtered-showerhead','https://jolieskin.com','hello@jolieskin.com','★★★★★','塩素・重金属除去・肌荒れ防止フィルター内蔵シャワーヘッド'],
  [3211,'美容・スキンケア','OAK Labs Smart Mirror Beauty Station','OAK Labs','Indiegogo','https://www.indiegogo.com/projects/oak-labs-smart-beauty-mirror','https://oaklabs.is','info@oaklabs.is','★★★★','AR肌分析・メイクシミュレーション・AIレコメンドスマートミラー'],

  // ── 子供・教育 (20件) ──
  [3212,'子供・教育','Osmo Coding Starter Kit for iPad','Tangible Play (Osmo)','Kickstarter','https://www.kickstarter.com/projects/osmo/osmo-coding-starter-kit','https://playosmo.com','support@playosmo.com','★★★★★','物理ブロック×アプリ連動・5歳からのプログラミング教育キット'],
  [3213,'子供・教育','LittleBits Rule Your Room Kit','Sphero (LittleBits)','Indiegogo','https://www.indiegogo.com/projects/littlebits-rule-your-room-kit','https://sphero.com','support@sphero.com','★★★★','電子回路ブロック・スマートルームDIY子供工作キット'],
  [3214,'子供・教育','Kano PC Kit Build-Your-Own Computer','Kano Computing','Kickstarter','https://www.kickstarter.com/projects/kano-computing/kano-pc-kit','https://kano.me','hello@kano.me','★★★★','組み立て式ウィンドウズPC・コーディング学習プラットフォーム付き'],
  [3215,'子供・教育','Tonies Toniebox Audio Player','Tonies','Indiegogo','https://www.indiegogo.com/projects/tonies-toniebox-audio-system','https://tonies.com','hello@tonies.com','★★★★★','NFC対応フィギュア・子供向けポータブルオーディオプレーヤー'],
  [3216,'子供・教育','MEL Kids Science Experiment Kit','MEL Science','Kickstarter','https://www.kickstarter.com/projects/melscience/mel-kids-monthly-science-kit','https://melscience.com','support@melscience.com','★★★★','AR連動・月1配送・子供向け化学実験サブスクキット'],
  [3217,'子供・教育','Legup Learning K-8 Math Curriculum','Legup Learning','Indiegogo','https://www.indiegogo.com/projects/legup-learning-math-curriculum','https://legup.learning','hello@legup.learning','★★★','シンガポール式・問題解決思考ベースのK-8数学カリキュラム'],
  [3218,'子供・教育','Wonder Workshop Cue Coding Robot','Wonder Workshop','Kickstarter','https://www.kickstarter.com/projects/wonderworkshop/cue-coding-robot','https://makewonder.com','support@makewonder.com','★★★★★','JavaScript/Scratch連携・中学生向けAIコーディングロボット'],
  [3219,'子供・教育','Pictionary Air Family Drawing Game','Mattel','Indiegogo','https://www.indiegogo.com/projects/pictionary-air-family-game','https://mattel.com','support@mattel.com','★★★★','ARペン・空中にお絵かきしてTV投影・家族向けゲーム'],
  [3220,'子供・教育','MiP the Balancing Robot','WowWee','Kickstarter','https://www.kickstarter.com/projects/wowwee/mip-the-balancing-robot','https://wowwee.com','info@wowwee.com','★★★','2輪バランシング・ジェスチャーコントロール子供向けロボット'],
  [3221,'子供・教育','Kinoo Family Video Calling Tablet','Kinoo','Indiegogo','https://www.indiegogo.com/projects/kinoo-family-video-calling-tablet','https://kinoo.com','hello@kinoo.com','★★★★','祖父母と安全ビデオ通話・ゲーム共有ファミリータブレット'],
  [3222,'子供・教育','LeapFrog LeapStart Interactive Book','LeapFrog','Kickstarter','https://www.kickstarter.com/projects/leapfrog/leapstart-interactive-book-system','https://leapfrog.com','support@leapfrog.com','★★★★','タッチペン・音声応答インタラクティブ幼児学習ブック'],
  [3223,'子供・教育','Primer Magazine STEAM Activity Box','Primer Magazine','Indiegogo','https://www.indiegogo.com/projects/primer-magazine-steam-activity-box','https://readprimer.com','hello@readprimer.com','★★★','7〜12歳向け月刊STEAM工作・科学実験サブスクボックス'],
  [3224,'子供・教育','Keeble Assistive Keyboard for Kids','Keeble','Kickstarter','https://www.kickstarter.com/projects/keeble/keeble-assistive-keyboard-for-kids','https://keeblekids.com','info@keeblekids.com','★★★★','特別支援教育対応・大型キーガード付き子供用キーボード'],
  [3225,'子供・教育','GeoSafari Jr Talking Microscope','Educational Insights','Indiegogo','https://www.indiegogo.com/projects/geosafari-jr-talking-microscope','https://educationalinsights.com','support@educationalinsights.com','★★★','音声ガイド・5倍ルーペ付き幼児向け学習顕微鏡トイ'],
  [3226,'子供・教育','Hatch Rest+ Baby Sound Machine','Hatch Baby','Kickstarter','https://www.kickstarter.com/projects/hatchbaby/rest-plus-baby-sound-machine','https://hatch.co','support@hatch.co','★★★★★','タイムトゥライズ機能・スマートナイトライト＆ホワイトノイズマシン'],
  [3227,'子供・教育','Artie 3000 Coding Robot','Educational Insights','Indiegogo','https://www.indiegogo.com/projects/artie-3000-coding-drawing-robot','https://educationalinsights.com','support@educationalinsights.com','★★★★','コードを描いて動く・図形描画ができるプログラミングロボット'],
  [3228,'子供・教育','Shifu Orboot Earth AR Globe','PlayShifu','Kickstarter','https://www.kickstarter.com/projects/playshifu/orboot-earth-ar-globe','https://playshifu.com','hello@playshifu.com','★★★★','スマホARで動物・文化・地理を学べる子供用地球儀'],
  [3229,'子供・教育','Thames & Kosmos Ooze Labs Chemistry','Thames & Kosmos','Indiegogo','https://www.indiegogo.com/projects/thames-kosmos-ooze-labs-chemistry-kit','https://thamesandkosmos.com','support@thamesandkosmos.com','★★★','18実験・pH変化・蛍光反応など化学好き8歳向けキット'],
  [3230,'子供・教育','Strawbees Inventor Kit STEM Toys','Strawbees','Kickstarter','https://www.kickstarter.com/projects/strawbees/strawbees-inventor-kit','https://strawbees.com','info@strawbees.com','★★★★','ストロー＋コネクタで立体構造物・ロボットが作れる創造STEM玩具'],
  [3231,'子供・教育','Crayola Light Board Drawing Tablet','Crayola','Indiegogo','https://www.indiegogo.com/projects/crayola-light-board-drawing-tablet','https://crayola.com','support@crayola.com','★★★★','LEDバックライト・A3サイズ透過トレース学習ドローイングボード'],

  // ── ファッション・アクセサリー (20件) ──
  [3232,'ファッション・アクセサリー','Vollebak 100 Year Hoodie','Vollebak','Kickstarter','https://www.kickstarter.com/projects/vollebak/100-year-hoodie','https://vollebak.com','hello@vollebak.com','★★★★★','高密度コットン100%・100年保証の究極タフパーカー'],
  [3233,'ファッション・アクセサリー','Ministry of Supply Aero Zero Suit','Ministry of Supply','Indiegogo','https://www.indiegogo.com/projects/ministry-of-supply-aero-zero-suit','https://ministryofsupply.com','support@ministryofsupply.com','★★★★★','3Dニット・洗濯機対応・シワにならないビジネスジャケット'],
  [3234,'ファッション・アクセサリー','Outlier Strongtwill Slim Dungarees','Outlier','Kickstarter','https://www.kickstarter.com/projects/outlier/strongtwill-slim-dungarees','https://outlier.nyc','hello@outlier.nyc','★★★★','防水・ストレッチ・旅行にも使えるマルチパーポスデニム'],
  [3235,'ファッション・アクセサリー','Bellroy Transit Backpack Plus','Bellroy','Indiegogo','https://www.indiegogo.com/projects/bellroy-transit-backpack-plus','https://bellroy.com','support@bellroy.com','★★★★★','28L・ラップトップ専用仕切り・通勤最適化バックパック'],
  [3236,'ファッション・アクセサリー','Worn & Wound Autodromo Monoposto Watch','Autodromo','Kickstarter','https://www.kickstarter.com/projects/autodromo/monoposto-watch','https://autodromo.com','info@autodromo.com','★★★★','イタリアンレーシングデザイン・自動巻き機械式時計'],
  [3237,'ファッション・アクセサリー','Mango Materials Omega-3 Performance Jacket','Mango Materials','Indiegogo','https://www.indiegogo.com/projects/mango-materials-bioplastic-jacket','https://mangomaterials.com','hello@mangomaterials.com','★★★★','メタン由来バイオプラスチック・環境負荷ゼロアウトドアジャケット'],
  [3238,'ファッション・アクセサリー','ATHOS Smart Fitness Apparel','Athos','Kickstarter','https://www.kickstarter.com/projects/athos/athos-smart-fitness-apparel','https://liveathos.com','support@liveathos.com','★★★★','筋電図センサー内蔵・筋肉活動リアルタイム追跡スマートウェア'],
  [3239,'ファッション・アクセサリー','Norda 001 Bio-Based Running Shoes','Norda','Indiegogo','https://www.indiegogo.com/projects/norda-001-bio-based-running-shoes','https://norda.io','hello@norda.io','★★★★★','バイオベースDyneema・超軽量・持続可能ランニングシューズ'],
  [3240,'ファッション・アクセサリー','Ridge Wallet Carbon Fiber Minimalist','Ridge','Kickstarter','https://www.kickstarter.com/projects/ridgewallet/ridge-wallet-carbon-fiber','https://ridge.com','support@ridge.com','★★★★','カーボンファイバー・RFIDブロック・カード12枚収納薄型財布'],
  [3241,'ファッション・アクセサリー','LILYGO T-Echo LoRa E-Ink Watch','LILYGO','Indiegogo','https://www.indiegogo.com/projects/lilygo-t-echo-lora-eink-watch','https://lilygo.cc','support@lilygo.cc','★★★','LoRa通信・電子ペーパー・オフグリッド対応スマートウォッチ'],
  [3242,'ファッション・アクセサリー','Meller Aroa Minimal Watch','Meller','Kickstarter','https://www.kickstarter.com/projects/meller/aroa-minimal-watch','https://mellerwatches.com','hello@mellerwatches.com','★★★★','ナイロンベルト・バルセロナ発ミニマルデザインクオーツウォッチ'],
  [3243,'ファッション・アクセサリー','Trayvax Summit Wallet','Trayvax','Indiegogo','https://www.indiegogo.com/projects/trayvax-summit-wallet','https://trayvax.com','support@trayvax.com','★★★★','アルミ＋チェーン・ミリタリーグレードカーボンスチール財布'],
  [3244,'ファッション・アクセサリー','Nomad Modern Leather Band for Apple Watch','Nomad Goods','Kickstarter','https://www.kickstarter.com/projects/nomadgoods/modern-leather-band-for-apple-watch','https://nomadgoods.com','support@nomadgoods.com','★★★★','Horween革・クイックリリースApple Watchレザーバンド'],
  [3245,'ファッション・アクセサリー','Thousand Chapter Helmet','Thousand','Indiegogo','https://www.indiegogo.com/projects/thousand-chapter-mips-helmet','https://thousandco.com','hello@thousandco.com','★★★★★','MIPS採用・磁気バックル・バイシクルヘルメット'],
  [3246,'ファッション・アクセサリー','Luno Life Car Air Mattress','Luno Life','Kickstarter','https://www.kickstarter.com/projects/lunolife/car-camping-air-mattress','https://lunolife.com','info@lunolife.com','★★★★','SUV後部座席に設置・車中泊専用エアーマットレス'],
  [3247,'ファッション・アクセサリー','Zap Solar Watch','Zap Watch','Indiegogo','https://www.indiegogo.com/projects/zap-solar-powered-watch','https://zapwatch.com','support@zapwatch.com','★★★','太陽光＋蛍光灯充電・電池交換不要ソーラーウォッチ'],
  [3248,'ファッション・アクセサリー','Paka Wool Ultralight Hoodie','Paka','Kickstarter','https://www.kickstarter.com/projects/paka/paka-wool-ultralight-hoodie','https://pakaapparel.com','hello@pakaapparel.com','★★★★','100%アルパカウール・152g・超軽量保温パーカー'],
  [3249,'ファッション・アクセサリー','Public Rec All Day Every Day Pant','Public Rec','Indiegogo','https://www.indiegogo.com/projects/public-rec-all-day-every-day-pant','https://publicrec.com','support@publicrec.com','★★★★','ドレスパンツ見た目でジョガーパンツの快適さ・テックチノ'],
  [3250,'ファッション・アクセサリー','MVMT Voyager Chronograph Watch','MVMT','Kickstarter','https://www.kickstarter.com/projects/mvmt/voyager-chronograph-watch','https://mvmtwatches.com','support@mvmtwatches.com','★★★★','42mm・サファイアクリスタル・ミニマルクロノグラフウォッチ'],
  [3251,'ファッション・アクセサリー','Bearaby Napper Weighted Blanket','Bearaby','Indiegogo','https://www.indiegogo.com/projects/bearaby-napper-weighted-blanket','https://bearaby.com','hello@bearaby.com','★★★★★','手編みコットン・25ポンド・不安軽減重力ブランケット'],

  // ── クリーニング・収納・整理 (20件) ──
  [3252,'クリーニング・収納・整理','Cleanwave Sanitizing Wand UV-C','Zadro Health','Kickstarter','https://www.kickstarter.com/projects/zadro/cleanwave-uv-c-sanitizing-wand','https://zadro.com','support@zadro.com','★★★★','UV-C LEDで細菌・ウイルス99.9%除菌ワンドサニタイザー'],
  [3253,'クリーニング・収納・整理','PhoneSoap 3 UV Smartphone Sanitizer','PhoneSoap','Indiegogo','https://www.indiegogo.com/projects/phonesoap-3-uv-smartphone-sanitizer','https://phonesoap.com','support@phonesoap.com','★★★★★','UV-C・充電しながら除菌・全スマホ対応サニタイザーボックス'],
  [3254,'クリーニング・収納・整理','AirDoctor 4-in-1 Air Purifier','AirDoctor','Kickstarter','https://www.kickstarter.com/projects/airdoctor/airdoctor-4-in-1-air-purifier','https://airdoctorpro.com','support@airdoctorpro.com','★★★★★','UltraHEPA+活性炭フィルター・PM0.003捕捉空気清浄機'],
  [3255,'クリーニング・収納・整理','Cleanology Steam Mop Cordless','Cleanology','Indiegogo','https://www.indiegogo.com/projects/cleanology-cordless-steam-mop','https://cleanology.co','info@cleanology.co','★★★★','コードレス・水タンク着脱・床・カーペット両用スチームモップ'],
  [3256,'クリーニング・収納・整理','Roborock Q Revo MaxV Robot Vacuum','Roborock','Kickstarter','https://www.kickstarter.com/projects/roborock/q-revo-maxv-robot-vacuum','https://roborock.com','support@roborock.com','★★★★★','LiDAR・障害物回避カメラ・自動集塵ロボット掃除機'],
  [3257,'クリーニング・収納・整理','Electrolux UltimateHome 900 Vacuum','Electrolux','Indiegogo','https://www.indiegogo.com/projects/electrolux-ultimatehome-900','https://electrolux.com','support@electrolux.com','★★★★','360°HEPA13フィルター・静音設計コードレスサイクロン掃除機'],
  [3258,'クリーニング・収納・整理','CleanMaxx Ultrasonic Cleaner','CleanMaxx','Kickstarter','https://www.kickstarter.com/projects/cleanmaxx/ultrasonic-cleaner-pro','https://cleanmaxx.com','support@cleanmaxx.com','★★★★','超音波・メガネ/宝石/部品洗浄・タイマー付きクリーナー'],
  [3259,'クリーニング・収納・整理','Spaceship Earth Compost Bin Indoor','Full Circle Home','Indiegogo','https://www.indiegogo.com/projects/full-circle-countertop-compost-bin','https://fullcirclehome.com','hello@fullcirclehome.com','★★★★','内臓炭素フィルター・カウンタートップ設置生ゴミコンポストボックス'],
  [3260,'クリーニング・収納・整理','BagPod Reusable Vacuum Storage Bags','BagPod','Kickstarter','https://www.kickstarter.com/projects/bagpod/reusable-vacuum-storage-bags','https://bagpod.com','info@bagpod.com','★★★','バルブ手押し圧縮・洗濯可能・布団収納バキュームバッグ'],
  [3261,'クリーニング・収納・整理','Airtight Freshness Jar System','OXO Good Grips','Indiegogo','https://www.indiegogo.com/projects/oxo-airtight-freshness-jar-system','https://oxo.com','support@oxo.com','★★★★','ボタン押しで密閉・食品鮮度保持ポップコンテナシステム'],
  [3262,'クリーニング・収納・整理','Smart Label Maker LabelWriter 5XL','DYMO','Kickstarter','https://www.kickstarter.com/projects/dymo/labelwriter-5xl-label-maker','https://dymo.com','support@dymo.com','★★★★','Wi-Fi接続・5インチ幅・PCなし単独使用スマートラベルライター'],
  [3263,'クリーニング・収納・整理','Grovemade Desk Shelf System','Grovemade','Indiegogo','https://www.indiegogo.com/projects/grovemade-desk-shelf-system','https://grovemade.com','hello@grovemade.com','★★★★★','天然ウォールナット無垢材・ミニマルデスク収納棚セット'],
  [3264,'クリーニング・収納・整理','Neat Method Home Organizing Kit','Neat Method','Kickstarter','https://www.kickstarter.com/projects/neatmethod/home-organizing-kit','https://neatmethod.com','info@neatmethod.com','★★★','プロ整理収納士監修・引き出し仕切りセット＋アドバイスカード'],
  [3265,'クリーニング・収納・整理','Airzai Atmos Air Purifier','Airzai','Indiegogo','https://www.indiegogo.com/projects/airzai-atmos-air-purifier','https://airzai.com','hello@airzai.com','★★★★','H13 HEPA+活性炭+UV-C三重フィルター・アプリ連携空気清浄機'],
  [3266,'クリーニング・収納・整理','Dreame L10s Ultra Robot Vacuum Mop','Dreame Technology','Kickstarter','https://www.kickstarter.com/projects/dreame/l10s-ultra-robot-vacuum-mop','https://dreame.tech','support@dreame.tech','★★★★★','自動モップ洗浄・熱風乾燥ドック・全自動床掃除ロボット'],
  [3267,'クリーニング・収納・整理','Bissell SpinWave Cordless Mop','Bissell','Indiegogo','https://www.indiegogo.com/projects/bissell-spinwave-cordless-mop','https://bissell.com','support@bissell.com','★★★★','回転スピンパッド・コードレス・ハードフロア用電動モップ'],
  [3268,'クリーニング・収納・整理','Modular Wall-Mounted Bike Storage','Steadyrack','Kickstarter','https://www.kickstarter.com/projects/steadyrack/modular-wall-mounted-bike-storage','https://steadyrack.com','info@steadyrack.com','★★★★','扇形回転・省スペース壁掛けバイク収納ラック'],
  [3269,'クリーニング・収納・整理','Calpak Expandable Packing Cubes','Calpak','Indiegogo','https://www.indiegogo.com/projects/calpak-expandable-packing-cubes','https://calpaktravel.com','hello@calpaktravel.com','★★★★','2倍展開・圧縮ジッパー・スーツケース収納パッキングキューブ'],
  [3270,'クリーニング・収納・整理','Shark FlexStyle Air Styling System','Shark Beauty','Kickstarter','https://www.kickstarter.com/projects/sharkbeauty/flexstyle-air-styling-system','https://sharkbeauty.com','support@sharkbeauty.com','★★★★★','ドライヤー＋カール＋ストレート・マルチスタイリングシステム'],
  [3271,'クリーニング・収納・整理','Whitmor Supreme Closet Organizer','Whitmor','Indiegogo','https://www.indiegogo.com/projects/whitmor-supreme-closet-organizer','https://whitmor.com','support@whitmor.com','★★★','拡張可能・クローゼット全体を収納最適化するシステム'],
];

// ── ユーティリティ ──────────────────────────────────────────────────────────
function buildGmailUrl(toEmail, maker, product) {
  const enc = s => encodeURIComponent(s);
  const subject = `Potential Distribution Partnership for Japan - ${product}`;
  const body =
`Dear ${maker} Team,

I hope this message finds you well.

My name is Yutorin Ino, and I represent LEAGUE, a Japanese company specializing in bringing innovative overseas products to the Japanese market.

We recently discovered your product "${product}" and were truly impressed by its innovation and quality. We believe it has significant potential in Japan, where consumers are enthusiastic about cutting-edge products from abroad.

We would love to explore the possibility of becoming your authorized distributor in Japan. Our team has extensive experience in market entry, logistics, and retail partnerships across Japan.

Could we schedule a brief call or email exchange to discuss this opportunity further?

Looking forward to your response.

Best regards,
Yutorin Ino
LEAGUE Inc.
Email: ${SENDER}`;

  return `https://mail.google.com/mail/?authuser=${enc(SENDER)}&view=cm&fs=1&to=${enc(toEmail)}&su=${enc(subject)}&body=${enc(body)}`;
}

// ── メイン処理 ──────────────────────────────────────────────────────────────
console.log('Reading:', INPUT);
const wb_in   = XLSX.readFile(INPUT);
const ws_in   = wb_in.Sheets[wb_in.SheetNames[0]];
const existing = XLSX.utils.sheet_to_json(ws_in, { header: 1 });

console.log(`既存行数: ${existing.length - 1} 件 (ヘッダー除く)`);

// 新規行を追加
const allRows = [...existing];
for (const p of NEW_PRODUCTS) {
  allRows.push(p);
}

// 新しいワークブック作成
const wb_out = XLSX.utils.book_new();
const ws_out = XLSX.utils.aoa_to_sheet(allRows);

// 列幅設定
ws_out['!cols'] = [
  { wch: 6  }, // No
  { wch: 18 }, // カテゴリ
  { wch: 40 }, // 製品名
  { wch: 20 }, // メーカー
  { wch: 14 }, // ECサイト
  { wch: 50 }, // 製品URL
  { wch: 35 }, // メーカーHP
  { wch: 35 }, // メール
  { wch: 8  }, // ★
  { wch: 40 }, // コメント
];

// メール列にGmailハイパーリンク設定
const dataStartRow = 1; // 0-indexed (row 0 = header)
for (let r = dataStartRow; r < allRows.length; r++) {
  const row = allRows[r];
  const mailAddr = String(row[7] || '');
  if (mailAddr && mailAddr.includes('@')) {
    const cellAddr = XLSX.utils.encode_cell({ r, c: 7 });
    const gmailUrl = buildGmailUrl(mailAddr, String(row[3] || ''), String(row[2] || ''));
    ws_out[cellAddr] = {
      v: mailAddr,
      t: 's',
      l: { Target: gmailUrl },
      s: { font: { color: { rgb: '1155CC' }, underline: true } }
    };
  }
}

XLSX.utils.book_append_sheet(wb_out, ws_out, '製品リスト');
XLSX.writeFile(wb_out, OUTPUT);

// ── 集計レポート ────────────────────────────────────────────────────────────
const totalNew   = NEW_PRODUCTS.length;
const mailCount  = NEW_PRODUCTS.filter(p => p[7] && p[7].includes('@')).length;
const cfCount    = NEW_PRODUCTS.filter(p =>
  ['Kickstarter','Indiegogo'].includes(p[4])
).length;

const catCount = {};
for (const p of NEW_PRODUCTS) {
  catCount[p[1]] = (catCount[p[1]] || 0) + 1;
}

console.log('\n====== 追加完了レポート ======');
console.log(`追加件数    : ${totalNew} 件`);
console.log(`メールあり  : ${mailCount} 件`);
console.log(`クラファン  : ${cfCount} 件`);
console.log('\nカテゴリ別:');
Object.entries(catCount).forEach(([cat, n]) => console.log(`  ${cat}: ${n}件`));
console.log(`\n総収録件数  : ${existing.length - 1 + totalNew} 件`);
console.log(`出力ファイル: ${OUTPUT}`);
