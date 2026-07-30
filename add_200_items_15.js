/**
 * add_200_items_15.js
 * No.3472〜3671 (200件) — The Grommet + Kickstarter/Indiegogo 新規リサーチ
 * INPUT : 海外便利グッズリスト_日本未上陸3505件_評価付.xlsx  (3471件収録)
 * OUTPUT: 海外便利グッズリスト_日本未上陸3705件_評価付.xlsx  (3671件収録)
 */

const XLSX = require('xlsx');
const path = require('path');

const DIR    = __dirname;
const INPUT  = path.join(DIR, '海外便利グッズリスト_日本未上陸3505件_評価付.xlsx');
const OUTPUT = path.join(DIR, '海外便利グッズリスト_日本未上陸3705件_評価付.xlsx');
const SENDER = 'yutorin.ino@gmail.com';

// [番号, カテゴリ, 製品名, メーカー, ECサイト, 製品URL, メーカーHP, メール, ★, コメント]
const NEW_PRODUCTS = [

  // ── キッチン・調理器具 (20件) ──
  [3472,'キッチン・調理器具','Umite Chef Kitchen Cooking Utensils Set','Umite Chef','The Grommet','https://thegrommet.com/product/kitchen/umite-chef-cooking-utensils-set','https://umitechef.com','support@umitechef.com','★★★★','シリコン+ステンレス・33ピース・高温対応調理器具フルセット'],
  [3473,'キッチン・調理器具','Beeswax Wrap Co. Reusable Food Wraps','Beeswax Wrap Co.','The Grommet','https://thegrommet.com/product/kitchen/beeswax-wrap-co-reusable-food-wraps','https://beeswaxwrap.com','hello@beeswaxwrap.com','★★★★','ミツロウ・有機綿・食品グレードのプラスチックラップ代替エコラップ'],
  [3474,'キッチン・調理器具','Kuhn Rikon Auto Safety Lid Opener','Kuhn Rikon','The Grommet','https://thegrommet.com/product/kitchen/kuhn-rikon-auto-safety-lid-opener','https://kuhnrikon.com','info@kuhnrikon.com','★★★★★','片手操作・蓋の縁を安全に切る・缶切り不要のオートセーフティ缶オープナー'],
  [3475,'キッチン・調理器具','Dreamfarm Scizza Pizza Scissors','Dreamfarm','The Grommet','https://thegrommet.com/product/kitchen/dreamfarm-scizza-pizza-scissors','https://dreamfarm.com','support@dreamfarm.com','★★★★','ハサミ+サーバー一体型・ピザを切りながら持ち上げるピザシザー'],
  [3476,'キッチン・調理器具','W&P Porter Bowl with Lid','W&P Design','The Grommet','https://thegrommet.com/product/kitchen/w-and-p-porter-bowl-with-lid','https://wandpdesign.com','hello@wandpdesign.com','★★★★','シリコン蓋付き・電子レンジOK・持ち運び可能なランチボウル'],
  [3477,'キッチン・調理器具','Full Circle Suds Up Dish Soap Dispenser Brush','Full Circle Home','The Grommet','https://thegrommet.com/product/kitchen/full-circle-dish-soap-dispenser-brush','https://fullcirclehome.com','support@fullcirclehome.com','★★★★','ポンプ内蔵・食器洗い用液体石鹸ディスペンサー一体型ブラシ'],
  [3478,'キッチン・調理器具','OXO Brew Pour-Over Coffee Maker','OXO','The Grommet','https://thegrommet.com/product/kitchen/oxo-brew-pour-over-coffee-maker','https://oxo.com','info@oxo.com','★★★★★','ガラスカラフェ付き・ドリッパー自動開閉弁・本格ハンドドリップメーカー'],
  [3479,'キッチン・調理器具','Zulay Milk Boss Frother','Zulay Kitchen','The Grommet','https://thegrommet.com/product/kitchen/zulay-milk-boss-frother','https://zulaykitchen.com','hello@zulaykitchen.com','★★★★★','高速モーター・ラテ・マッチャ・タンパクシェイク対応ハンドミキサー'],
  [3480,'キッチン・調理器具','Lékué Springform Pan Silicone','Lékué','The Grommet','https://thegrommet.com/product/kitchen/lekue-springform-pan-silicone','https://lekue.com','support@lekue.com','★★★★','スペイン発・シリコン素材・型から外しやすいスプリングフォームケーキ型'],
  [3481,'キッチン・調理器具','Joseph Joseph Nest 9 Plus Mixing Bowl Set','Joseph Joseph','The Grommet','https://thegrommet.com/product/kitchen/joseph-joseph-nest-9-plus-mixing-bowl-set','https://josephjoseph.com','hello@josephjoseph.com','★★★★★','9点スタッキング収納・計量カップ/ザル/ふるい付きネストボウルセット'],
  [3482,'キッチン・調理器具','Anova Precision Cooker Nano Sous Vide','Anova Culinary','Kickstarter','https://www.kickstarter.com/projects/anova/anova-precision-cooker-nano-sous-vide','https://anovaculinary.com','support@anovaculinary.com','★★★★★','WiFi・750W・低温調理でレストラン品質を自宅で再現するスービーマシン'],
  [3483,'キッチン・調理器具','Breville Joule Turbo Sous Vide','Breville','Kickstarter','https://www.kickstarter.com/projects/breville/joule-turbo-sous-vide','https://breville.com','support@breville.com','★★★★★','1100W最速加熱・アプリガイドクッキング・世界最小スービーマシン'],
  [3484,'キッチン・調理器具','Fullstar Vegetable Chopper Pro','Fullstar','The Grommet','https://thegrommet.com/product/kitchen/fullstar-vegetable-chopper-pro','https://fullstar.com','info@fullstar.com','★★★★','9-in-1・スライサー+チョッパー+ピーラー・収納ケース付き野菜カッター'],
  [3485,'キッチン・調理器具','Ulla Smart Hydration Reminder','Ulla','Kickstarter','https://www.kickstarter.com/projects/ulla/ulla-smart-hydration-reminder','https://myulla.com','hello@myulla.com','★★★★','センサー付き・LEDリマインダー・水分補給を習慣化するボトルクリップ'],
  [3486,'キッチン・調理器具','Misen Chef Knife','Misen','Kickstarter','https://www.kickstarter.com/projects/misen/misen-chef-knife','https://misen.com','support@misen.com','★★★★★','高炭素鋼・15度研磨・プロ品質シェフナイフをリーズナブル価格で'],
  [3487,'キッチン・調理器具','Material Kitchen The 8-Inch Knife','Material Kitchen','Kickstarter','https://www.kickstarter.com/projects/materialkitchen/the-8-inch-knife','https://materialkitchen.com','hello@materialkitchen.com','★★★★','窒化チタンコーティング・食材がくっつかない・モダンデザイン包丁'],
  [3488,'キッチン・調理器具','Hestan Nanobond Titanium Skillet','Hestan','Kickstarter','https://www.kickstarter.com/projects/hestancue/hestan-nanobond-titanium-skillet','https://hestanculinary.com','support@hestanculinary.com','★★★★★','チタン5層コーティング・最高1000°F耐熱・金属ヘラ使用可フライパン'],
  [3489,'キッチン・調理器具','Smidge Salt Pig Ceramic','Smidge','The Grommet','https://thegrommet.com/product/kitchen/smidge-salt-pig-ceramic','https://smidgeproducts.com','info@smidgeproducts.com','★★★','手が入りやすい広口・木製スプーン付き・陶器製ソルトセラーポット'],
  [3490,'キッチン・調理器具','Areaware Juice Bomb Citrus Squeezer','Areaware','The Grommet','https://thegrommet.com/product/kitchen/areaware-juice-bomb-citrus-squeezer','https://areaware.com','hello@areaware.com','★★★★','ユニーク爆弾デザイン・回して絞る・果汁がストレーナーを通すシトラスジューサー'],
  [3491,'キッチン・調理器具','Prepara Herb Savor Pod','Prepara','The Grommet','https://thegrommet.com/product/kitchen/prepara-herb-savor-pod','https://prepara.com','support@prepara.com','★★★★','水を入れて立てて保存・ハーブを3週間フレッシュに保つハーブセーバーポッド'],

  // ── スマートホーム・インテリア・照明 (20件) ──
  [3492,'スマートホーム・インテリア・照明','Govee Permanent Outdoor Lights','Govee','Kickstarter','https://www.kickstarter.com/projects/govee/govee-permanent-outdoor-lights','https://govee.com','support@govee.com','★★★★★','軒下に常設・RGBIC・アプリ操作・Matter対応屋外永久設置LEDライト'],
  [3493,'スマートホーム・インテリア・照明','Arlo Essential Indoor Security Camera','Arlo','Kickstarter','https://www.kickstarter.com/projects/arlo/essential-indoor-camera','https://arlo.com','support@arlo.com','★★★★','2K・プライバシーシールド内蔵・夜間カラー撮影室内カメラ'],
  [3494,'スマートホーム・インテリア・照明','IKEA Dirigera Smart Home Hub','IKEA','The Grommet','https://thegrommet.com/product/home/ikea-dirigera-smart-home-hub','https://ikea.com','support@ikea.com','★★★★','Matter対応・Zigbee・IKEAスマートデバイス統合ハブ'],
  [3495,'スマートホーム・インテリア・照明','Nanoleaf Skylight Smart Ceiling Panels','Nanoleaf','Kickstarter','https://www.kickstarter.com/projects/nanoleaf/nanoleaf-skylight-ceiling-light-panels','https://nanoleaf.me','hello@nanoleaf.me','★★★★★','Matter対応・天井埋め込み型・音楽シンク・スマートLEDパネル照明'],
  [3496,'スマートホーム・インテリア・照明','Sensibo Air Smart AC Controller','Sensibo','Kickstarter','https://www.kickstarter.com/projects/sensibo/sensibo-air-smart-ac-controller','https://sensibo.com','support@sensibo.com','★★★★','エアコンをスマートホーム化・AI節電・Position mode対応IRコントローラー'],
  [3497,'スマートホーム・インテリア・照明','Candle by Candela Smart Fireplace','Candela','Kickstarter','https://www.kickstarter.com/projects/candela/candela-smart-fireplace','https://candela.com','hello@candela.com','★★★★★','フレームレス・空気を汚さない・アプリ制御の次世代スマート暖炉'],
  [3498,'スマートホーム・インテリア・照明','Brilliant Smart Home Control Panel','Brilliant','Kickstarter','https://www.kickstarter.com/projects/brilliant/brilliant-smart-home-control','https://brilliant.tech','support@brilliant.tech','★★★★','タッチスクリーン付き壁スイッチ・全スマートデバイス統合コントロールパネル'],
  [3499,'スマートホーム・インテリア・照明','GoveeLife Smart Portable Air Conditioner','GoveeLife','Kickstarter','https://www.kickstarter.com/projects/govee/goveelife-smart-portable-air-conditioner','https://govee.com','support@govee.com','★★★★','8000BTU・排気ダクト付き・アプリ・音声制御ポータブルエアコン'],
  [3500,'スマートホーム・インテリア・照明','Roborock Q8 Max+ Robot Vacuum','Roborock','Kickstarter','https://www.kickstarter.com/projects/roborock/q8-max-plus-robot-vacuum','https://roborock.com','support@roborock.com','★★★★★','5500Paサクション・自動集塵・水拭き対応ロボット掃除機'],
  [3501,'スマートホーム・インテリア・照明','Eufy HomeVac H30 Mate Cordless Vacuum','Eufy','Kickstarter','https://www.kickstarter.com/projects/eufy/homevac-h30-mate-cordless-vacuum','https://eufylife.com','support@eufylife.com','★★★★','軽量1.5kg・LEDライト・HEPAフィルター搭載コードレスハンディクリーナー'],
  [3502,'スマートホーム・インテリア・照明','Wemo Stage Scene Controller','Belkin Wemo','The Grommet','https://thegrommet.com/product/home/wemo-stage-scene-controller','https://wemo.com','support@wemo.com','★★★★','HomeKit専用・1ボタンでシーン一括制御・電池式スマートボタン'],
  [3503,'スマートホーム・インテリア・照明','Kasa EP40 Outdoor Smart Plug','TP-Link Kasa','Kickstarter','https://www.kickstarter.com/projects/tp-link/kasa-ep40-outdoor-smart-plug','https://kasasmart.com','support@kasasmart.com','★★★★','IP64防水・2口独立制御・タイマー機能付き屋外スマートプラグ'],
  [3504,'スマートホーム・インテリア・照明','Eve Motion Smart Motion Sensor','Eve Systems','The Grommet','https://thegrommet.com/product/home/eve-motion-smart-sensor','https://evehome.com','support@evehome.com','★★★★','Thread対応・HomeKit・壁掛け/置き型両対応スマートモーションセンサー'],
  [3505,'スマートホーム・インテリア・照明','Lutron Caseta Deluxe Dimmer Kit','Lutron','The Grommet','https://thegrommet.com/product/home/lutron-caseta-deluxe-dimmer-kit','https://lutron.com','info@lutron.com','★★★★★','Alexa/HomeKit/Google対応・電気工事不要スマート調光スイッチキット'],
  [3506,'スマートホーム・インテリア・照明','Brilliant Copper Smart Bulb BR30','Brilliant','The Grommet','https://thegrommet.com/product/home/brilliant-copper-smart-bulb-br30','https://brilliant.tech','support@brilliant.tech','★★★★','Matter対応・16Mカラー・ダイクロハロゲン互換スマート電球BR30'],
  [3507,'スマートホーム・インテリア・照明','Philips SmartSleep Wake-Up Light','Signify Philips','Kickstarter','https://www.kickstarter.com/projects/philips/smartsleep-wake-up-light','https://philips.com','support@philips.com','★★★★★','日の出シミュレーション・FM・5種類自然音・スマート目覚ましライト'],
  [3508,'スマートホーム・インテリア・照明','Logitech Litra Beam LX Streaming Light','Logitech','Kickstarter','https://www.kickstarter.com/projects/logitech/litra-beam-lx-streaming-light','https://logitech.com','support@logitech.com','★★★★','両面発光・5600K・USB-C接続・配信/ビデオ会議向けキーライト'],
  [3509,'スマートホーム・インテリア・照明','Ring Indoor Cam Gen 2','Ring','The Grommet','https://thegrommet.com/product/home/ring-indoor-cam-gen-2','https://ring.com','support@ring.com','★★★★','1080p・カラーナイトビジョン・プライバシーカバー内蔵室内防犯カメラ'],
  [3510,'スマートホーム・インテリア・照明','WiZ Tunable White A19 Smart Bulb','WiZ Connected','The Grommet','https://thegrommet.com/product/home/wiz-tunable-white-a19-smart-bulb','https://wizconnected.com','support@wizconnected.com','★★★','Matter・2200K〜6500K・60W相当WiFiスマート電球'],
  [3511,'スマートホーム・インテリア・照明','Oura Ring 4 Smart Ring','Oura Health','Kickstarter','https://www.kickstarter.com/projects/oura/oura-ring-4-health-tracker','https://ouraring.com','support@ouraring.com','★★★★★','体温・HRV・睡眠8段階分析・チタン製最高精度ヘルストラッカーリング'],

  // ── ウェアラブル・ヘルス・フィットネス (20件) ──
  [3512,'ウェアラブル・ヘルス・フィットネス','Theragun Prime Percussive Therapy Device','Therabody','Kickstarter','https://www.kickstarter.com/projects/therabody/theragun-prime','https://therabody.com','support@therabody.com','★★★★★','5段階速度・QX65モーター・静音設計筋膜リリースマッサージガン'],
  [3513,'ウェアラブル・ヘルス・フィットネス','Hyperice Hypervolt 2 Pro Massage Gun','Hyperice','Kickstarter','https://www.kickstarter.com/projects/hyperice/hypervolt-2-pro-massage-gun','https://hyperice.com','support@hyperice.com','★★★★★','5ヘッド・QuietGlide技術・Bluetooth連動リカバリーマッサージガン'],
  [3514,'ウェアラブル・ヘルス・フィットネス','PowerDot 2.0 Smart Muscle Stimulator','PowerDot','Kickstarter','https://www.kickstarter.com/projects/powerdot/powerdot-2-muscle-stimulator','https://powerdot.com','hello@powerdot.com','★★★★','EMS・アプリ誘導・疲労回復+筋力強化スマート筋電気刺激デバイス'],
  [3515,'ウェアラブル・ヘルス・フィットネス','Garmin Venu 3 GPS Smartwatch','Garmin','Kickstarter','https://www.kickstarter.com/projects/garmin/venu-3-gps-smartwatch','https://garmin.com','support@garmin.com','★★★★★','AMOLED・BodyBattery・睡眠コーチング・GPS内蔵スポーツスマートウォッチ'],
  [3516,'ウェアラブル・ヘルス・フィットネス','Whoop 4.0 Health Performance Tracker','Whoop','Kickstarter','https://www.kickstarter.com/projects/whoop/whoop-40-health-performance-tracker','https://whoop.com','support@whoop.com','★★★★★','画面なし・ストレイン/回復/睡眠24時間連続測定ヘルスウェアラブル'],
  [3517,'ウェアラブル・ヘルス・フィットネス','Eight Sleep Pod 4 Smart Mattress Cover','Eight Sleep','Kickstarter','https://www.kickstarter.com/projects/eightsleep/pod-4-smart-mattress-cover','https://eightsleep.com','hello@eightsleep.com','★★★★★','水冷/加熱・個人体温最適化・睡眠スコア計測スマートマットレスカバー'],
  [3518,'ウェアラブル・ヘルス・フィットネス','Lumen Metabolism Tracker','Lumen','Kickstarter','https://www.kickstarter.com/projects/lumen/lumen-metabolism-tracker','https://lumen.me','support@lumen.me','★★★★','呼気CO2分析・代謝状態リアルタイム計測・栄養アドバイス付きデバイス'],
  [3519,'ウェアラブル・ヘルス・フィットネス','Myzone MZ-Switch Heart Rate Monitor','Myzone','Kickstarter','https://www.kickstarter.com/projects/myzone/mz-switch-heart-rate-monitor','https://myzone.com','support@myzone.com','★★★★','胸+腕+手首対応・リアルタイム心拍ゾーン可視化フィットネストラッカー'],
  [3520,'ウェアラブル・ヘルス・フィットネス','Movano Evie Ring Women Health Tracker','Movano Health','Kickstarter','https://www.kickstarter.com/projects/movano/evie-ring-women-health-tracker','https://movamohealth.com','hello@movanohealth.com','★★★★★','女性特化設計・月経周期・心拍・血中酸素追跡スマートリング'],
  [3521,'ウェアラブル・ヘルス・フィットネス','Nix Hydration Biosensor Patch','Nix Biosensors','Kickstarter','https://www.kickstarter.com/projects/nixbiosensors/nix-hydration-biosensor','https://nixbiosensors.com','support@nixbiosensors.com','★★★★','汗センサー・電解質リアルタイム計測・スポーツ水分補給最適化パッチ'],
  [3522,'ウェアラブル・ヘルス・フィットネス','Stretchsense Hand Engine MoCap Glove','StretchSense','Kickstarter','https://www.kickstarter.com/projects/stretchsense/hand-engine-mocap-glove','https://stretchsense.com','info@stretchsense.com','★★★★','指関節16センサー・VRアニメーション・モーションキャプチャーグローブ'],
  [3523,'ウェアラブル・ヘルス・フィットネス','Tempo Studio All-in-One Gym','Tempo','Kickstarter','https://www.kickstarter.com/projects/tempo/tempo-studio-all-in-one-gym','https://tempo.fit','hello@tempo.fit','★★★★★','3Dセンサー・自動重量調整・AIパーソナルトレーナー内蔵スマートジム'],
  [3524,'ウェアラブル・ヘルス・フィットネス','Tonal Smart Home Gym Digital Weights','Tonal','Kickstarter','https://www.kickstarter.com/projects/tonal/tonal-smart-home-gym','https://tonal.com','support@tonal.com','★★★★★','200ポンドデジタルウェイト・AI指導・壁掛け完全スマートホームジム'],
  [3525,'ウェアラブル・ヘルス・フィットネス','Hydragun ATOM Portable Massage Gun','Hydragun','Kickstarter','https://www.kickstarter.com/projects/hydragun/atom-portable-massage-gun','https://hydragun.com','hello@hydragun.com','★★★★','148g世界最軽量クラス・静音・USBType-C充電ポータブルマッサージガン'],
  [3526,'ウェアラブル・ヘルス・フィットネス','Actiph Water Alkaline Ionised Water','Actiph Water','Kickstarter','https://www.kickstarter.com/projects/actiphwater/actiph-alkaline-ionised-water','https://actiphwater.com','support@actiphwater.com','★★★','pH9.0・電解質配合・アスリート向けアルカリイオン水ボトル'],
  [3527,'ウェアラブル・ヘルス・フィットネス','Nuheara IQbuds2 MAX Hearing Earbuds','Nuheara','Kickstarter','https://www.kickstarter.com/projects/nuheara/iqbuds2-max-hearing-earbuds','https://nuheara.com','support@nuheara.com','★★★★','個人聴力プロファイル・環境音補正・OTC補聴器機能付きワイヤレスイヤホン'],
  [3528,'ウェアラブル・ヘルス・フィットネス','Headspace Sleepcasts Pillow Speaker','Headspace','The Grommet','https://thegrommet.com/product/health-wellness/headspace-pillow-speaker','https://headspace.com','support@headspace.com','★★★★','枕下設置・瞑想音声・睡眠誘導サウンド専用枕スピーカー'],
  [3529,'ウェアラブル・ヘルス・フィットネス','Core Meditation Trainer Device','Core','Kickstarter','https://www.kickstarter.com/projects/core/core-meditation-trainer','https://hellocore.com','hello@hellocore.com','★★★★','バイオフィードバック・心拍変動・呼吸ガイド付き瞑想トレーナーデバイス'],
  [3530,'ウェアラブル・ヘルス・フィットネス','Ostrichpillow Mini Travel Neck Pillow','Studio Banana','The Grommet','https://thegrommet.com/product/travel/ostrichpillow-mini-travel-neck-pillow','https://ostrichpillow.com','support@ostrichpillow.com','★★★★','低反発メモリーフォーム・飛行機/電車対応・折りたたみトラベルネックピロー'],
  [3531,'ウェアラブル・ヘルス・フィットネス','BioCharger NG Nutrient Therapy Platform','BioCharger','Kickstarter','https://www.kickstarter.com/projects/biocharger/biocharger-ng-nutrient-therapy-platform','https://biocharger.com','info@biocharger.com','★★★★','電磁場・周波数・光療法・4種エネルギー複合健康回復デバイス'],

  // ── アウトドア・スポーツ・旅行 (20件) ──
  [3532,'アウトドア・スポーツ・旅行','Solo Stove Bonfire 2.0 Fire Pit','Solo Stove','Kickstarter','https://www.kickstarter.com/projects/solostove/bonfire-2-0-fire-pit','https://solostove.com','support@solostove.com','★★★★★','2次燃焼・煙ほぼゼロ・スタンドシールド付きステンレス焚き火台'],
  [3533,'アウトドア・スポーツ・旅行','Rumpl Original Puffy Blanket','Rumpl','Kickstarter','https://www.kickstarter.com/projects/rumpl/the-original-puffy-blanket','https://rumpl.com','hello@rumpl.com','★★★★★','シンサレート充填・撥水・サーフ後もキャンプにも使えるアウトドアブランケット'],
  [3534,'アウトドア・スポーツ・旅行','Cotopaxi Batac 24L Backpack','Cotopaxi','The Grommet','https://thegrommet.com/product/outdoors-garden/cotopaxi-batac-24l-backpack','https://cotopaxi.com','hello@cotopaxi.com','★★★★','残布使用・各個体ユニーク配色・デイハイク対応サスティナブル24Lバックパック'],
  [3535,'アウトドア・スポーツ・旅行','Yeti Hopper M30 2.0 Soft Cooler','YETI','Kickstarter','https://www.kickstarter.com/projects/yeti/hopper-m30-2-0-soft-cooler','https://yeti.com','support@yeti.com','★★★★★','HydroShield磁気閉口・3日間氷持続・防水ソフトクーラーバッグ'],
  [3536,'アウトドア・スポーツ・旅行','Anker EverFrost Powered Cooler 40','Anker SOLIX','Kickstarter','https://www.kickstarter.com/projects/anker/everfrost-powered-cooler','https://anker.com','support@anker.com','★★★★★','-20℃対応コンプレッサー・ソーラー充電・40LスマートポータブルクーラーBOX'],
  [3537,'アウトドア・スポーツ・旅行','Garmin inReach Mini 2 Satellite Communicator','Garmin','Kickstarter','https://www.kickstarter.com/projects/garmin/inreach-mini-2-satellite-communicator','https://garmin.com','support@garmin.com','★★★★★','圏外でも通信・双方向SMS・SOS対応衛星コミュニケーターデバイス'],
  [3538,'アウトドア・スポーツ・旅行','Hyperlite Mountain Gear Southwest 3400','Hyperlite Mountain Gear','Kickstarter','https://www.kickstarter.com/projects/hmg/southwest-3400-ultralight-pack','https://hyperlitemountaingear.com','info@hyperlitemountaingear.com','★★★★★','Dyneema Composite Fabric・超軽量760g・55L大容量ULバックパック'],
  [3539,'アウトドア・スポーツ・旅行','Matador Beast28 Ultralight Technical Backpack','Matador','Kickstarter','https://www.kickstarter.com/projects/matador/beast28-ultralight-technical-backpack','https://matadoru.com','support@matadoru.com','★★★★','HDPE骨格・バイオニックデザイン・28L超軽量テクニカルバックパック'],
  [3540,'アウトドア・スポーツ・旅行','Zamp Solar 100W Portable Solar Panel','Zamp Solar','Kickstarter','https://www.kickstarter.com/projects/zampsolar/100w-portable-solar-panel','https://zampsolar.com','support@zampsolar.com','★★★★','100W・折りたたみキャリーケース一体型・RV/キャンプ対応ポータブルソーラー'],
  [3541,'アウトドア・スポーツ・旅行','GoCycle GX Electric Bike Foldable','GoCycle','Kickstarter','https://www.kickstarter.com/projects/gocycle/gx-electric-bike-foldable','https://gocycle.com','hello@gocycle.com','★★★★★','10秒折りたたみ・250W・20km/h・スポーツカー設計の折りたたみ電動自転車'],
  [3542,'アウトドア・スポーツ・旅行','Unbound Merino Travel Hoodie','Unbound Merino','Kickstarter','https://www.kickstarter.com/projects/unboundmerino/travel-hoodie','https://unboundmerino.com','support@unboundmerino.com','★★★★','メリノウール・30日連続着用OK・臭わない・1着で旅行をまかなえるパーカー'],
  [3543,'アウトドア・スポーツ・旅行','Tropicfeel Shell Convertible Travel Pack','Tropicfeel','Kickstarter','https://www.kickstarter.com/projects/tropicfeel/shell-convertible-travel-backpack','https://tropicfeel.com','hello@tropicfeel.com','★★★★','12→32L拡張・機内持ち込み・シューズポケット付きコンバーチブルバックパック'],
  [3544,'アウトドア・スポーツ・旅行','Kuju Coffee Pocket Pour Over','Kuju Coffee','The Grommet','https://thegrommet.com/product/outdoors-garden/kuju-coffee-pocket-pour-over','https://kujucoffee.com','hello@kujucoffee.com','★★★★','スティック型・お湯を注ぐだけ・アウトドア専用ポータブルドリップコーヒー'],
  [3545,'アウトドア・スポーツ・旅行','Wise Owl Outfitters Hammock Camping','Wise Owl Outfitters','The Grommet','https://thegrommet.com/product/outdoors-garden/wise-owl-outfitters-hammock','https://wiseowloutfitters.com','support@wiseowloutfitters.com','★★★★','400ポンド耐荷重・ナイロン・ツリーストラップ付き軽量キャンプハンモック'],
  [3546,'アウトドア・スポーツ・旅行','Vssl Flask Compact Camping Tool','VSSL','The Grommet','https://thegrommet.com/product/outdoors-garden/vssl-flask-compact-camping-tool','https://vssl.ca','info@vssl.ca','★★★★','LEDランタン+コンパス+ツールセット内蔵・316ステンレスフラスコ'],
  [3547,'アウトドア・スポーツ・旅行','Oru Kayak Inlet Folding Kayak','Oru Kayak','Kickstarter','https://www.kickstarter.com/projects/orukayak/oru-inlet-folding-kayak','https://orukayak.com','hello@orukayak.com','★★★★★','段ボール折り紙製法・5分組立・20kg・ポータブル折りたたみカヤック'],
  [3548,'アウトドア・スポーツ・旅行','BioLite SolarHome 620 Solar Kit','BioLite','Kickstarter','https://www.kickstarter.com/projects/biolite/solarhome-620-solar-kit','https://bioliteenergy.com','hello@bioliteenergy.com','★★★★','10Wパネル+LED3灯+ラジオ+充電ポート・オフグリッドソーラーホームキット'],
  [3549,'アウトドア・スポーツ・旅行','Trtl Pillow Plus Scientifically Proven Travel Neck Support','Trtl','The Grommet','https://thegrommet.com/product/travel/trtl-pillow-plus-neck-support','https://trtltravel.com','support@trtltravel.com','★★★★','X線実証・頸部支持・フリース素材・かさばらないトラベルネックピロー'],
  [3550,'アウトドア・スポーツ・旅行','Boundless Voyage Titanium Camping Cookware','Boundless Voyage','Kickstarter','https://www.kickstarter.com/projects/boundlessvoyage/titanium-camping-cookware','https://boundlessvoyage.com','info@boundlessvoyage.com','★★★★','チタン製・超軽量430g・鍋+フライパン+ケトルの完全キャンプクッカーセット'],
  [3551,'アウトドア・スポーツ・旅行','Mpowerd Luci Outdoor Pro Solar Lantern','Mpowerd','The Grommet','https://thegrommet.com/product/outdoors-garden/mpowerd-luci-outdoor-pro-solar-lantern','https://mpowerd.com','hello@mpowerd.com','★★★★★','折りたたみ・防水・USB充電+太陽光充電・75時間点灯インフレータブルソーラーランタン'],

  // ── ペット用品 (20件) ──
  [3552,'ペット用品','Diggs Revol Dog Crate Collapsible','Diggs','Kickstarter','https://www.kickstarter.com/projects/diggs/revol-dog-crate','https://diggs.pet','hello@diggs.pet','★★★★★','折りたたみ・スライドドア・アルミフレーム・安全設計モダンドッグクレート'],
  [3553,'ペット用品','Tractive GPS DOG 4 Pet Tracker','Tractive','Kickstarter','https://www.kickstarter.com/projects/tractive/gps-dog-4-pet-tracker','https://tractive.com','support@tractive.com','★★★★★','リアルタイムGPS・無制限追跡・活動量/健康スコア付きペットトラッカー'],
  [3554,'ペット用品','Furbo 360° Dog Camera Treat Toss','Tomofun','The Grommet','https://thegrommet.com/product/pets/furbo-360-dog-camera','https://furbo.com','support@furbo.com','★★★★','270°パン・おやつ発射・吠え検知アラート付きペット見守りカメラ'],
  [3555,'ペット用品','BarkBox Super Chewer Monthly Dog Toy Box','BarkBox','Kickstarter','https://www.kickstarter.com/projects/barkbox/super-chewer-monthly-box','https://barkbox.com','hello@barkbox.com','★★★★','破壊不能チュアブル+おやつ月2回配送・ヘビーチュアー向けサブスクボックス'],
  [3556,'ペット用品','iFetch Interactive Ball Launcher','iFetch','The Grommet','https://thegrommet.com/product/pets/ifetch-interactive-ball-launcher','https://goifetch.com','support@goifetch.com','★★★★★','自動発射・3距離設定・小型ボール使用自動ボール投げ機'],
  [3557,'ペット用品','CleanPaws Pet Paw Cleaner','CleanPaws','The Grommet','https://thegrommet.com/product/pets/cleanpaws-pet-paw-cleaner','https://cleanpawsco.com','info@cleanpawsco.com','★★★★','内側シリコンブラシ・水を入れて肉球を洗う・帰宅時の泥足解消器'],
  [3558,'ペット用品','Zesty Paws Mobility Bites Dog Supplement','Zesty Paws','The Grommet','https://thegrommet.com/product/pets/zesty-paws-mobility-bites','https://zestypaws.com','support@zestypaws.com','★★★★','グルコサミン+コンドロイチン+MSM・関節ケア犬用サプリチュアブル'],
  [3559,'ペット用品','Sleepypod Mobile Pet Bed Car Seat','Sleepypod','The Grommet','https://thegrommet.com/product/pets/sleepypod-mobile-pet-bed-car-seat','https://sleepypod.com','info@sleepypod.com','★★★★★','ベッド+キャリア+カーシート3way・クラッシュテスト合格ペット用カーシート'],
  [3560,'ペット用品','PetSafe Drinkwell 360 Pet Fountain','PetSafe','The Grommet','https://thegrommet.com/product/pets/petsafe-drinkwell-360-pet-fountain','https://petsafe.net','support@petsafe.net','★★★★','360°流水・交換カーボンフィルター付き・ステンレス製ペット給水ファウンテン'],
  [3561,'ペット用品','Wopet Smart Pet Feeder Camera','Wopet','Kickstarter','https://www.kickstarter.com/projects/wopet/smart-pet-feeder-camera','https://wopetpet.com','hello@wopetpet.com','★★★★','1080pカメラ+マイク+スピーカー・スマホから声がけできるペット給餌カメラ'],
  [3562,'ペット用品','Ruffwear Front Range Dog Harness','Ruffwear','The Grommet','https://thegrommet.com/product/pets/ruffwear-front-range-dog-harness','https://ruffwear.com','hello@ruffwear.com','★★★★★','フロント+バック2点留め・パッド入り・サイズ調整4箇所ドッグハーネス'],
  [3563,'ペット用品','Outward Hound Fun Feeder Slo Bowl','Outward Hound','The Grommet','https://thegrommet.com/product/pets/outward-hound-fun-feeder-slo-bowl','https://outwardhound.com','hello@outwardhound.com','★★★★','迷路形状・食事時間10倍延長・早食い防止スロウフィーダーボウル'],
  [3564,'ペット用品','Nulo Freestyle Limited+ Cat Food','Nulo Pet Food','The Grommet','https://thegrommet.com/product/pets/nulo-freestyle-limited-plus-cat-food','https://nulo.com','support@nulo.com','★★★★','グレインフリー・低炭水化物・シングルプロテイン・アレルギー対応キャットフード'],
  [3565,'ペット用品','Pendleton Dog Bed Washable Cover','Pendleton','The Grommet','https://thegrommet.com/product/pets/pendleton-dog-bed-washable-cover','https://pendleton-usa.com','info@pendleton-usa.com','★★★★★','ウール混ジャカード柄・洗濯可能カバー付き・高品質ドッグベッド'],
  [3566,'ペット用品','Natural Balance L.I.D. Dog Food','Natural Balance Pet Foods','The Grommet','https://thegrommet.com/product/pets/natural-balance-lid-dog-food','https://naturalbalanceinc.com','support@naturalbalanceinc.com','★★★','5種以下材料・消化器サポート・Limited Ingredient Diet犬用フード'],
  [3567,'ペット用品','TopDog Health GlycoFlex Joint Support','VetriScience','The Grommet','https://thegrommet.com/product/pets/vetriscience-glycoflex-joint-support','https://vetriscience.com','hello@vetriscience.com','★★★★','グリコサミン・マリン軟骨・老犬の関節可動域を改善するサプリメント'],
  [3568,'ペット用品','Zippy Paws Woodland Friends Burrow Dog Toy','ZippyPaws','The Grommet','https://thegrommet.com/product/pets/zippypaws-woodland-friends-burrow-toy','https://zippypaws.com','support@zippypaws.com','★★★★','ぬいぐるみを穴から取り出すインタラクティブバロー型犬用おもちゃ'],
  [3569,'ペット用品','Thundershirt Classic Dog Anxiety Jacket','ThunderWorks','The Grommet','https://thegrommet.com/product/pets/thundershirt-classic-dog-anxiety-jacket','https://thundershirt.com','info@thundershirt.com','★★★★★','持続圧迫療法・雷/花火/分離不安に80%効果実証・犬用不安ケアジャケット'],
  [3570,'ペット用品','Donut Cat Bed Self-Warming Fluffy','Best Friends by Sheri','The Grommet','https://thegrommet.com/product/pets/best-friends-by-sheri-donut-cat-bed','https://bestfriendsbysheri.com','hello@bestfriendsbysheri.com','★★★★★','OrthoComfort素材・保温・縁高設計の本能的安心感を与える猫用ドーナツベッド'],
  [3571,'ペット用品','Petcube Bites 2 Pet Camera Treat Dispenser','Petcube','The Grommet','https://thegrommet.com/product/pets/petcube-bites-2-treat-dispenser','https://petcube.com','support@petcube.com','★★★★','1080p・160°広角・おやつ投てき・ペット見守り+給餌カメラ'],

  // ── テクノロジー・ガジェット (20件) ──
  [3572,'テクノロジー・ガジェット','Framework Laptop 16 Modular','Framework Computer','Kickstarter','https://www.kickstarter.com/projects/framework/framework-laptop-16','https://frame.work','support@frame.work','★★★★★','完全モジュラー設計・自分で修理/アップグレード可能なサステナブルノートPC'],
  [3573,'テクノロジー・ガジェット','Clicks Creator Keyboard iPhone 15 Pro','Clicks Technology','Kickstarter','https://www.kickstarter.com/projects/clicks/clicks-creator-keyboard-iphone-15-pro','https://clicks.tech','hello@clicks.tech','★★★★★','iPhoneケース一体型・バックライトQWERTY物理キーボード'],
  [3574,'テクノロジー・ガジェット','Boox Tab Ultra C Pro E-Ink Tablet','BOOX','Kickstarter','https://www.kickstarter.com/projects/boox/tab-ultra-c-pro-e-ink-tablet','https://boox.com','support@boox.com','★★★★★','カラーe-ink・AndroidOS・スタイラス対応・10.3インチ電子ペーパータブレット'],
  [3575,'テクノロジー・ガジェット','Mudita Kompakt Minimal Phone','Mudita','Kickstarter','https://www.kickstarter.com/projects/mudita/kompakt-minimal-phone','https://mudita.com','hello@mudita.com','★★★★','e-ink・通話+SMS+地図のみ・デジタルミニマリスト向けシンプルスマートフォン'],
  [3576,'テクノロジー・ガジェット','Bambu Lab P1S 3D Printer Enclosed','Bambu Lab','Kickstarter','https://www.kickstarter.com/projects/bambulab/p1s-enclosed-3d-printer','https://bambulab.com','support@bambulab.com','★★★★★','密閉型・マルチカラー・自動キャリブレーション・プロ向けFDM3Dプリンター'],
  [3577,'テクノロジー・ガジェット','Creality K1C 3D Printer Carbon Fiber','Creality','Kickstarter','https://www.kickstarter.com/projects/creality/k1c-3d-printer-carbon-fiber','https://creality.com','support@creality.com','★★★★','カーボンファイバー対応・600mm/s・AI自動トラブルシューティング3Dプリンター'],
  [3578,'テクノロジー・ガジェット','Laser Pecker 4 Dual Laser Engraver','LaserPecker','Kickstarter','https://www.kickstarter.com/projects/laserpecker/laserpecker-4-dual-laser-engraver','https://laserpecker.net','support@laserpecker.net','★★★★★','2波長・金属直彫り・旋盤アタッチメント付き卓上デュアルレーザー彫刻機'],
  [3579,'テクノロジー・ガジェット','xTool M1 Desktop Laser Vinyl Cutter','xTool','Kickstarter','https://www.kickstarter.com/projects/xtool/m1-desktop-laser-vinyl-cutter','https://xtool.com','support@xtool.com','★★★★★','レーザー+カッター2in1・自動材料認識・クラフト用デスクトップDIYマシン'],
  [3580,'テクノロジー・ガジェット','Govee RGBICWW Floor Lamp Pro','Govee','Kickstarter','https://www.kickstarter.com/projects/govee/rgbicww-floor-lamp-pro','https://govee.com','support@govee.com','★★★★','Matter・ミュージックモード・グラデーション対応・スマートRGBICフロアランプ'],
  [3581,'テクノロジー・ガジェット','Epson EcoTank ET-3850 All-in-One Printer','Epson','The Grommet','https://thegrommet.com/product/tech/epson-ecotank-et-3850-all-in-one','https://epson.com','info@epson.com','★★★★★','インクカートリッジ不要・2年分インク同梱・WiFi・大容量エコタンクプリンター'],
  [3582,'テクノロジー・ガジェット','Rocketbook Fusion Smart Notebook','Rocketbook','The Grommet','https://thegrommet.com/product/tech/rocketbook-fusion-smart-notebook','https://getrocketbook.com','hello@getrocketbook.com','★★★★★','週間カレンダー+ノート+リスト統合・クラウド同期・繰り返し使えるノート'],
  [3583,'テクノロジー・ガジェット','Momento Ring Video Doorbell Pro','Momento','Kickstarter','https://www.kickstarter.com/projects/momento/ring-video-doorbell-pro','https://momentocam.com','support@momentocam.com','★★★★','3Dモーション検知・4K・顔認識・サブスク不要ローカル録画ビデオドアベル'],
  [3584,'テクノロジー・ガジェット','Shure MV7+ USB-XLR Podcast Microphone','Shure','Kickstarter','https://www.kickstarter.com/projects/shure/mv7-plus-usb-xlr-podcast-microphone','https://shure.com','support@shure.com','★★★★★','USB+XLR両対応・AIノイズキャンセル・ポッドキャスト/配信用ダイナミックマイク'],
  [3585,'テクノロジー・ガジェット','Elgato Stream Deck + Audio Mixer','Elgato','Kickstarter','https://www.kickstarter.com/projects/elgato/stream-deck-plus-audio-mixer','https://elgato.com','support@elgato.com','★★★★★','タッチバー+8ボタン・ダイヤルつまみ・配信/制作ワークフロー自動化コントローラー'],
  [3586,'テクノロジー・ガジェット','DJI Mic 2 Wireless Microphone System','DJI','Kickstarter','https://www.kickstarter.com/projects/dji/mic-2-wireless-microphone-system','https://dji.com','support@dji.com','★★★★★','250m伝送・内蔵録音・3.5mm/USB-C/Lightning対応ワイヤレスマイクシステム'],
  [3587,'テクノロジー・ガジェット','Insta360 X4 360° Action Camera','Insta360','Kickstarter','https://www.kickstarter.com/projects/insta360/insta360-x4-360-action-camera','https://insta360.com','hello@insta360.com','★★★★★','8K・360°撮影・アクティブHDR・AI編集アシスト360°アクションカメラ'],
  [3588,'テクノロジー・ガジェット','Halide Mark III Pro Camera App','Lux Optics','Kickstarter','https://www.kickstarter.com/projects/luxoptics/halide-mark-iii-camera-app','https://halide.cam','hello@halide.cam','★★★★','RAW+ProRAW対応・ニューラルデプス処理・プロカメラマン向けiPhoneカメラアプリ'],
  [3589,'テクノロジー・ガジェット','Loupedeck Live S Stream Controller','Loupedeck','Kickstarter','https://www.kickstarter.com/projects/loupedeck/live-s-stream-controller','https://loupedeck.com','support@loupedeck.com','★★★★','タッチディスプレイ+ダイヤル・OBS/Photoshop/Spotify連動クリエイターコントローラー'],
  [3590,'テクノロジー・ガジェット','Anker SOLIX F3800 Portable Power Station','Anker SOLIX','Kickstarter','https://www.kickstarter.com/projects/anker/solix-f3800-portable-power-station','https://anker.com','support@anker.com','★★★★★','3840Wh・LFP電池・6000W出力・停電/キャンプ/EV充電対応大型蓄電システム'],
  [3591,'テクノロジー・ガジェット','DJI Osmo Action 5 Pro Camera','DJI','Kickstarter','https://www.kickstarter.com/projects/dji/osmo-action-5-pro-camera','https://dji.com','support@dji.com','★★★★★','4K/120fps・-20℃動作・コールドスタート0.8秒・最強防水アクションカメラ'],

  // ── 美容・スキンケア (20件) ──
  [3592,'美容・スキンケア','Skin Gym Facial Cupping Set','Skin Gym','The Grommet','https://thegrommet.com/product/beauty/skin-gym-facial-cupping-set','https://skingym.com','info@skingym.com','★★★★','シリコン・セルフカッピング・リンパ流しとフェイスリフト効果のカッピングセット'],
  [3593,'美容・スキンケア','FOREO LUNA 4 Smart Facial Cleanser','FOREO','Kickstarter','https://www.kickstarter.com/projects/foreo/luna-4-smart-facial-cleanser','https://foreo.com','support@foreo.com','★★★★★','8,000振動/分・シリコンブラシ・AI肌分析連動スマートフェイシャルクレンザー'],
  [3594,'美容・スキンケア','CurrentBody Skin Tone Device','CurrentBody','Kickstarter','https://www.kickstarter.com/projects/currentbody/skin-tone-device','https://currentbody.com','support@currentbody.com','★★★★★','FDA認可・EMS+赤色光+温熱・コンビネーションリフティングデバイス'],
  [3595,'美容・スキンケア','Osmosis Rescue Epigenetic Activator Serum','Osmosis Beauty','The Grommet','https://thegrommet.com/product/beauty/osmosis-rescue-epigenetic-serum','https://osmosisbeauty.com','hello@osmosisbeauty.com','★★★★★','エピジェネティクス・DNA修復誘導・次世代エイジングケアセラム'],
  [3596,'美容・スキンケア','Augustinus Bader The Rich Cream','Augustinus Bader','The Grommet','https://thegrommet.com/product/beauty/augustinus-bader-the-rich-cream','https://augustinusbader.com','support@augustinusbader.com','★★★★★','TFC8技術・幹細胞活性化・世界的に注目の最高峰エイジングケアクリーム'],
  [3597,'美容・スキンケア','Tower 28 SOS Daily Rescue Facial Spray','Tower 28 Beauty','The Grommet','https://thegrommet.com/product/beauty/tower-28-sos-daily-rescue-spray','https://tower28beauty.com','hello@tower28beauty.com','★★★★','生理食塩水・敏感肌/湿疹対応・肌荒れ即鎮静ミスト型フェイシャルスプレー'],
  [3598,'美容・スキンケア','Byoma Hydrating Serum','Byoma','The Grommet','https://thegrommet.com/product/beauty/byoma-hydrating-serum','https://byoma.com','info@byoma.com','★★★★','TriSystem Complex・バリア機能強化・TikTok話題のバリアケアセラム'],
  [3599,'美容・スキンケア','Dieux Instant Angel Moisturizer','Dieux Skin','Kickstarter','https://www.kickstarter.com/projects/dieuxskin/instant-angel-moisturizer','https://dieuxskin.com','support@dieuxskin.com','★★★★','ペプチド・スクワラン・パンテノール・成分透明性を徹底したシンプルモイスチャライザー'],
  [3600,'美容・スキンケア','Typology Targeted Serum 10% Niacinamide','Typology','Kickstarter','https://www.kickstarter.com/projects/typology/targeted-serum-niacinamide','https://typology.com','hello@typology.com','★★★★','10%ナイアシンアミド+亜鉛・毛穴/テカリ/くすみ改善ターゲットセラム'],
  [3601,'美容・スキンケア','Drunk Elephant B-Goldi Bright Drops','Drunk Elephant','The Grommet','https://thegrommet.com/product/beauty/drunk-elephant-b-goldi-bright-drops','https://drunkelephant.com','support@drunkelephant.com','★★★★★','20%ビタミンC+Hが配合・肌荒れゼロで透明感アップの美白ドロップセラム'],
  [3602,'美容・スキンケア','Nécessaire The Body Wash','Nécessaire','The Grommet','https://thegrommet.com/product/beauty/necessaire-the-body-wash','https://necessaire.com','hello@necessaire.com','★★★★','ナイアシンアミド+ヒアルロン酸+ビタミンC/E・全身スキンケアボディウォッシュ'],
  [3603,'美容・スキンケア','Act+Acre Scalp Detox Treatment','Act+Acre','Kickstarter','https://www.kickstarter.com/projects/actandacre/scalp-detox-treatment','https://actandacre.com','info@actandacre.com','★★★★','コールドプロセス製法・乳酸菌・過酸化脂質除去スカルプデトックストリートメント'],
  [3604,'美容・スキンケア','Olaplex No.9 Bond Protector Nourishing Hair Serum','Olaplex','The Grommet','https://thegrommet.com/product/beauty/olaplex-no-9-bond-protector-serum','https://olaplex.com','support@olaplex.com','★★★★★','ボンドプロテクター・抗酸化・熱ダメージから髪を守るヘアセラム'],
  [3605,'美容・スキンケア','Florence by Mills Cloud Cushion Setting Powder','Florence by Mills','The Grommet','https://thegrommet.com/product/beauty/florence-by-mills-cloud-cushion-powder','https://florencebymills.com','hello@florencebymills.com','★★★★','マシュマロ質感・全トーン対応・テカリ抑制・Gen Z向けクラウドクッションパウダー'],
  [3606,'美容・スキンケア','e.l.f. Power Grip Primer','e.l.f. Cosmetics','The Grommet','https://thegrommet.com/product/beauty/elf-power-grip-primer','https://elfcosmetics.com','support@elfcosmetics.com','★★★★★','ヒアルロン酸・ジェル質感・メイクを12時間キープ・コスパ最強グリッププライマー'],
  [3607,'美容・スキンケア','Drmtlgy Needle-Less Serum Universal Tinted','Drmtlgy','Kickstarter','https://www.kickstarter.com/projects/drmtlgy/needle-less-serum-tinted','https://drmtlgy.com','info@drmtlgy.com','★★★★','ヒアルロン酸+コラーゲン+ペプチド・UVケア同時叶える万能カバーセラム'],
  [3608,'美容・スキンケア','Tatcha The Water Cream Oil-Free Moisturizer','Tatcha','The Grommet','https://thegrommet.com/product/beauty/tatcha-the-water-cream','https://tatcha.com','support@tatcha.com','★★★★★','Hadasei-3複合体・博多絹エキス・日本古来成分配合・皮脂ゼロ保湿クリーム'],
  [3609,'美容・スキンケア','Youth To The People Superfood Air-Whip Moisturizer','Youth To The People','The Grommet','https://thegrommet.com/product/beauty/yttp-superfood-air-whip-moisturizer','https://youthtothepeople.com','hello@youthtothepeople.com','★★★★','スーパーフード・ケール+ブルーベリー・ビーガン認証軽量エアウィップ保湿剤'],
  [3610,'美容・スキンケア','Loops Beauty Triple Dip Hydrating Face Mask','Loops Beauty','Kickstarter','https://www.kickstarter.com/projects/loopsbeauty/triple-dip-hydrating-face-mask','https://loopsbeauty.com','support@loopsbeauty.com','★★★★','三層浸透・即効性保湿・10分ケア完結のシートマスク'],
  [3611,'美容・スキンケア','Beekman 1802 Milk Drop Probiotic Serum','Beekman 1802','The Grommet','https://thegrommet.com/product/beauty/beekman-1802-milk-drop-probiotic-serum','https://beekman1802.com','info@beekman1802.com','★★★★★','ヤギミルク・プロバイオティクス・マイクロバイオーム研究に基づく肌菌バランスセラム'],

  // ── 子供・教育 (20件) ──
  [3612,'子供・教育','KiwiCo Tinker Crate STEM Kit','KiwiCo','Kickstarter','https://www.kickstarter.com/projects/kiwico/tinker-crate-stem-kit','https://kiwico.com','support@kiwico.com','★★★★★','月1配送・9-16歳向け・実験+エンジニアリング挑戦STEMサブスクキット'],
  [3613,'子供・教育','Miko 3 AI-Powered Robot for Kids','Miko','Kickstarter','https://www.kickstarter.com/projects/miko/miko-3-ai-powered-robot-for-kids','https://miko.ai','support@miko.ai','★★★★★','感情認識・会話AI・STEM学習コンテンツ内蔵子供向けAIコンパニオンロボット'],
  [3614,'子供・教育','Root Coding Robot for iPad','iRobot Education','Kickstarter','https://www.kickstarter.com/projects/irobot/root-coding-robot-for-ipad','https://rootrobotics.com','hello@rootrobotics.com','★★★★','磁石で壁/床を走行・3レベルコーディング・iPad連動プログラミングロボット'],
  [3615,'子供・教育','Toca Life World Build Stories App Bundle','Toca Boca','The Grommet','https://thegrommet.com/product/family-kids/toca-life-world-starter-kit','https://tocaboca.com','support@tocaboca.com','★★★★★','スウェーデン発・広告なし・オープンエンドの子供向けデジタル世界観ゲーム'],
  [3616,'子供・教育','Goldie Blox Construction Toy','GoldieBlox','Kickstarter','https://www.kickstarter.com/projects/goldieblox/goldieblox-the-engineering-toy-for-girls','https://goldieblox.com','hello@goldieblox.com','★★★★','女の子向けエンジニアリング玩具・絵本+ブロック・STEM興味を引き出す'],
  [3617,'子供・教育','Roominate Wired for STEM Building Kit','Roominate','Kickstarter','https://www.kickstarter.com/projects/roominate/roominate-wired-for-stem','https://roominatetoy.com','support@roominatetoy.com','★★★★','回路内蔵・LED/モーター付き・部屋をDIYする女の子向け建築STEMキット'],
  [3618,'子供・教育','Sphero BOLT Programmable Robot Ball','Sphero','Kickstarter','https://www.kickstarter.com/projects/sphero/sphero-bolt-programmable-robot-ball','https://sphero.com','support@sphero.com','★★★★★','LED 8x8マトリクス・障害物センシング・Scratch/JS対応プログラミングロボットボール'],
  [3619,'子供・教育','Primo Cubetto Playset Coding Toy','Primo Toys','Kickstarter','https://www.kickstarter.com/projects/primotoys/cubetto-playset','https://primotoys.com','hello@primotoys.com','★★★★★','画面なし・木製ブロック・3歳からのMontessori型プログラミング玩具'],
  [3620,'子供・教育','Smartivity Marble Run STEM Kit','Smartivity','Kickstarter','https://www.kickstarter.com/projects/smartivity/marble-run-stem-kit','https://smartivity.in','support@smartivity.in','★★★★','レーザーカット木材・物理学・ビー玉コースを設計するSTEM建築キット'],
  [3621,'子供・教育','FabricLab Sewable Electronics Kit','FabricLab','Kickstarter','https://www.kickstarter.com/projects/fabriclab/sewable-electronics-kit','https://fabriclab.com','info@fabriclab.com','★★★★','縫い付けるLED回路・ウェアラブル作成・ファッション×電子工学STEM体験キット'],
  [3622,'子供・教育','Playmation Marvel Avengers Starter Pack','Hasbro Playmation','The Grommet','https://thegrommet.com/product/family-kids/playmation-marvel-avengers-starter-pack','https://hasbro.com','support@hasbro.com','★★★★','AR+物理アクション連動・体を動かすマーベルヒーロー体験ゲームセット'],
  [3623,'子供・教育','Cubelets Robot Blocks Curiosity Set','Modular Robotics','Kickstarter','https://www.kickstarter.com/projects/modrobotics/cubelets-robot-blocks','https://modrobotics.com','support@modrobotics.com','★★★★★','磁石連結・プログラム不要・積み方でロボット動作が変わる知育ロボットブロック'],
  [3624,'子供・教育','Play Shifu Tacto Chess Digital Board Game','PlayShifu','Kickstarter','https://www.kickstarter.com/projects/playshifu/tacto-chess-digital-board-game','https://playshifu.com','hello@playshifu.com','★★★★','スマートフォンをボードに置くデジタルチェス・AIコーチ付き学習将棋ゲーム'],
  [3625,'子供・教育','SmartMax My First Jungle Animals','SmartMax','The Grommet','https://thegrommet.com/product/family-kids/smartmax-my-first-jungle-animals','https://smartmax.com','info@smartmax.com','★★★★','BPA不使用・磁石接続・1歳から安心して遊べるジャングル動物磁気ブロック'],
  [3626,'子供・教育','VTech DigiGo Kids Electronic Tablet','VTech','The Grommet','https://thegrommet.com/product/family-kids/vtech-digigo-kids-tablet','https://vtech.com','support@vtech.com','★★★★','子供用OS・保護者コントロール・教育アプリ搭載ジュニアタブレット'],
  [3627,'子供・教育','ThinkFun Gravity Maze Marble Run Logic Game','ThinkFun','The Grommet','https://thegrommet.com/product/family-kids/thinkfun-gravity-maze-marble-run','https://thinkfun.com','hello@thinkfun.com','★★★★★','60チャレンジ・3Dマンスリン迷路・論理的思考を鍛える重力迷路パズル'],
  [3628,'子供・教育','Crayola Light Up Tracing Pad','Crayola','The Grommet','https://thegrommet.com/product/family-kids/crayola-light-up-tracing-pad','https://crayola.com','info@crayola.com','★★★★','LEDバックライト・20枚トレーシングシート付き・子供のお絵描き学習パッド'],
  [3629,'子供・教育','MBI Tinker Bell Glitter Globe','Enesco','The Grommet','https://thegrommet.com/product/family-kids/enesco-tinker-bell-glitter-globe','https://enesco.com','support@enesco.com','★★★','手描きディテール・回転する光・ティンカーベルミュージカルグリッターグローブ'],
  [3630,'子供・教育','Hatch Toddler Sleep Trainer Rest+','Hatch Baby','The Grommet','https://thegrommet.com/product/family-kids/hatch-rest-plus-toddler-sleep-trainer','https://hatch.co','support@hatch.co','★★★★★','起床OKライト・音・WiFi・スマホ管理可能な幼児向けスマートナイトライト'],
  [3631,'子供・教育','Pictionary Air Drawing Game','Mattel Games','The Grommet','https://thegrommet.com/product/family-kids/pictionary-air-drawing-game','https://mattel.com','support@mattel.com','★★★★','ARペン・空中に描いた絵がTV映し出される・家族向けデジタルお絵かきゲーム'],

  // ── ファッション・アクセサリー (20件) ──
  [3632,'ファッション・アクセサリー','Blunt Classic Large Umbrella','Blunt Umbrellas','The Grommet','https://thegrommet.com/product/fashion/blunt-classic-large-umbrella','https://bluntumbrellas.com','support@bluntumbrellas.com','★★★★★','フレームレスエッジ・風速100km/h耐性・ニュージーランド発最強耐久傘'],
  [3633,'ファッション・アクセサリー','STNKY Laundry Bag Sports Odor Control','STNKY','Kickstarter','https://www.kickstarter.com/projects/stnky/stnky-laundry-bag-odor-control','https://stnky.com','hello@stnky.com','★★★★','炭素フィルター内蔵・スポーツウェアの臭いを閉じ込める消臭ランドリーバッグ'],
  [3634,'ファッション・アクセサリー','MUD Jeans Lease a Jeans Circular Fashion','MUD Jeans','Kickstarter','https://www.kickstarter.com/projects/mudjeans/lease-a-jeans-circular-fashion','https://mudjeans.eu','support@mudjeans.eu','★★★★★','オーガニックコットン・完全循環型・リース＆リサイクルサーキュラーデニム'],
  [3635,'ファッション・アクセサリー','Stormy Kromer Original Wool Cap','Stormy Kromer','The Grommet','https://thegrommet.com/product/fashion/stormy-kromer-original-wool-cap','https://stormykromer.com','info@stormykromer.com','★★★★★','ウール100%・耳当て付き・1903年創業・米国製クラシックウールキャップ'],
  [3636,'ファッション・アクセサリー','Bellroy Venture Hip Pack','Bellroy','The Grommet','https://thegrommet.com/product/fashion/bellroy-venture-hip-pack','https://bellroy.com','support@bellroy.com','★★★★','リサイクルナイロン・多機能ポケット・ウエスト/クロスボディ両対応ヒップパック'],
  [3637,'ファッション・アクセサリー','Pela Case Compostable iPhone Case','Pela Case','Kickstarter','https://www.kickstarter.com/projects/pela/pela-case-compostable-iphone-case','https://pelacase.com','hello@pelacase.com','★★★★','植物由来・堆肥化可能・カーボンネガティブな環境配慮スマホケース'],
  [3638,'ファッション・アクセサリー','Cariuma OCA Low Sneaker Sustainable','Cariuma','Kickstarter','https://www.kickstarter.com/projects/cariuma/oca-low-sustainable-sneaker','https://cariuma.com','support@cariuma.com','★★★★★','竹+サトウキビ素材・B Corp認証・カーボンニュートラルスニーカー'],
  [3639,'ファッション・アクセサリー','tentree Classic Hoodie Sustainable','tentree','Kickstarter','https://www.kickstarter.com/projects/tentree/classic-hoodie-sustainable','https://tentree.com','hello@tentree.com','★★★★','1枚購入で10本植樹・オーガニックコットン・サステナブルパーカー'],
  [3640,'ファッション・アクセサリー','Fjällräven Kånken Mini Backpack','Fjällräven','The Grommet','https://thegrommet.com/product/fashion/fjallraven-kanken-mini-backpack','https://fjallraven.com','support@fjallraven.com','★★★★★','Vinylon F素材・スウェーデン発・1978年誕生の定番ミニバックパック'],
  [3641,'ファッション・アクセサリー','Rains Bag Base Backpack','Rains','Kickstarter','https://www.kickstarter.com/projects/rains/bag-base-backpack','https://rains.com','info@rains.com','★★★★','完全防水・ミニマルデザイン・デンマーク発都市型レインバックパック'],
  [3642,'ファッション・アクセサリー','Huckberry Ranch Road Boots Western','Huckberry Ranch Road Boots','Kickstarter','https://www.kickstarter.com/projects/huckberry/ranch-road-boots-western','https://ranchroadboots.com','support@ranchroadboots.com','★★★★','グッドイヤーウェルト・フルグレインレザー・日常使いウェスタンブーツ'],
  [3643,'ファッション・アクセサリー','Anker SoundCore Liberty 4 NC Earbuds','Anker Soundcore','The Grommet','https://thegrommet.com/product/tech/soundcore-liberty-4-nc-earbuds','https://soundcore.com','support@soundcore.com','★★★★★','98.5%ANC・LDAC・ハイレゾ対応・コスパ最強ワイヤレスノイキャンイヤホン'],
  [3644,'ファッション・アクセサリー','Tom Bihn Synik 30 Travel Backpack','Tom Bihn','The Grommet','https://thegrommet.com/product/fashion/tom-bihn-synik-30-travel-backpack','https://tombihn.com','info@tombihn.com','★★★★★','米国製・30L・15インチPC対応・機内持ち込み最適化旅行バックパック'],
  [3645,'ファッション・アクセサリー','Seiko Presage Sharp Edged Automatic Watch','Seiko','The Grommet','https://thegrommet.com/product/fashion/seiko-presage-sharp-edged-watch','https://seiko.com','support@seikousa.com','★★★★★','有田焼ダイヤル・自動巻き・日本工芸とウォッチクラフトの融合機械式時計'],
  [3646,'ファッション・アクセサリー','Nau Merino Hoodie Sustainable','Nau','Kickstarter','https://www.kickstarter.com/projects/nau/merino-hoodie-sustainable','https://nau.com','hello@nau.com','★★★★','メリノウール・再生ポリエステル混・オレゴン発サステナブルアウトドアパーカー'],
  [3647,'ファッション・アクセサリー','Halo Smart Wireless Security Tag','Halo Wearables','Kickstarter','https://www.kickstarter.com/projects/halowearables/halo-smart-security-tag','https://halowearables.com','support@halowearables.com','★★★★','超軽量・Find My対応・GPSトラッカー機能付きスマートタグアクセサリー'],
  [3648,'ファッション・アクセサリー','Kin Euphorics Non-Alcoholic Spirits','Kin Euphorics','Kickstarter','https://www.kickstarter.com/projects/kineuphoric/kin-euphorics-non-alcoholic-spirits','https://kineuphoric.com','hello@kineuphoric.com','★★★★','アダプトゲン+ノートロピクス配合・アルコール不使用のスピリチュアルドリンク'],
  [3649,'ファッション・アクセサリー','Paka Alpaca Wool Performance Quarter Zip','Paka Apparel','Kickstarter','https://www.kickstarter.com/projects/paka/alpaca-wool-performance-quarter-zip','https://pakaapparel.com','hello@pakaapparel.com','★★★★★','100%アルパカウール・温度調整・消臭・化学薬品不使用ハイパフォーマンスウェア'],
  [3650,'ファッション・アクセサリー','Birdies The Blackbird Heeled Bootie','Birdies','The Grommet','https://thegrommet.com/product/fashion/birdies-blackbird-heeled-bootie','https://birdiesshoes.com','info@birdiesshoes.com','★★★★★','医師設計インソール・スタイリッシュ・1日中履いても痛くないヒールブーティ'],
  [3651,'ファッション・アクセサリー','Darn Tough Hiker Boot Cushion Sock','Darn Tough Vermont','The Grommet','https://thegrommet.com/product/fashion/darn-tough-hiker-boot-cushion-sock','https://darntough.com','info@darntough.com','★★★★★','Merino Wool・生涯保証・ミッドカッション・バーモント製ハイキングソックス'],

  // ── クリーニング・収納・整理 (20件) ──
  [3652,'クリーニング・収納・整理','Blueland Clean Essentials Starter Kit','Blueland','Kickstarter','https://www.kickstarter.com/projects/blueland/clean-essentials-starter-kit','https://blueland.com','info@blueland.com','★★★★★','錠剤を水に溶かすだけ・プラスチック削減・全種エコ掃除スターターキット'],
  [3653,'クリーニング・収納・整理','Grove Co. Plastic Free Dish Soap','Grove Collaborative','The Grommet','https://thegrommet.com/product/home/grove-co-plastic-free-dish-soap','https://grove.co','support@grove.co','★★★★','バー状・プラスチックなし・パーム油不使用・濃縮食器用洗剤バー'],
  [3654,'クリーニング・収納・整理','Dropps Laundry & Dish Pods','Dropps','Kickstarter','https://www.kickstarter.com/projects/dropps/laundry-dish-pods-plastic-free','https://dropps.com','hello@dropps.com','★★★★','PVAフィルム包装・濃縮・乳幼児用・プラスチックフリー洗濯&食洗機ポッド'],
  [3655,'クリーニング・収納・整理','Truman\'s Cleaning Concentrates Refill Kit','Truman\'s','Kickstarter','https://www.kickstarter.com/projects/trumans/cleaning-concentrates-refill-kit','https://trumans.com','support@trumans.com','★★★★','カートリッジ交換式・水で薄める・廃棄物90%削減サステナブル洗剤システム'],
  [3656,'クリーニング・収納・整理','Kognito Bamboo Dish Drying Rack','Kognito','The Grommet','https://thegrommet.com/product/home/kognito-bamboo-dish-drying-rack','https://kognito.co','hello@kognito.co','★★★★','天然竹・折りたたみ・伸縮ドレンボード付きエコ食器水切りラック'],
  [3657,'クリーニング・収納・整理','Umundo Stacking Containers Set','Umundo','Kickstarter','https://www.kickstarter.com/projects/umundo/stacking-containers-modular-storage','https://umundo.com','support@umundo.com','★★★★★','モジュラー・スタッキング・カスタマイズ可能な次世代収納コンテナシステム'],
  [3658,'クリーニング・収納・整理','iDesign AFFIXX Adhesive Kitchen Organizer','iDesign','The Grommet','https://thegrommet.com/product/home/idesign-affixx-adhesive-kitchen-organizer','https://idesignlivesimply.com','info@idesignlivesimply.com','★★★★','貼り付け式・取り外し可能・賃貸OK・冷蔵庫側面活用収納オーガナイザー'],
  [3659,'クリーニング・収納・整理','OXO Good Grips Sweep & Swipe Laptop Cleaner','OXO','The Grommet','https://thegrommet.com/product/tech/oxo-sweep-swipe-laptop-cleaner','https://oxo.com','info@oxo.com','★★★','マイクロファイバー+ブラシ一体型・PCスクリーン&キーボード清掃ツール'],
  [3660,'クリーニング・収納・整理','Simplehuman Sensor Trash Can 45L','simplehuman','The Grommet','https://thegrommet.com/product/home/simplehuman-sensor-trash-can-45l','https://simplehuman.com','hello@simplehuman.com','★★★★★','ボイスアクティベーション・USB充電・防臭リング付き45L大型センサーゴミ箱'],
  [3661,'クリーニング・収納・整理','Elfa Utility Garage Wall System','Container Store','The Grommet','https://thegrommet.com/product/home/elfa-utility-garage-wall-system','https://containerstore.com','support@containerstore.com','★★★★','カスタマイズ自在・スウェーデン発・ガレージ壁面収納最適化エルファシステム'],
  [3662,'クリーニング・収納・整理','Casabrews 20 Bar Espresso Machine','Casabrews','The Grommet','https://thegrommet.com/product/kitchen/casabrews-espresso-machine','https://casabrews.com','support@casabrews.com','★★★★','20バール・スチームワンド付き・デロンギの半額以下の本格エスプレッソマシン'],
  [3663,'クリーニング・収納・整理','OXO SteeL Pop Container Set','OXO','The Grommet','https://thegrommet.com/product/home/oxo-steel-pop-container-set','https://oxo.com','info@oxo.com','★★★★★','ボタン一押し密封・スタック収納・ステンレス蓋付き食品保存容器セット'],
  [3664,'クリーニング・収納・整理','Yamazaki Tower Wall-Mounted Paper Towel Holder','Yamazaki Home','The Grommet','https://thegrommet.com/product/home/yamazaki-tower-paper-towel-holder','https://yamazakihome.com','hello@yamazakihome.com','★★★★','スチール・日本発ミニマルデザイン・壁掛け+スタンド兼用ペーパータオルホルダー'],
  [3665,'クリーニング・収納・整理','Umbra Trigg Display Shelf Floating','Umbra','The Grommet','https://thegrommet.com/product/home/umbra-trigg-display-shelf','https://umbra.com','hello@umbra.com','★★★★','三角形ウォールシェルフ+フックセット・カナダ発インダストリアルデザイン壁収納'],
  [3666,'クリーニング・収納・整理','Method Laundry Detergent Pump Bottle','Method Products','The Grommet','https://thegrommet.com/product/home/method-laundry-detergent-pump-bottle','https://methodproducts.com','support@methodproducts.com','★★★★','濃縮50回分・ポンプ式・天然由来成分・プラスチック削減洗濯洗剤ポンプボトル'],
  [3667,'クリーニング・収納・整理','Brandless Cleaning Wipes Multi-Surface','Brandless','The Grommet','https://thegrommet.com/product/home/brandless-cleaning-wipes-multi-surface','https://brandless.com','info@brandless.com','★★★','植物由来・香料フリー・100枚入り敏感肌対応多目的クリーニングワイプ'],
  [3668,'クリーニング・収納・整理','Tidy Tiger Professional Organize Your Life','Tidy Tiger','Kickstarter','https://www.kickstarter.com/projects/tidytiger/professional-organize-your-life-system','https://tidytiger.com','hello@tidytiger.com','★★★★','プロ整理収納士監修・デジタル+物理ハイブリッド・ライフ整理管理システム'],
  [3669,'クリーニング・収納・整理','Rubbermaid Reveal Spray Mop Kit','Rubbermaid','The Grommet','https://thegrommet.com/product/home/rubbermaid-reveal-spray-mop-kit','https://rubbermaid.com','support@rubbermaid.com','★★★★','詰め替え式スプレー・洗濯可能パッド・フローリング対応スプレーモップキット'],
  [3670,'クリーニング・収納・整理','Stardrops The Pink Stuff Cleaning Paste Bundle','Star Drops','The Grommet','https://thegrommet.com/product/home/stardrops-pink-stuff-cleaning-paste-bundle','https://thepinkstuffcleaner.com','hello@thepinkstuffcleaner.com','★★★★★','ペースト+スプレー+クリーム3点セット・頑固汚れ撃退英国発万能洗剤セット'],
  [3671,'クリーニング・収納・整理','Casabella Premium Rubber Kitchen Gloves','Casabella','The Grommet','https://thegrommet.com/product/home/casabella-premium-rubber-kitchen-gloves','https://casabella.com','support@casabella.com','★★★★','ベルベット裏地・手首ガード付き・洗い物/ガーデニング兼用プレミアムゴム手袋'],
];

// ── ユーティリティ ──────────────────────────────────────────────────────────
function buildGmailUrl(toEmail, maker, product) {
  const enc = s => encodeURIComponent(s);
  const subject = 'Potential Distribution Partnership for Japan';
  const body =
`Dear ${maker} Team,

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

E-mail: ${SENDER}
Address: 49-5 Kitazakuno, Higashi-Itsushiro, Ichinomiya City, Aichi Prefecture, Japan

=================================================`;
  return `https://mail.google.com/mail/?authuser=${enc(SENDER)}&view=cm&fs=1&to=${enc(toEmail)}&su=${enc(subject)}&body=${enc(body)}`;
}

// ── メイン処理 ──────────────────────────────────────────────────────────────
console.log('Reading:', INPUT);
const wb_in   = XLSX.readFile(INPUT);
const ws_in   = wb_in.Sheets[wb_in.SheetNames[0]];
const existing = XLSX.utils.sheet_to_json(ws_in, { header: 1 });
console.log(`既存行数: ${existing.length - 1} 件`);

const allRows = [...existing];
for (const p of NEW_PRODUCTS) allRows.push(p);

const wb_out = XLSX.utils.book_new();
const ws_out = XLSX.utils.aoa_to_sheet(allRows);

ws_out['!cols'] = [
  { wch: 6 }, { wch: 18 }, { wch: 40 }, { wch: 22 },
  { wch: 14 }, { wch: 52 }, { wch: 35 }, { wch: 35 },
  { wch: 8  }, { wch: 45 },
];

for (let r = 1; r < allRows.length; r++) {
  const row = allRows[r];
  const mailAddr = String(row[7] || '');
  if (mailAddr && mailAddr.includes('@')) {
    const cellAddr = XLSX.utils.encode_cell({ r, c: 7 });
    const gmailUrl = buildGmailUrl(mailAddr, String(row[3] || ''), String(row[2] || ''));
    ws_out[cellAddr] = {
      v: mailAddr, t: 's',
      l: { Target: gmailUrl },
      s: { font: { color: { rgb: '1155CC' }, underline: true } }
    };
  }
}

XLSX.utils.book_append_sheet(wb_out, ws_out, '製品リスト');
XLSX.writeFile(wb_out, OUTPUT);

const mailCount = NEW_PRODUCTS.filter(p => p[7] && p[7].includes('@')).length;
const cfCount   = NEW_PRODUCTS.filter(p => ['Kickstarter','Indiegogo'].includes(p[4])).length;
const grCount   = NEW_PRODUCTS.filter(p => p[4] === 'The Grommet').length;
const catCount  = {};
for (const p of NEW_PRODUCTS) catCount[p[1]] = (catCount[p[1]] || 0) + 1;

console.log('\n====== 追加完了レポート ======');
console.log(`追加件数       : ${NEW_PRODUCTS.length} 件`);
console.log(`メールあり     : ${mailCount} 件`);
console.log(`Kickstarter/IF : ${cfCount} 件`);
console.log(`The Grommet    : ${grCount} 件`);
console.log('\nカテゴリ別:');
Object.entries(catCount).forEach(([c,n]) => console.log(`  ${c}: ${n}件`));
console.log(`\n総収録件数     : ${existing.length - 1 + NEW_PRODUCTS.length} 件`);
console.log(`出力ファイル   : ${OUTPUT}`);
