/**
 * add_grommet_new_107.js
 * The Grommet専用ダッシュボード用・新規追加107件 (No.294〜400)
 * 既存293件（add_200_items_14.js全200件 + add_200_items_15.jsのGrommet分93件）と合算してNo.1〜400完成
 * 各カテゴリ40件になるようバランス配分
 */

// [番号, カテゴリ, 製品名, メーカー, ECサイト, 製品URL, メーカーHP, メール, ★, コメント]
const NEW_PRODUCTS_107 = [
  // ── キッチン・調理器具 (6件・294-299) ──
  [294,'キッチン・調理器具','GreenPan Valencia Pro Ceramic Nonstick Skillet','GreenPan','The Grommet','https://thegrommet.com/product/kitchen/greenpan-valencia-pro-skillet','https://greenpan.us','support@greenpan.us','★★★★','セラミックコーティング・PFAS不使用・耐久性強化ノンスティックフライパン'],
  [295,'キッチン・調理器具','Instant Pot Duo Crisp 11-in-1','Instant Brands','The Grommet','https://thegrommet.com/product/kitchen/instant-pot-duo-crisp','https://instantpot.com','support@instantpot.com','★★★★★','圧力調理+エアフライ11役・ワンポットで揚げ物まで作れる多機能調理器'],
  [296,'キッチン・調理器具','Caraway Non-Toxic Cookware Set','Caraway Home','The Grommet','https://thegrommet.com/product/kitchen/caraway-cookware-set','https://carawayhome.com','hello@carawayhome.com','★★★★★','有害物質不使用セラミックコーティング・収納ラック付き調理器具セット'],
  [297,'キッチン・調理器具','Zulay Kitchen Original Avocado Slicer','Zulay Kitchen','The Grommet','https://thegrommet.com/product/kitchen/zulay-avocado-slicer','https://zulaykitchen.com','hello@zulaykitchen.com','★★★','スライス・種取り・皮むき一体型アボカドツール'],
  [298,'キッチン・調理器具','Our Place Always Pan 2.0','Our Place','The Grommet','https://thegrommet.com/product/kitchen/our-place-always-pan','https://fromourplace.com','support@fromourplace.com','★★★★★','8-in-1機能・セラミック非粘着・SNSで話題の万能フライパン'],
  [299,'キッチン・調理器具',"Chef'n FreshForce Citrus Juicer","Chef'n",'The Grommet','https://thegrommet.com/product/kitchen/chefn-freshforce-citrus-juicer','https://chefn.com','info@chefn.com','★★★★','てこの原理で果汁を最大限搾り取るシトラスジューサー'],

  // ── スマートホーム・インテリア・照明 (13件・300-312) ──
  [300,'スマートホーム・インテリア・照明','Govee Glide Wall Light','Govee','The Grommet','https://thegrommet.com/product/home/govee-glide-wall-light','https://govee.com','support@govee.com','★★★★','タイルモジュラー式・音楽同期・アプリ制御ウォールライト'],
  [301,'スマートホーム・インテリア・照明','Wyze Cam v4','Wyze Labs','The Grommet','https://thegrommet.com/product/home/wyze-cam-v4','https://wyze.com','support@wyze.com','★★★★','カラーナイトビジョン・スマート通知・低価格スマートホームカメラ'],
  [302,'スマートホーム・インテリア・照明','Ecobee Smart Thermostat Premium','ecobee','The Grommet','https://thegrommet.com/product/home/ecobee-smart-thermostat-premium','https://ecobee.com','support@ecobee.com','★★★★★','室内空気質センサー内蔵・Alexa搭載・省エネスマートサーモスタット'],
  [303,'スマートホーム・インテリア・照明','Nanoleaf Lines Modular Light Panels','Nanoleaf','The Grommet','https://thegrommet.com/product/home/nanoleaf-lines','https://nanoleaf.me','hello@nanoleaf.me','★★★★','直線モジュール式・音楽反応・壁面デザイン照明'],
  [304,'スマートホーム・インテリア・照明','LIFX Color A19 Smart Bulb','LIFX','The Grommet','https://thegrommet.com/product/home/lifx-color-a19','https://lifx.com','support@lifx.com','★★★','ハブ不要・WiFi直結・1600万色対応スマート電球'],
  [305,'スマートホーム・インテリア・照明','Yale Assure Lock 2','Yale','The Grommet','https://thegrommet.com/product/home/yale-assure-lock-2','https://yalehome.com','support@yalehome.com','★★★★★','タッチスクリーン・指紋認証対応・後付けスマート玄関錠'],
  [306,'スマートホーム・インテリア・照明','SwitchBot Curtain Robot','SwitchBot','The Grommet','https://thegrommet.com/product/home/switchbot-curtain-robot','https://switch-bot.com','support@switch-bot.com','★★★★','既存カーテンレールに後付け・スマホ/音声操作自動開閉ロボット'],
  [307,'スマートホーム・インテリア・照明','Levoit Core 400S Smart Air Purifier','Levoit','The Grommet','https://thegrommet.com/product/home/levoit-core-400s','https://levoit.com','support@levoit.com','★★★★','HEPA H13・アプリ連携・PM2.5対応スマート空気清浄機'],
  [308,'スマートホーム・インテリア・照明','Aqara Smart Hub M3','Aqara','The Grommet','https://thegrommet.com/product/home/aqara-smart-hub-m3','https://aqara.com','support@aqara.com','★★★★','Zigbee/Matter対応・ローカル処理・拡張性抜群スマートホームハブ'],
  [309,'スマートホーム・インテリア・照明','Casper Glow Light','Casper','The Grommet','https://thegrommet.com/product/home/casper-glow-light','https://casper.com','hello@casper.com','★★★★','タップで明滅・タイマー付き・入眠/起床サポートライト'],
  [310,'スマートホーム・インテリア・照明','Ring Alarm Pro Security Kit','Ring','The Grommet','https://thegrommet.com/product/home/ring-alarm-pro-security-kit','https://ring.com','support@ring.com','★★★★','eeroルーター内蔵・侵入検知・DIY設置セキュリティシステム'],
  [311,'スマートホーム・インテリア・照明','Brightech Litespan LED Floor Lamp','Brightech','The Grommet','https://thegrommet.com/product/home/brightech-litespan-led-floor-lamp','https://brightechled.com','support@brightechled.com','★★★','20年寿命LED・省エネ・北欧デザインフロアランプ'],
  [312,'スマートホーム・インテリア・照明','Eufy Security Video Doorbell E340','Eufy','The Grommet','https://thegrommet.com/product/home/eufy-video-doorbell-e340','https://eufylife.com','support@eufylife.com','★★★★','デュアルカメラ・ローカルストレージ・月額不要スマートドアベル'],

  // ── ウェアラブル・ヘルス・フィットネス (18件・313-330) ──
  [313,'ウェアラブル・ヘルス・フィットネス','Withings Body Scan Smart Scale','Withings','The Grommet','https://thegrommet.com/product/health-wellness/withings-body-scan','https://withings.com','support@withings.com','★★★★★','心電図+体組成+神経系スキャン・医療グレードスマート体重計'],
  [314,'ウェアラブル・ヘルス・フィットネス','Apollo Neuro Wearable Stress Relief','Apollo Neuroscience','The Grommet','https://thegrommet.com/product/health-wellness/apollo-neuro-wearable','https://apolloneuro.com','hello@apolloneuro.com','★★★★★','触覚振動・自律神経調整・ストレス緩和ウェアラブル'],
  [315,'ウェアラブル・ヘルス・フィットネス','Muse S Meditation Headband','Muse','The Grommet','https://thegrommet.com/product/health-wellness/muse-s-meditation-headband','https://choosemuse.com','support@choosemuse.com','★★★★','EEG脳波センサー・リアルタイムフィードバック瞑想サポートヘッドバンド'],
  [316,'ウェアラブル・ヘルス・フィットネス','Biostrap EVO Wearable','Biostrap','The Grommet','https://thegrommet.com/product/health-wellness/biostrap-evo-wearable','https://biostrap.com','support@biostrap.com','★★★★','PPG+モーションセンサー・睡眠/回復精密解析リストバンド'],
  [317,'ウェアラブル・ヘルス・フィットネス','Amazfit Balance Smartwatch','Amazfit','The Grommet','https://thegrommet.com/product/health-wellness/amazfit-balance-smartwatch','https://amazfit.com','support@amazfit.com','★★★★','BioTracker5.0・栄養/ストレススコア・長時間バッテリースマートウォッチ'],
  [318,'ウェアラブル・ヘルス・フィットネス','Therabody SmartGoggles','Therabody','The Grommet','https://thegrommet.com/product/health-wellness/therabody-smartgoggles','https://therabody.com','support@therabody.com','★★★★','振動+熱+圧迫・目元疲労回復スマートゴーグル'],
  [319,'ウェアラブル・ヘルス・フィットネス','Wahoo TRACKR Heart Rate Monitor','Wahoo Fitness','The Grommet','https://thegrommet.com/product/health-wellness/wahoo-trackr-heart-rate','https://wahoofitness.com','support@wahoofitness.com','★★★','胸部装着・ANT+/Bluetooth両対応心拍センサー'],
  [320,'ウェアラブル・ヘルス・フィットネス','Bala Bangles Wearable Weights','Bala','The Grommet','https://thegrommet.com/product/health-wellness/bala-bangles-wearable-weights','https://yourbala.com','hello@yourbala.com','★★★★','おしゃれデザイン・1kg×2個・日常に負荷を加えるリスト/アンクルウェイト'],
  [321,'ウェアラブル・ヘルス・フィットネス','Nuyu Sleep System Wearable','Nuyu','The Grommet','https://thegrommet.com/product/health-wellness/nuyu-sleep-system','https://nuyusleep.com','support@nuyusleep.com','★★★','体温調整パッチ・深部体温を下げ入眠をサポートするウェアラブル'],
  [322,'ウェアラブル・ヘルス・フィットネス','Kegg Fertility Tracker','Kegg','The Grommet','https://thegrommet.com/product/health-wellness/kegg-fertility-tracker','https://kegg.tech','support@kegg.tech','★★★★','膣分泌液センサー・排卵予測女性向けヘルストラッカー'],
  [323,'ウェアラブル・ヘルス・フィットネス','Carol Bike AI Resistance Bike','Carol Bike','The Grommet','https://thegrommet.com/product/health-wellness/carol-bike-ai-resistance','https://carolbike.com','support@carolbike.com','★★★★★','AI抵抗調整・8分REHIT・短時間高効率フィットネスバイク'],
  [324,'ウェアラブル・ヘルス・フィットネス','Hyperice Normatec 3 Leg Recovery System','Hyperice','The Grommet','https://thegrommet.com/product/health-wellness/hyperice-normatec-3','https://hyperice.com','support@hyperice.com','★★★★★','動的空気圧縮・パルス技術・アスリート向け脚部リカバリーシステム'],
  [325,'ウェアラブル・ヘルス・フィットネス','Withings BeamO Multi-Scope','Withings','The Grommet','https://thegrommet.com/product/health-wellness/withings-beamo-multi-scope','https://withings.com','support@withings.com','★★★★','体温+心拍+SpO2+聴診・4in1家庭用健康モニタリングデバイス'],
  [326,'ウェアラブル・ヘルス・フィットネス','Modius Health Weight Loss Headset','Modius Health','The Grommet','https://thegrommet.com/product/health-wellness/modius-health-headset','https://modiushealth.com','support@modiushealth.com','★★★','前庭神経刺激・脳への電気刺激で代謝サポートするヘッドセット'],
  [327,'ウェアラブル・ヘルス・フィットネス','Moonbird Breathing Coach','Moonbird','The Grommet','https://thegrommet.com/product/health-wellness/moonbird-breathing-coach','https://moonbird.life','hello@moonbird.life','★★★★','手のひらサイズ・膨張収縮で呼吸ペースを導く触覚呼吸コーチ'],
  [328,'ウェアラブル・ヘルス・フィットネス','Bosign Anti-Snore Pillow','Bosign','The Grommet','https://thegrommet.com/product/health-wellness/bosign-anti-snore-pillow','https://bosign.com','info@bosign.com','★★★','横向き寝誘導形状・いびき軽減エルゴノミクス枕'],
  [329,'ウェアラブル・ヘルス・フィットネス','Compex Fit 5.0 Muscle Stimulator','Compex','The Grommet','https://thegrommet.com/product/health-wellness/compex-fit-5-muscle-stimulator','https://compexstore.com','support@compexstore.com','★★★★','EMS・筋力強化+リカバリー・アスリート向け電気刺激トレーナー'],
  [330,'ウェアラブル・ヘルス・フィットネス','Kinsa QuickCare Smart Thermometer','Kinsa','The Grommet','https://thegrommet.com/product/health-wellness/kinsa-quickcare-smart-thermometer','https://kinsahealth.com','support@kinsahealth.com','★★★★','スマホ連動・症状記録・家族の健康管理スマート体温計'],

  // ── アウトドア・スポーツ・旅行 (14件・331-344) ──
  [331,'アウトドア・スポーツ・旅行','Nemo Tensor Ultralight Sleeping Pad','NEMO Equipment','The Grommet','https://thegrommet.com/product/outdoors-garden/nemo-tensor-ultralight-sleeping-pad','https://nemoequipment.com','support@nemoequipment.com','★★★★','バフルデザイン・軽量490g・静音構造スリーピングパッド'],
  [332,'アウトドア・スポーツ・旅行','Jetboil Flash Cooking System','Jetboil','The Grommet','https://thegrommet.com/product/outdoors-garden/jetboil-flash-cooking-system','https://jetboil.com','support@jetboil.com','★★★★★','100秒で沸騰・一体型ゴトク・超高速アウトドア調理システム'],
  [333,'アウトドア・スポーツ・旅行','Therm-a-Rest NeoAir XLite Sleeping Pad','Therm-a-Rest','The Grommet','https://thegrommet.com/product/outdoors-garden/thermarest-neoair-xlite','https://thermarest.com','support@thermarest.com','★★★★','R値4.2・350g軽量・4シーズン対応エアマット'],
  [334,'アウトドア・スポーツ・旅行','Big Agnes Copper Spur HV UL2 Tent','Big Agnes','The Grommet','https://thegrommet.com/product/outdoors-garden/big-agnes-copper-spur-hv-ul2','https://bigagnes.com','support@bigagnes.com','★★★★★','1.4kg軽量・自立式・2人用超軽量バックパッキングテント'],
  [335,'アウトドア・スポーツ・旅行','Kammok Mantis All-in-One Hammock','Kammok','The Grommet','https://thegrommet.com/product/outdoors-garden/kammok-mantis-hammock','https://kammok.com','hello@kammok.com','★★★★','蚊帳+タープ一体型・アトランタ発オールインワンハンモック'],
  [336,'アウトドア・スポーツ・旅行','LuminAID PackLite Solar Lantern','LuminAID','The Grommet','https://thegrommet.com/product/outdoors-garden/luminaid-packlite-solar-lantern','https://luminaid.com','hello@luminaid.com','★★★★','空気で膨らむ・防水・太陽光充電インフレータブルランタン'],
  [337,'アウトドア・スポーツ・旅行','Grayl GeoPress Water Purifier','Grayl','The Grommet','https://thegrommet.com/product/outdoors-garden/grayl-geopress-water-purifier','https://thegrayl.com','support@thegrayl.com','★★★★★','押すだけ8秒浄水・ウイルス/細菌除去可能なプレス式浄水ボトル'],
  [338,'アウトドア・スポーツ・旅行','Nite Ize Radiant Rechargeable Headlamp','Nite Ize','The Grommet','https://thegrommet.com/product/outdoors-garden/nite-ize-radiant-headlamp','https://niteize.com','support@niteize.com','★★★','USB充電・防水・複数照射モードLEDヘッドランプ'],
  [339,'アウトドア・スポーツ・旅行','Yeti Roadie 24 Hard Cooler','YETI','The Grommet','https://thegrommet.com/product/outdoors-garden/yeti-roadie-24-hard-cooler','https://yeti.com','support@yeti.com','★★★★★','頑丈ローラー付き・氷持続・車載しやすいコンパクトクーラー'],
  [340,'アウトドア・スポーツ・旅行','Osprey Daylite Plus Daypack','Osprey','The Grommet','https://thegrommet.com/product/outdoors-garden/osprey-daylite-plus-daypack','https://osprey.com','support@osprey.com','★★★★','20L・ヒップベルト付き・日帰りハイキング向けデイパック'],
  [341,'アウトドア・スポーツ・旅行','Sea to Summit X-Set Collapsible Cookware','Sea to Summit','The Grommet','https://thegrommet.com/product/outdoors-garden/sea-to-summit-x-set-cookware','https://seatosummit.com','info@seatosummit.com','★★★★','シリコン折りたたみ式・収納時1/3サイズになるクックウェアセット'],
  [342,'アウトドア・スポーツ・旅行','Black Diamond Distance Carbon Z Trekking Poles','Black Diamond','The Grommet','https://thegrommet.com/product/outdoors-garden/black-diamond-distance-carbon-z','https://blackdiamondequipment.com','support@blackdiamondequipment.com','★★★★','3つ折り式・カーボン・150gの超軽量トレッキングポール'],
  [343,'アウトドア・スポーツ・旅行','Coleman Sundome Camping Tent','Coleman','The Grommet','https://thegrommet.com/product/outdoors-garden/coleman-sundome-camping-tent','https://coleman.com','support@coleman.com','★★★','WeatherTecシステム・4人用・初心者に優しい定番キャンプテント'],
  [344,'アウトドア・スポーツ・旅行','Hydro Flask Wide Mouth Insulated Bottle','Hydro Flask','The Grommet','https://thegrommet.com/product/outdoors-garden/hydro-flask-wide-mouth-bottle','https://hydroflask.com','support@hydroflask.com','★★★★★','TempShield二重壁断熱・24時間保冷・米国発定番ステンレスボトル'],

  // ── ペット用品 (4件・345-348) ──
  [345,'ペット用品','Whistle Go Explore GPS Pet Tracker','Whistle','The Grommet','https://thegrommet.com/product/pets/whistle-go-explore-gps','https://whistle.com','support@whistle.com','★★★★★','GPS+活動量+健康アラート・首輪装着型ペットトラッカー'],
  [346,'ペット用品','Fable Pets The Big Bowl','Fable Pets','The Grommet','https://thegrommet.com/product/pets/fable-pets-the-big-bowl','https://meetfable.com','hello@meetfable.com','★★★★','高栄養・厳選素材・サブスク対応プレミアムドッグフードボウルセット'],
  [347,'ペット用品','Whisker Litter-Robot 4','Whisker','The Grommet','https://thegrommet.com/product/pets/whisker-litter-robot-4','https://litter-robot.com','support@litter-robot.com','★★★★★','センサー式自動清掃・アプリ連携・多頭飼い対応全自動猫トイレ'],
  [348,'ペット用品','Modkat Flip Litter Box','Modkat','The Grommet','https://thegrommet.com/product/pets/modkat-flip-litter-box','https://modkat.com','hello@modkat.com','★★★','反転式清掃・飛び散り防止フード付きモダンデザイン猫トイレ'],

  // ── テクノロジー・ガジェット (18件・349-366) ──
  [349,'テクノロジー・ガジェット','Rabbit R1 AI Companion Device','Rabbit Inc.','The Grommet','https://thegrommet.com/product/tech/rabbit-r1-ai-companion','https://rabbit.tech','support@rabbit.tech','★★★★','LAM(大規模行動モデル)搭載・アプリ操作を代行するAIガジェット'],
  [350,'テクノロジー・ガジェット','Humane Ai Pin Wearable Projector','Humane','The Grommet','https://thegrommet.com/product/tech/humane-ai-pin','https://hu.ma.ne','support@hu.ma.ne','★★★','画面レス・胸元装着・投影ディスプレイ型AIウェアラブル'],
  [351,'テクノロジー・ガジェット','Anker Nebula Capsule 3 Laser Projector','Anker','The Grommet','https://thegrommet.com/product/tech/anker-nebula-capsule-3','https://nebula.anker.com','support@nebula.anker.com','★★★★★','手のひらサイズ・Google TV内蔵・自動フォーカス小型プロジェクター'],
  [352,'テクノロジー・ガジェット','reMarkable Paper Pro E-ink Tablet','reMarkable','The Grommet','https://thegrommet.com/product/tech/remarkable-paper-pro','https://remarkable.com','support@remarkable.com','★★★★★','カラーE Ink・紙のような書き心地・ペーパーレスノートタブレット'],
  [353,'テクノロジー・ガジェット','Flipper Zero Multi-Tool Device','Flipper Devices','The Grommet','https://thegrommet.com/product/tech/flipper-zero-multi-tool','https://flipperzero.one','support@flipperzero.one','★★★★','RFID/NFC/赤外線・セキュリティ研究向けポータブルマルチツール'],
  [354,'テクノロジー・ガジェット','Nreal Air 2 AR Glasses','XREAL','The Grommet','https://thegrommet.com/product/tech/nreal-air-2-ar-glasses','https://xreal.com','support@xreal.com','★★★★','130インチ相当仮想スクリーン・軽量120g・携帯型ARグラス'],
  [355,'テクノロジー・ガジェット','Loop Earplugs Engage Plus','Loop','The Grommet','https://thegrommet.com/product/tech/loop-earplugs-engage-plus','https://loopearplugs.com','hello@loopearplugs.com','★★★★','会話クリア・騒音のみ軽減・北欧発デザイン性耳栓'],
  [356,'テクノロジー・ガジェット','Twelve South AirFly Pro Bluetooth Transmitter','Twelve South','The Grommet','https://thegrommet.com/product/tech/twelve-south-airfly-pro','https://twelvesouth.com','support@twelvesouth.com','★★★★','機内エンタメ・ジム機器対応Bluetoothオーディオトランスミッター'],
  [357,'テクノロジー・ガジェット','Anker Soundcore Frames Audio Sunglasses','Anker Soundcore','The Grommet','https://thegrommet.com/product/tech/anker-soundcore-frames','https://soundcore.com','support@soundcore.com','★★★★','スピーカー内蔵・UV400レンズ・音楽が聴けるサングラス'],
  [358,'テクノロジー・ガジェット','Belkin Auto-Tracking Stand Pro','Belkin','The Grommet','https://thegrommet.com/product/tech/belkin-auto-tracking-stand-pro','https://belkin.com','support@belkin.com','★★★★','AI顔認識・自動追尾・ビデオ通話用回転スタンド'],
  [359,'テクノロジー・ガジェット','Rocketbook Smart Reusable Sticky Notes','Rocketbook','The Grommet','https://thegrommet.com/product/tech/rocketbook-smart-sticky-notes','https://getrocketbook.com','hello@getrocketbook.com','★★★','クラウド自動保存・貼って剥がせる再利用可能スマート付箋'],
  [360,'テクノロジー・ガジェット','Anker 737 Power Bank PowerCore 24K','Anker','The Grommet','https://thegrommet.com/product/tech/anker-737-power-bank','https://anker.com','support@anker.com','★★★★★','24,000mAh・140W出力・ノートPCも充電できる大容量モバイルバッテリー'],
  [361,'テクノロジー・ガジェット','Satechi USB-C Slim Dock for iMac','Satechi','The Grommet','https://thegrommet.com/product/tech/satechi-usb-c-slim-dock','https://satechi.com','support@satechi.com','★★★','iMacデザインに調和・USB-C/SD対応スリムドッキングステーション'],
  [362,'テクノロジー・ガジェット','Logitech MX Master 3S Wireless Mouse','Logitech','The Grommet','https://thegrommet.com/product/tech/logitech-mx-master-3s','https://logitech.com','support@logitech.com','★★★★★','8Kセンサー・静音クリック・マルチデバイス対応高性能マウス'],
  [363,'テクノロジー・ガジェット','Keychron Q1 Pro Wireless Mechanical Keyboard','Keychron','The Grommet','https://thegrommet.com/product/tech/keychron-q1-pro','https://keychron.com','support@keychron.com','★★★★','ガスケットマウント・ホットスワップ対応・カスタムメカニカルキーボード'],
  [364,'テクノロジー・ガジェット','Ecoflow River 2 Pro Portable Power Station','EcoFlow','The Grommet','https://thegrommet.com/product/tech/ecoflow-river-2-pro','https://ecoflow.com','support@ecoflow.com','★★★★★','768Wh・1600W出力・1時間フル充電対応ポータブル電源'],
  [365,'テクノロジー・ガジェット','Withings U-Scan Urine Lab Cartridge','Withings','The Grommet','https://thegrommet.com/product/tech/withings-u-scan','https://withings.com','support@withings.com','★★★★','尿検体を自動分析・健康トレンドを可視化するトイレ設置型スキャナー'],
  [366,'テクノロジー・ガジェット','Yubico YubiKey 5C NFC Security Key','Yubico','The Grommet','https://thegrommet.com/product/tech/yubico-yubikey-5c-nfc','https://yubico.com','support@yubico.com','★★★★','物理2要素認証・NFC/USB-C対応セキュリティキー'],

  // ── 美容・スキンケア (7件・367-373) ──
  [367,'美容・スキンケア','Therabody TheraFace Pro','Therabody','The Grommet','https://thegrommet.com/product/beauty/therabody-theraface-pro','https://therabody.com','support@therabody.com','★★★★★','LED光+マイクロカレント+percussive・オールインワン美顔デバイス'],
  [368,'美容・スキンケア','Solawave Radiant Renewal Skincare Wand','Solawave','The Grommet','https://thegrommet.com/product/beauty/solawave-radiant-renewal-wand','https://solawave.co','hello@solawave.co','★★★★','レッドライト+マイクロカレント+温熱・4in1携帯美顔器'],
  [369,'美容・スキンケア','Dyson Airwrap Multi-Styler','Dyson','The Grommet','https://thegrommet.com/product/beauty/dyson-airwrap-multi-styler','https://dyson.com','support@dyson.com','★★★★★','コアンダ効果・熱ダメージ抑制・巻き髪も乾かしも一台のスタイラー'],
  [370,'美容・スキンケア','Foreo Bear 2 Microcurrent Device','FOREO','The Grommet','https://thegrommet.com/product/beauty/foreo-bear-2','https://foreo.com','support@foreo.com','★★★★','T-Sonic技術・痛みなしリフトアップ・マイクロカレント美顔器'],
  [371,'美容・スキンケア','Ouai Detox Shampoo','OUAI','The Grommet','https://thegrommet.com/product/beauty/ouai-detox-shampoo','https://theouai.com','hello@theouai.com','★★★★','アップルサイダービネガー配合・毛穴/頭皮デトックスシャンプー'],
  [372,'美容・スキンケア','Necessaire The Eye Cream','Nécessaire','The Grommet','https://thegrommet.com/product/beauty/necessaire-the-eye-cream','https://necessaire.com','hello@necessaire.com','★★★★','ペプチド+ナイアシンアミド・低刺激処方アイクリーム'],
  [373,'美容・スキンケア','PMD Personal Microderm Elite','PMD Beauty','The Grommet','https://thegrommet.com/product/beauty/pmd-personal-microderm-elite','https://pmdbeauty.com','support@pmdbeauty.com','★★★★','吸引+研磨・自宅でできるマイクロダーマブレーション'],

  // ── 子供・教育 (11件・374-384) ──
  [374,'子供・教育','LeapFrog LeapStart Interactive Learning System','LeapFrog','The Grommet','https://thegrommet.com/product/family-kids/leapfrog-leapstart','https://leapfrog.com','support@leapfrog.com','★★★★','タッチペン式・音声フィードバック・幼児向けインタラクティブ学習システム'],
  [375,'子供・教育','Yoto Player Screen-Free Audio Player','Yoto','The Grommet','https://thegrommet.com/product/family-kids/yoto-player','https://yotoplay.com','hello@yotoplay.com','★★★★★','カード挿入式・画面なし・子供が自分で操作できるオーディオプレイヤー'],
  [376,'子供・教育','Osmo Genius Starter Kit','Tangible Play','The Grommet','https://thegrommet.com/product/family-kids/osmo-genius-starter-kit','https://playosmo.com','support@playosmo.com','★★★★★','タブレット+物理ピース・書く/組む/創る5ゲーム収録知育キット'],
  [377,'子供・教育','Botzees Coding Robot Building Set','Botzees','The Grommet','https://thegrommet.com/product/family-kids/botzees-coding-robot-building-set','https://botzees.com','hello@botzees.com','★★★★','ブロック組立+アプリコーディング・5歳からのSTEMロボット玩具'],
  [378,'子供・教育','Melissa & Doug Wooden Puzzle Set','Melissa & Doug','The Grommet','https://thegrommet.com/product/family-kids/melissa-doug-wooden-puzzle-set','https://melissaanddoug.com','info@melissaanddoug.com','★★★★','頑丈な木製ピース・幼児の手指発達を促すパズルセット'],
  [379,'子供・教育','Plum Print Keepsake Book Service','Plum Print','The Grommet','https://thegrommet.com/product/family-kids/plum-print-keepsake-book','https://plumprint.com','hello@plumprint.com','★★★','子供の作品を写真に撮って一冊の本にまとめるキープセイクサービス'],
  [380,'子供・教育','Little Passports World Edition Subscription','Little Passports','The Grommet','https://thegrommet.com/product/family-kids/little-passports-world-edition','https://littlepassports.com','support@littlepassports.com','★★★★★','毎月届く・世界の国を学べる子供向け地理学習サブスクボックス'],
  [381,'子供・教育','Kiwi Crate Art & Design Subscription','KiwiCo','The Grommet','https://thegrommet.com/product/family-kids/kiwi-crate-art-design','https://kiwico.com','support@kiwico.com','★★★★','月1配送・年齢別・アート&デザイン工作キットサブスク'],
  [382,'子供・教育','Codeybear Coding for Kids Book Series','Codeybear','The Grommet','https://thegrommet.com/product/family-kids/codeybear-coding-for-kids','https://codeybear.com','hello@codeybear.com','★★★','画面を使わずコーディング概念を学ぶ絵本+パズルシリーズ'],
  [383,'子供・教育','Boogie Board Jot 8.5 Writing Tablet','Boogie Board','The Grommet','https://thegrommet.com/product/family-kids/boogie-board-jot-writing-tablet','https://myboogieboard.com','support@myboogieboard.com','★★★','電池不要・ワンクリック消去・紙を使わないお絵かきタブレット'],
  [384,'子供・教育','Melissa & Doug Water Wow Coloring Pad','Melissa & Doug','The Grommet','https://thegrommet.com/product/family-kids/melissa-doug-water-wow','https://melissaanddoug.com','info@melissaanddoug.com','★★★','水だけで色が出る・繰り返し遊べるお絵かきパッド'],

  // ── ファッション・アクセサリー (11件・385-395) ──
  [385,'ファッション・アクセサリー','Ridge Wallet Slim Minimalist Wallet','Ridge Wallet','The Grommet','https://thegrommet.com/product/fashion/ridge-wallet-slim-minimalist','https://ridge.com','support@ridge.com','★★★★★','アルミプレート・RFIDブロック・ミニマリスト向けスリムウォレット'],
  [386,'ファッション・アクセサリー','Vessi Weekend Waterproof Sneakers','Vessi','The Grommet','https://thegrommet.com/product/fashion/vessi-weekend-waterproof-sneakers','https://vessi.com','hello@vessi.com','★★★★','Dyma-tex防水生地・通気性抜群・雨の日も快適スニーカー'],
  [387,'ファッション・アクセサリー','Allbirds Tree Runners','Allbirds','The Grommet','https://thegrommet.com/product/fashion/allbirds-tree-runners','https://allbirds.com','support@allbirds.com','★★★★','ユーカリ由来繊維・軽量通気性・サステナブルランニングシューズ'],
  [388,'ファッション・アクセサリー','Feetures Elite Max Cushion Socks','Feetures','The Grommet','https://thegrommet.com/product/fashion/feetures-elite-max-cushion-socks','https://feetures.com','support@feetures.com','★★★★','アーチサポート・ブリスターガード・ランナー向けクッションソックス'],
  [389,'ファッション・アクセサリー','Pact Organic Cotton Essentials','Pact','The Grommet','https://thegrommet.com/product/fashion/pact-organic-cotton-essentials','https://wearpact.com','hello@wearpact.com','★★★★','フェアトレード認証・オーガニックコットン・基本インナーウェア'],
  [390,'ファッション・アクセサリー',"Rothy's The Merino Loafer","Rothy's",'The Grommet','https://thegrommet.com/product/fashion/rothys-merino-loafer','https://rothys.com','support@rothys.com','★★★★★','再生素材+メリノウール・洗濯機OKサステナブルローファー'],
  [391,'ファッション・アクセサリー','Baggu Standard Reusable Bag','BAGGU','The Grommet','https://thegrommet.com/product/fashion/baggu-standard-reusable-bag','https://baggu.com','hello@baggu.com','★★★★','リップストップナイロン・大容量・カラフルデザインエコバッグ'],
  [392,'ファッション・アクセサリー','Article22 Peace Bomb Bracelet','Article22','The Grommet','https://thegrommet.com/product/fashion/article22-peace-bomb-bracelet','https://article22.com','info@article22.com','★★★','ラオスの不発弾アルミを再利用・フェアトレード認証ブレスレット'],
  [393,'ファッション・アクセサリー','Poppy Barley Custom Ankle Boots','Poppy Barley','The Grommet','https://thegrommet.com/product/fashion/poppy-barley-custom-ankle-boots','https://poppybarley.com','support@poppybarley.com','★★★★','サイズカスタムオーダー・エシカル製造カナダ発アンクルブーツ'],
  [394,'ファッション・アクセサリー','United By Blue Bison Fiber Beanie','United By Blue','The Grommet','https://thegrommet.com/product/fashion/united-by-blue-bison-fiber-beanie','https://unitedbyblue.com','hello@unitedbyblue.com','★★★','バイソン繊維混紡・購入ごとに海洋ゴミ回収に貢献するビーニー'],
  [395,'ファッション・アクセサリー','Parker Clay Leather Tote Bag','Parker Clay','The Grommet','https://thegrommet.com/product/fashion/parker-clay-leather-tote-bag','https://parkerclay.com','info@parkerclay.com','★★★★','エチオピア職人手作り・雇用創出プロジェクト発フルグレインレザートート'],

  // ── クリーニング・収納・整理 (5件・396-400) ──
  [396,'クリーニング・収納・整理','Grove Collaborative Reusable Storage Bags','Grove Collaborative','The Grommet','https://thegrommet.com/product/home/grove-reusable-storage-bags','https://grove.co','support@grove.co','★★★★','シリコン製・食洗機OK・ジップロック代替再利用保存袋'],
  [397,'クリーニング・収納・整理','Bluecrate Bamboo Drawer Organizer Set','Bluecrate','The Grommet','https://thegrommet.com/product/home/bluecrate-bamboo-drawer-organizer','https://bluecrateshop.com','hello@bluecrateshop.com','★★★','天然竹・モジュール式・引き出し内を仕切る収納オーガナイザーセット'],
  [398,'クリーニング・収納・整理','Simplehuman Rectangular Step Trash Can','simplehuman','The Grommet','https://thegrommet.com/product/home/simplehuman-rectangular-step-can','https://simplehuman.com','hello@simplehuman.com','★★★★','フィンガープリントプルーフ・静音クローズ・省スペース角型ゴミ箱'],
  [399,'クリーニング・収納・整理','Container Store Elfa Closet Drawer System','The Container Store','The Grommet','https://thegrommet.com/product/home/elfa-closet-drawer-system','https://containerstore.com','support@containerstore.com','★★★★','カスタマイズ自在・スウェーデン発クローゼット収納システム'],
  [400,'クリーニング・収納・整理','Full Circle Fresh Air Odor Eliminator','Full Circle Home','The Grommet','https://thegrommet.com/product/home/full-circle-fresh-air-odor-eliminator','https://fullcirclehome.com','support@fullcirclehome.com','★★★','竹炭配合・化学薬品不使用・ゴミ箱/冷蔵庫用消臭剤'],
];

module.exports = NEW_PRODUCTS_107;
