/**
 * add_200_items_14.js
 * No.3272〜3471 (200件) — The Grommet 掲載製品
 * INPUT : 海外便利グッズリスト_日本未上陸3305件_評価付.xlsx  (3271件収録)
 * OUTPUT: 海外便利グッズリスト_日本未上陸3505件_評価付.xlsx  (3471件収録)
 * ★評価基準: アップボート数 10,000+→★★★★★ / 3,000+→★★★★ / 500+→★★★ / 〜→★★★
 */

const XLSX = require('xlsx');
const path = require('path');

const DIR    = __dirname;
const INPUT  = path.join(DIR, '海外便利グッズリスト_日本未上陸3305件_評価付.xlsx');
const OUTPUT = path.join(DIR, '海外便利グッズリスト_日本未上陸3505件_評価付.xlsx');
const SENDER = 'yutorin.ino@gmail.com';

// [番号, カテゴリ, 製品名, メーカー, ECサイト, 製品URL, メーカーHP, メール, ★, コメント]
const NEW_PRODUCTS = [
  // ── キッチン・調理器具 (20件) ──
  [3272,'キッチン・調理器具','Negg Egg Peeler','Negg Egg Products','The Grommet','https://thegrommet.com/product/kitchen/negg-egg-peeler','https://thenegg.com','hello@thenegg.com','★★★★★','水と振動で殻をむく革新的ゆで卵ピーラー・米国製 (12,384 upvotes)'],
  [3273,'キッチン・調理器具','Polygons 3-in-1 Flat Measuring Spoon','Polygons Design','The Grommet','https://thegrommet.com/product/kitchen/polygons-3-in-1-flat-measuring-spoon','https://polygonsdesign.com','info@polygonsdesign.com','★★★★★','1本で大さじ・小さじ・tsp計量できるフラット収納計量スプーン (4,218 upvotes)'],
  [3274,'キッチン・調理器具','Formaticum Cheese Storage Bags','Formaticum','The Grommet','https://thegrommet.com/product/kitchen/formaticum-cheese-storage-bags','https://formaticum.com','hello@formaticum.com','★★★★★','通気性素材でチーズを最適保存・食品廃棄削減バッグ (4,863 upvotes)'],
  [3275,'キッチン・調理器具','The Whiskey Grail','Whiskey Grail','The Grommet','https://thegrommet.com/product/kitchen/the-whiskey-grail','https://whiskeygrail.com','support@whiskeygrail.com','★★★★★','スコッチグラス内蔵ミニチュア樽・ウイスキー風味熟成カップ (3,775 upvotes)'],
  [3276,'キッチン・調理器具','Microwave-Safe Stainless Steel Container Set','Magisso Design','The Grommet','https://thegrommet.com/product/kitchen/microwave-safe-stainless-steel-container-set','https://magisso.com','info@magisso.com','★★★★','電子レンジ対応ステンレス容器セット・冷蔵から温めまで一容器 (3,783 upvotes)'],
  [3277,'キッチン・調理器具','NexTrend Garlic Twister','NexTrend','The Grommet','https://thegrommet.com/product/kitchen/nextrend-garlic-twister','https://nextrend.com','hello@nextrend.com','★★★★','片手でひねるだけにんにくみじん切り・手が匂わないガーリックミンサー (3,156 upvotes)'],
  [3278,'キッチン・調理器具','Titanium Zero-Tox Cutting Board','Titanium Cuttingboard','The Grommet','https://thegrommet.com/product/kitchen/titanium-zero-tox-cutting-board','https://titaniumcuttingboard.com','support@titaniumcuttingboard.com','★★★★','マイクロプラスチックゼロ・チタン製軽量まな板 (1,983 upvotes)'],
  [3279,'キッチン・調理器具','Nextrend Baker Scraper','NexTrend','The Grommet','https://thegrommet.com/product/kitchen/nextrend-baker-scraper','https://nextrend.com','hello@nextrend.com','★★★★','生地カット・すくい・成形まで1本でこなすベーカーズスクレーパー (1,547 upvotes)'],
  [3280,'キッチン・調理器具','Micro Grill Microwave Egg Cooker','Micro Grill','The Grommet','https://thegrommet.com/product/kitchen/micro-grill-microwave-egg-cooker','https://microgrill.com','info@microgrill.com','★★★★','電子レンジで目玉焼き・電気不要・3分でプロ仕上げ卵調理器 (826 upvotes)'],
  [3281,'キッチン・調理器具','FreshPaper Produce Saver Sheets','The FRESHGLOW Co.','The Grommet','https://thegrommet.com/product/kitchen/freshpaper-produce-saver-sheets','https://thefreshglowco.com','hello@thefreshglowco.com','★★★★','有機ボタニカル配合・野菜・果物を2〜4倍長持ちさせる保存シート (941 upvotes)'],
  [3282,'キッチン・調理器具','GRIPmitt BBQ Kitchen Mitts','GRIPmitt','The Grommet','https://thegrommet.com/product/kitchen/gripmitt-bbq-kitchen-mitts','https://gripmitt.com','support@gripmitt.com','★★★★','シリコン&テリー地の両用鍋つかみ・指先グリップ付き耐熱ミット (865 upvotes)'],
  [3283,'キッチン・調理器具','Hyvance Smart Fried Egg Cooker','Hyvance','The Grommet','https://thegrommet.com/product/kitchen/hyvance-smart-fried-egg-cooker','https://hyvance.com','info@hyvance.com','★★★★','プッシュ一つで完璧な目玉焼き・自動温度管理スマート卵調理器 (993 upvotes)'],
  [3284,'キッチン・調理器具','Revomax Classic Bottle','Revomax','The Grommet','https://thegrommet.com/product/kitchen/revomax-classic-bottle','https://revomax.com','hello@revomax.com','★★★','ワンタップで開閉・高耐圧・炭酸飲料対応真空断熱ボトル'],
  [3285,'キッチン・調理器具','SLOTDOG Hot Dog Slicing Tool','SlotDog','The Grommet','https://thegrommet.com/product/kitchen/slotdog-hot-dog-slicing-tool','https://slotdog.com','support@slotdog.com','★★★','格子状切り込みでカリカリ食感・ホットドッグスライサー (303 upvotes)'],
  [3286,'キッチン・調理器具','CRIMPiT Tortilla Sealer','CRIMPiT','The Grommet','https://thegrommet.com/product/kitchen/crimpit','https://crimpit.co.uk','hello@crimpit.co.uk','★★★','トルティーヤを折り畳みクリンプする専用シール器 (406 upvotes)'],
  [3287,'キッチン・調理器具','Burger Master Burger Press','Burger Master','The Grommet','https://thegrommet.com/product/kitchen/burger-master','https://burgermastertools.com','info@burgermastertools.com','★★★','均等な厚さのパティを簡単成形・詰め物対応バーガープレス (329 upvotes)'],
  [3288,'キッチン・調理器具','Novaz Crystal Glass Flexible Cup','Novaz','The Grommet','https://thegrommet.com/product/kitchen/novaz-crystal-glass','https://novazglass.com','support@novazglass.com','★★★','割れない・柔軟なトライタン素材クリスタルルックカップ'],
  [3289,'キッチン・調理器具','Vestiges Kitchen Towels','Vestiges Home','The Grommet','https://thegrommet.com/product/kitchen/vestiges-kitchen-towels','https://vestigeshome.com','hello@vestigeshome.com','★★★','アーティスト作品プリント・飾っておきたいデザインキッチンタオル (170 upvotes)'],
  [3290,'キッチン・調理器具','Mix & Carry Stainless Steel Bowl Set','Mix & Carry','The Grommet','https://thegrommet.com/product/kitchen/mix-carry-stainless-steel-bowl-set','https://mixandcarry.com','info@mixandcarry.com','★★★','下ごしらえ・保存・サーブが1つのボウルで完結するステンレスセット'],
  [3291,'キッチン・調理器具','The Micro Grill Crispy Cooker','Micro Grill Inc.','The Grommet','https://thegrommet.com/product/kitchen/the-micro-grill','https://microgrill.com','info@microgrill.com','★★★★','電子レンジでグリル風カリカリ食感を実現するシリコントレー (2,273 upvotes)'],

  // ── スマートホーム・インテリア・照明 (20件) ──
  [3292,'スマートホーム・インテリア・照明','MagicOutlet Smart USB Outlet Cover','MagicOutlet','The Grommet','https://thegrommet.com/product/home/magicoutlet-smart-usb-outlet-cover','https://magicoutlet.com','support@magicoutlet.com','★★★★','既存コンセントを交換不要でUSB給電ポート追加スマートカバー'],
  [3293,'スマートホーム・インテリア・照明','Spiral Light Candles Self-Consuming','Spiral Light Candle Corp.','The Grommet','https://thegrommet.com/product/home/spiral-light-candles','https://spirallightcandles.com','hello@spirallightcandles.com','★★★★','燃え終わると自然消灯・ロウがこぼれないスパイラル構造キャンドル'],
  [3294,'スマートホーム・インテリア・照明','Brightech Sparq LED Arc Floor Lamp','Brightech','The Grommet','https://thegrommet.com/product/home/brightech-sparq-led-arc-floor-lamp','https://brightechled.com','support@brightechled.com','★★★★','調光3段階・省エネLED・アーチ型フロアランプ'],
  [3295,'スマートホーム・インテリア・照明','Cork Pops Wine Opener Aerator','Cork Pops','The Grommet','https://thegrommet.com/product/home/cork-pops-wine-opener','https://corkpops.com','info@corkpops.com','★★★★','ガスカートリッジ式・コルクをワンプッシュで抜栓するワインオープナー'],
  [3296,'スマートホーム・インテリア・照明','Kikkerland Retractable Luggage Tag','Kikkerland Design','The Grommet','https://thegrommet.com/product/home/kikkerland-retractable-luggage-tag','https://kikkerland.com','support@kikkerland.com','★★★','伸縮ストラップ・ユニークデザインのラゲッジタグ'],
  [3297,'スマートホーム・インテリア・照明','WRAP-IT Storage Straps','WRAP-IT Inc.','The Grommet','https://thegrommet.com/product/home/wrap-it-storage-straps','https://wrapit-inc.com','hello@wrapit-inc.com','★★★★','コード・ケーブルをスッキリまとめる再利用可能シリコンストラップ'],
  [3298,'スマートホーム・インテリア・照明','Mosen Stencil Roller Painting Kit','Mosen','The Grommet','https://thegrommet.com/product/home/mosen-stencil-roller-painting-kit','https://mosenpaint.com','info@mosenpaint.com','★★★','模様ローラーで壁紙風アクセントが手軽に作れるDIYペイントキット'],
  [3299,'スマートホーム・インテリア・照明','Candle Shack Wax Melt Burner','Candle Shack','The Grommet','https://thegrommet.com/product/home/candle-shack-wax-melt-burner','https://candleshack.com','hello@candleshack.com','★★★','炎なし・電動ウォーマー・長時間香りが広がるワックスメルトバーナー'],
  [3300,'スマートホーム・インテリア・照明','Honey-Can-Do Rolling Laundry Cart','Honey-Can-Do','The Grommet','https://thegrommet.com/product/home/honey-can-do-rolling-laundry-cart','https://honeycando.com','support@honeycando.com','★★★','3段カゴ付き・折りたたみ収納・キャスター付きランドリーカート'],
  [3301,'スマートホーム・インテリア・照明','Umbra Trigg Floating Shelf','Umbra','The Grommet','https://thegrommet.com/product/home/umbra-trigg-floating-shelf','https://umbra.com','hello@umbra.com','★★★★','三角形シェルフ+ウォールフック・ミニマルデザイン壁掛け飾り棚'],
  [3302,'スマートホーム・インテリア・照明','GreenLife Soft Grip Cookware Set','GreenLife','The Grommet','https://thegrommet.com/product/home/greenlife-soft-grip-cookware-set','https://greenlifecookware.com','support@greenlifecookware.com','★★★★','セラミックノンスティック・PFAS不使用・パステルカラー調理器具セット'],
  [3303,'スマートホーム・インテリア・照明','Canopy Humidifier Filtered','Canopy','The Grommet','https://thegrommet.com/product/home/canopy-humidifier','https://canopyhumidifier.com','hello@canopyhumidifier.com','★★★★★','抗菌フィルター内蔵・カビ発生ゼロ・スマートアプリ連携加湿器'],
  [3304,'スマートホーム・インテリア・照明','PinPoint Blinds Cordless Shade','PinPoint Blinds','The Grommet','https://thegrommet.com/product/home/pinpoint-blinds-cordless-shade','https://pinpointblinds.com','info@pinpointblinds.com','★★★','コードレス・子供安全対応・遮光カーテンシェード'],
  [3305,'スマートホーム・インテリア・照明','Bright Innovations LED Closet Light','Bright Innovations','The Grommet','https://thegrommet.com/product/home/bright-innovations-led-closet-light','https://brightinnovations.com','support@brightinnovations.com','★★★','モーションセンサー・電池式・設置簡単クローゼット用LEDライト'],
  [3306,'スマートホーム・インテリア・照明','OrganizeIt Cabinet Door Organizer','OrganizeIt','The Grommet','https://thegrommet.com/product/home/organizeit-cabinet-door-organizer','https://organizeit.com','hello@organizeit.com','★★★','扉裏に設置・スパイス/ラップ類の整理に最適なキャビネットオーガナイザー'],
  [3307,'スマートホーム・インテリア・照明','Poppin Office Accessories Set','Poppin','The Grommet','https://thegrommet.com/product/home/poppin-office-accessories','https://poppin.com','support@poppin.com','★★★★','カラフル・モジュラーデザインのデスクアクセサリーセット'],
  [3308,'スマートホーム・インテリア・照明','DockATot Deluxe+ Baby Lounger','DockATot','The Grommet','https://thegrommet.com/product/home/dockatot-deluxe-baby-lounger','https://dockatot.com','hello@dockatot.com','★★★★★','北欧発・オーガニックコットン・赤ちゃんのためのラウンジャー'],
  [3309,'スマートホーム・インテリア・照明','Sip & Spin Records Cocktail Coaster','Sip & Spin Records','The Grommet','https://thegrommet.com/product/home/sip-and-spin-records-cocktail-coaster','https://sipandspinrecords.com','info@sipandspinrecords.com','★★★','レコード型・回転する個性的なコースター'],
  [3310,'スマートホーム・インテリア・照明','Furrion Vision S Wireless RV Backup Camera','Furrion','The Grommet','https://thegrommet.com/product/home/furrion-wireless-rv-backup-camera','https://furrion.com','support@furrion.com','★★★★','RV専用・7インチモニター付きワイヤレスバックカメラシステム'],
  [3311,'スマートホーム・インテリア・照明','LUCID Comfort Collection Mattress Topper','LUCID','The Grommet','https://thegrommet.com/product/home/lucid-comfort-mattress-topper','https://lucidmattress.com','hello@lucidmattress.com','★★★★','竹炭配合メモリーフォーム・3インチ寝心地改善マットレストッパー'],

  // ── ウェアラブル・ヘルス・フィットネス (20件) ──
  [3312,'ウェアラブル・ヘルス・フィットネス','Intake Breathing Starter Kit','Intake Breathing Technology','The Grommet','https://thegrommet.com/product/health-wellness/intake-breathing-starter-kit','https://intakebreathing.com','support@intakebreathing.com','★★★★★','磁気鼻孔拡張器・平均ストリップの2倍の開口・いびき軽減 (7,701 upvotes)'],
  [3313,'ウェアラブル・ヘルス・フィットネス','LIVFRESH Dental Gel','LiveFresh','The Grommet','https://thegrommet.com/product/health-wellness/livfresh-dental-gel','https://livfresh.com','hello@livfresh.com','★★★★★','就寝中も歯を守る・フッ素不使用・酵素配合歯磨きジェル (8,621 upvotes)'],
  [3314,'ウェアラブル・ヘルス・フィットネス','Bodifresh Toilet Paper Foam','Bodifresh','The Grommet','https://thegrommet.com/product/health-wellness/bodifresh-toilet-paper-foam','https://bodifresh.com','hello@bodifresh.com','★★★★★','植物由来フォームでトイレットペーパーをウェットシート代わりに (325 upvotes)'],
  [3315,'ウェアラブル・ヘルス・フィットネス','The Body Scratcher Telescoping Back Scratcher','The Body Scratcher','The Grommet','https://thegrommet.com/product/health-wellness/the-body-scratcher','https://thebodyscratcher.com','support@thebodyscratcher.com','★★★★★','伸縮式・全身どこでも届く・柔軟先端のかゆみ解消スクラッチャー (6,338 upvotes)'],
  [3316,'ウェアラブル・ヘルス・フィットネス','Pulse & Heat Intelligent Neck Massager','Pulse & Heat','The Grommet','https://thegrommet.com/product/health-wellness/pulse-heat-intelligent-neck-massager','https://pulseandheat.com','info@pulseandheat.com','★★★★★','電気パルス+温熱・首周り360度・ワイヤレスネックマッサージャー (8,101 upvotes)'],
  [3317,'ウェアラブル・ヘルス・フィットネス','Setex Nose Pads for Eyeglasses','Setex','The Grommet','https://thegrommet.com/product/health-wellness/setex-nose-pads-for-eyeglasses','https://setextech.com','support@setextech.com','★★★★★','ゲッコーグリップ素材・メガネのズレを瞬時に解消ノーズパッド (8,894 upvotes)'],
  [3318,'ウェアラブル・ヘルス・フィットネス','Nova Hearing Aid OTC','Nova Hearing','The Grommet','https://thegrommet.com/product/health-wellness/nova-hearing-aid','https://novahearing.com','hello@novahearing.com','★★★★★','処方箋不要OTC補聴器・Bluetooth対応・スマホ音量調整 (284 upvotes)'],
  [3319,'ウェアラブル・ヘルス・フィットネス','Zen Bamboo Electric Toothbrush','Zen Bamboo','The Grommet','https://thegrommet.com/product/health-wellness/zen-bamboo-electric-toothbrush','https://zenbamboo.com','info@zenbamboo.com','★★★','竹製ハンドル・プラスチック削減・充電式電動歯ブラシ'],
  [3320,'ウェアラブル・ヘルス・フィットネス','Sting Stick Bug Bite Relief Device','Sting Stick','The Grommet','https://thegrommet.com/product/health-wellness/sting-stick-bug-bite-relief','https://stingstick.com','support@stingstick.com','★★★★','熱パルス技術・虫刺され・蚊・ハチの腫れとかゆみを即解消 (671 upvotes)'],
  [3321,'ウェアラブル・ヘルス・フィットネス','Ugokku Posture Corrector Wearable','Ugokku','The Grommet','https://thegrommet.com/product/health-wellness/ugokku-posture-corrector','https://ugokku.com','hello@ugokku.com','★★★★','背中に貼る・振動フィードバックで猫背を矯正するウェアラブル'],
  [3322,'ウェアラブル・ヘルス・フィットネス','Spandits Stretching Strap','Spandits','The Grommet','https://thegrommet.com/product/health-wellness/spandits-stretching-strap','https://spandits.com','info@spandits.com','★★★','12ループ設計・全身ストレッチ対応・理学療法士推奨ストラップ'],
  [3323,'ウェアラブル・ヘルス・フィットネス','Revive Light Therapy Lamp','Revive Light Therapy','The Grommet','https://thegrommet.com/product/health-wellness/revive-light-therapy-lamp','https://revivelighttherapy.com','support@revivelighttherapy.com','★★★★','10,000ルクス・UV除去・季節性気分障害改善ライトセラピーランプ'],
  [3324,'ウェアラブル・ヘルス・フィットネス','TriggerPoint GRID Foam Roller','TriggerPoint Performance','The Grommet','https://thegrommet.com/product/health-wellness/triggerpoint-grid-foam-roller','https://tptherapy.com','hello@tptherapy.com','★★★★','多密度グリッドデザイン・筋膜リリース・スポーツリカバリーローラー'],
  [3325,'ウェアラブル・ヘルス・フィットネス','Chirp Wheel Back Pain Relief','Chirp','The Grommet','https://thegrommet.com/product/health-wellness/chirp-wheel-back-pain-relief','https://chirpwheels.com','support@chirpwheels.com','★★★★★','背骨の間を的確に刺激・腰痛・肩こり解消ストレッチホイール'],
  [3326,'ウェアラブル・ヘルス・フィットネス','Withings Thermo Smart Temporal Thermometer','Withings','The Grommet','https://thegrommet.com/product/health-wellness/withings-thermo-thermometer','https://withings.com','support@withings.com','★★★★','16センサー・2秒測定・スマホ自動記録スマート体温計'],
  [3327,'ウェアラブル・ヘルス・フィットネス','Upright GO 2 Posture Trainer','Upright Technologies','The Grommet','https://thegrommet.com/product/health-wellness/upright-go-2-posture-trainer','https://uprightpose.com','hello@uprightpose.com','★★★★★','背中に貼る・姿勢データ記録・リアルタイム振動フィードバック'],
  [3328,'ウェアラブル・ヘルス・フィットネス','Acupoint Physical Therapy Massage Balls','Acupoint','The Grommet','https://thegrommet.com/product/health-wellness/acupoint-physical-therapy-massage-balls','https://acupoint.com','info@acupoint.com','★★★★','ピーナッツ型・スパイク型・スムース3種類のマッサージボールセット'],
  [3329,'ウェアラブル・ヘルス・フィットネス','Gua Sha Sculpting Tool Set','Gua Sha','The Grommet','https://thegrommet.com/product/health-wellness/gua-sha-sculpting-tool-set','https://guabeautytool.com','support@guabeautytool.com','★★★','翡翠・ローズクォーツ・本石製フェイスリフティンググアシャセット'],
  [3330,'ウェアラブル・ヘルス・フィットネス','Comrad Companion Compression Socks','Comrad Socks','The Grommet','https://thegrommet.com/product/health-wellness/comrad-companion-compression-socks','https://comradsocks.com','hello@comradsocks.com','★★★★','段階着圧・通気メッシュ・長時間立ち仕事・旅行対応圧迫ソックス'],
  [3331,'ウェアラブル・ヘルス・フィットネス','Curveball Hydration Reminder Water Bottle','Curveball','The Grommet','https://thegrommet.com/product/health-wellness/curveball-hydration-reminder-bottle','https://curveballbottle.com','info@curveballbottle.com','★★★','時間マーカー付き・1日の水分補給目標を可視化するボトル'],

  // ── アウトドア・スポーツ・旅行 (20件) ──
  [3332,'アウトドア・スポーツ・旅行','Portable Tire Inflator Compact','Portable Power Technology','The Grommet','https://thegrommet.com/product/auto/portable-tire-inflator','https://portabletireinflator.com','support@portabletireinflator.com','★★★★★','コードレス・自動停止・LED搭載コンパクト電動タイヤインフレーター (5,887 upvotes)'],
  [3333,'アウトドア・スポーツ・旅行','YARD-X Multi-Use Garden Tool','Yard-X','The Grommet','https://thegrommet.com/product/outdoors-garden/yard-x-multi-use-garden-tool','https://yard-x.com','hello@yard-x.com','★★★','5つの庭仕事をこなす回転ヘッド付き多機能ガーデンツール (200 upvotes)'],
  [3334,'アウトドア・スポーツ・旅行','The Beach Towel Chair Blanket','Beach Towel Chair','The Grommet','https://thegrommet.com/product/outdoors-garden/the-beach-towel-chair-blanket','https://beachtowelchair.com','info@beachtowelchair.com','★★★','風に飛ばない・砂がつかない・椅子に固定できるビーチタオル'],
  [3335,'アウトドア・スポーツ・旅行','Pops Birding Original Hummingbird Swing','Pops Birding','The Grommet','https://thegrommet.com/product/outdoors-garden/pops-birding-original-hummingbird-swing','https://popsbirding.com','hello@popsbirding.com','★★★★★','ハチドリ専用木製パーチ・フィーダーそばに設置・休憩場所を提供 (3,763 upvotes)'],
  [3336,'アウトドア・スポーツ・旅行','FunL Stretchable Automotive Funnel','FunL','The Grommet','https://thegrommet.com/product/auto/funl-stretchable-automotive-funnel','https://funlfunnel.com','support@funlfunnel.com','★★★','蛇腹伸縮式・こぼれない・全車種対応オートモーティブファンネル (235 upvotes)'],
  [3337,'アウトドア・スポーツ・旅行','Vista Vail RFID Crossbody Purse','Vista Vail','The Grommet','https://thegrommet.com/product/travel/vista-vail-rfid-crossbody-purse','https://vistavail.com','hello@vistavail.com','★★★','スマホ出し入れ不要・RFIDブロック・スマホ収納型クロスボディバッグ'],
  [3338,'アウトドア・スポーツ・旅行','Waterfly Fanny Pack Sling Bag','Waterfly','The Grommet','https://thegrommet.com/product/outdoors-garden/waterfly-fanny-pack-sling-bag','https://waterflybag.com','support@waterflybag.com','★★★★','防水・USB充電ポート付き・ハイキング向けスリングバッグ'],
  [3339,'アウトドア・スポーツ・旅行','Rhino USA Heavy Duty Tow Strap','Rhino USA','The Grommet','https://thegrommet.com/product/auto/rhino-usa-tow-strap','https://rhinousa.com','info@rhinousa.com','★★★★','31,518ポンド耐荷重・レーシングフック付きリカバリートウストラップ'],
  [3340,'アウトドア・スポーツ・旅行','GCI Outdoor Freestyle Rocker Chair','GCI Outdoor','The Grommet','https://thegrommet.com/product/outdoors-garden/gci-outdoor-freestyle-rocker','https://gcioutdoor.com','hello@gcioutdoor.com','★★★★','スプリング搭載・折りたたみ・地面どこでも揺れるキャンプロッキングチェア'],
  [3341,'アウトドア・スポーツ・旅行','Igloo Trailmate Journey Cooler','Igloo Products','The Grommet','https://thegrommet.com/product/outdoors-garden/igloo-trailmate-journey-cooler','https://igloo.com','support@igloo.com','★★★★','全地形対応タイヤ・バンジーコード付き・ドリンクホルダー装備クーラーボックス'],
  [3342,'アウトドア・スポーツ・旅行','Ovie Smarterware Freshness Tracking System','Ovie','The Grommet','https://thegrommet.com/product/travel/ovie-smarterware-freshness-tracking','https://ovie.life','hello@ovie.life','★★★★','食品鮮度を光でお知らせ・IoT対応スマートストレージシステム'],
  [3343,'アウトドア・スポーツ・旅行','BioLite Campstove 2+ Bundle','BioLite','The Grommet','https://thegrommet.com/product/outdoors-garden/biolite-campstove-2-plus','https://bioliteenergy.com','hello@bioliteenergy.com','★★★★★','枝で発電・USB充電・送風ファン内蔵キャンプストーブ'],
  [3344,'アウトドア・スポーツ・旅行','Coghlan\'s Backpacker Stove','Coghlan\'s','The Grommet','https://thegrommet.com/product/outdoors-garden/coghlans-backpacker-stove','https://coghlans.com','support@coghlans.com','★★★','折りたたみ式・固形燃料対応・超軽量バックパッカーストーブ'],
  [3345,'アウトドア・スポーツ・旅行','Frog Toggs Ultra-Lite2 Rain Suit','Frog Toggs','The Grommet','https://thegrommet.com/product/outdoors-garden/frog-toggs-ultra-lite-rain-suit','https://frogtoggs.com','info@frogtoggs.com','★★★★','超軽量170g・ポーチに収納・通気性防水レインスーツ'],
  [3346,'アウトドア・スポーツ・旅行','Eureka! Suma Camp Chair','Eureka!','The Grommet','https://thegrommet.com/product/outdoors-garden/eureka-suma-camp-chair','https://eurekacamping.com','hello@eurekacamping.com','★★★','アルミフレーム・収納バッグ付き・275ポンド耐荷重キャンプチェア'],
  [3347,'アウトドア・スポーツ・旅行','Foxelli Trekking Poles Carbon Fiber','Foxelli','The Grommet','https://thegrommet.com/product/outdoors-garden/foxelli-trekking-poles-carbon-fiber','https://foxelli.com','support@foxelli.com','★★★★','カーボンファイバー・クイックロック・軽量225gトレッキングポール'],
  [3348,'アウトドア・スポーツ・旅行','Nomader Collapsible Wine Glasses Travel Set','Nomader','The Grommet','https://thegrommet.com/product/travel/nomader-collapsible-wine-glasses','https://nomader.com','hello@nomader.com','★★★','折りたたみ・割れない・ピクニック・キャンプ用シリコンワイングラスセット'],
  [3349,'アウトドア・スポーツ・旅行','Kikkerland Solar Rechargeable Lantern','Kikkerland Design','The Grommet','https://thegrommet.com/product/outdoors-garden/kikkerland-solar-lantern','https://kikkerland.com','support@kikkerland.com','★★★','ソーラー充電・折りたたみ・アウトドア・停電対応LEDランタン'],
  [3350,'アウトドア・スポーツ・旅行','Sea to Summit Aeros Premium Pillow','Sea to Summit','The Grommet','https://thegrommet.com/product/travel/sea-to-summit-aeros-premium-pillow','https://seatosummit.com','info@seatosummit.com','★★★★','マルチファンクションバルブ・100g超軽量・エアー注入式旅行枕'],
  [3351,'アウトドア・スポーツ・旅行','Camco RhinoFLEX 20ft Sewer Hose Kit','Camco','The Grommet','https://thegrommet.com/product/auto/camco-rhinoflex-sewer-hose-kit','https://camco.net','support@camco.net','★★★','RV専用・伸縮式・臭気遮断バルブ付き6mセワーホースキット'],

  // ── ペット用品 (20件) ──
  [3352,'ペット用品','FurZapper Pet Hair Remover','FurZapper','The Grommet','https://thegrommet.com/product/pets/furzapper-pet-hair-remover','https://furzapper.com','hello@furzapper.com','★★★★★','洗濯機に入れるだけ・ペットの毛を根こそぎ除去するシリコンパッド (4,610 upvotes)'],
  [3353,'ペット用品','Uproot Washing Machine Cleaner','Uproot Clean','The Grommet','https://thegrommet.com/product/home/uproot-washing-machine-cleaner','https://uprootclean.com','support@uprootclean.com','★★★★★','洗濯槽の汚れ・ペット毛・カビを根こそぎ除去する洗濯機クリーナー (505 upvotes)'],
  [3354,'ペット用品','Cozy Cave Dog Bed Snuggle Sack','Snoozer Pet Products','The Grommet','https://thegrommet.com/product/pets/cozy-cave-dog-bed-snuggle-sack','https://snoozerpeproducts.com','info@snoozerpetproducts.com','★★★★','洞窟型・フリース素材・犬猫の隠れたい本能を満たすコージーベッド'],
  [3355,'ペット用品','PetFusion Ultimate Dog Bed','PetFusion','The Grommet','https://thegrommet.com/product/pets/petfusion-ultimate-dog-bed','https://petfusion.com','support@petfusion.com','★★★★★','メモリーフォーム・防水カバー・洗濯可能・整形外科サポートベッド'],
  [3356,'ペット用品','Outward Hound Hide-A-Squirrel Puzzle Toy','Outward Hound','The Grommet','https://thegrommet.com/product/pets/outward-hound-hide-a-squirrel-puzzle-toy','https://outwardhound.com','hello@outwardhound.com','★★★★','ぬいぐるみを樹に隠して探させるインタラクティブ犬用パズルトイ'],
  [3357,'ペット用品','Catit Flower Fountain Water Dispenser','Catit','The Grommet','https://thegrommet.com/product/pets/catit-flower-fountain-water-dispenser','https://catit.com','support@catit.com','★★★★','花型三段流水・トリプルフィルター・ねこの飲水量を増やすウォーターファウンテン'],
  [3358,'ペット用品','EzyDog Quick Fit Dog Harness','EzyDog','The Grommet','https://thegrommet.com/product/pets/ezydog-quick-fit-dog-harness','https://ezydog.com','info@ezydog.com','★★★★','5秒装着・胸部保護パッド・引っ張り防止設計の犬用ハーネス'],
  [3359,'ペット用品','Kurgo Journey Dog Backpack','Kurgo','The Grommet','https://thegrommet.com/product/pets/kurgo-journey-dog-backpack','https://kurgo.com','hello@kurgo.com','★★★','犬が自分のおやつ・用品を運ぶ・腹部クリップ付きドッグバックパック'],
  [3360,'ペット用品','SmartyKat Hot Pursuit Cat Toy','SmartyKat','The Grommet','https://thegrommet.com/product/pets/smartykat-hot-pursuit-cat-toy','https://smartykat.com','support@smartykat.com','★★★★','不規則動作・速度調整・ねこの狩猟本能を刺激する電動おもちゃ'],
  [3361,'ペット用品','Chuckit! Ultra Ball Medium','Chuckit!','The Grommet','https://thegrommet.com/product/pets/chuckit-ultra-ball','https://chuckit.com','info@chuckit.com','★★★★','天然ゴム・高弾性・ランチャー対応耐久性チューイングボール'],
  [3362,'ペット用品','Ruffwear Grip Trex Dog Boots','Ruffwear','The Grommet','https://thegrommet.com/product/pets/ruffwear-grip-trex-dog-boots','https://ruffwear.com','hello@ruffwear.com','★★★★','Vibramソール・反射素材・岩場・雪道対応アウトドア犬用ブーツ'],
  [3363,'ペット用品','Kong Classic Dog Toy Rubber','KONG Company','The Grommet','https://thegrommet.com/product/pets/kong-classic-dog-toy','https://kongcompany.com','support@kongcompany.com','★★★★★','世界的定番・おやつ充填・噛みごたえ長持ちゴム製知育玩具'],
  [3364,'ペット用品','Sleepypod Air Airline Approved Pet Carrier','Sleepypod','The Grommet','https://thegrommet.com/product/pets/sleepypod-air-airline-pet-carrier','https://sleepypod.com','info@sleepypod.com','★★★★','航空機機内持ち込み承認・クッション付き超軽量ペットキャリア'],
  [3365,'ペット用品','CANIDAE PURE Limited Ingredient Dog Food','CANIDAE','The Grommet','https://thegrommet.com/product/pets/canidae-pure-limited-ingredient-dog-food','https://canidae.com','hello@canidae.com','★★★★','7種以下の素材・グレインフリー・アレルギー対応プレミアムドッグフード'],
  [3366,'ペット用品','PupRenew Paw Balm','PupRenew','The Grommet','https://thegrommet.com/product/pets/puprenew-paw-balm','https://puprenew.com','support@puprenew.com','★★★','オーガニックシアバター・ビーズワックス配合・肉球保湿バーム'],
  [3367,'ペット用品','Radio Systems PetSafe ScoopFree Litter','Radio Systems PetSafe','The Grommet','https://thegrommet.com/product/pets/petsafe-scoopfree-self-cleaning-litter-box','https://petsafe.net','support@petsafe.net','★★★★','クリスタルトレイ・30日交換不要・自動猫トイレ'],
  [3368,'ペット用品','Dogtopia Calming Supplement Chews','Dogtopia','The Grommet','https://thegrommet.com/product/pets/dogtopia-calming-supplement-chews','https://dogtopia.com','hello@dogtopia.com','★★★','L-テアニン・カモミール配合・불안解消ドッグカームチュアブル'],
  [3369,'ペット用品','PetLibro Automatic Cat Feeder Wifi','PetLibro','The Grommet','https://thegrommet.com/product/pets/petlibro-automatic-cat-feeder-wifi','https://petlibro.com','support@petlibro.com','★★★★','WiFi・カメラ内蔵・スマホ遠隔管理可能自動ペット給餌器'],
  [3370,'ペット用品','West Paw Zogoflex Tux Treat Toy','West Paw Design','The Grommet','https://thegrommet.com/product/pets/west-paw-zogoflex-tux-treat-toy','https://westpaw.com','info@westpaw.com','★★★★','米国製・食洗機対応・エコフレンドリー素材知育犬用おもちゃ'],
  [3371,'ペット用品','Outward Hound Zip & Zoom Outdoor Dog Agility','Outward Hound','The Grommet','https://thegrommet.com/product/pets/outward-hound-zip-zoom-dog-agility','https://outwardhound.com','hello@outwardhound.com','★★★','6点セット・ポール・フープ・スラローム・自宅アジリティコースキット'],

  // ── テクノロジー・ガジェット (20件) ──
  [3372,'テクノロジー・ガジェット','Battarix Power Card Emergency Battery','Battarix','The Grommet','https://thegrommet.com/product/tech/battarix-power-card','https://battarix.com','support@battarix.com','★★★★★','財布に入るクレジットカードサイズ・8年保存可能使い捨て充電器 (6,958 upvotes)'],
  [3373,'テクノロジー・ガジェット','MagSnap Magnetic Snap Wristband','MagSnap','The Grommet','https://thegrommet.com/product/tech/magsnap-magnetic-snap-wristband','https://magsnap.com','hello@magsnap.com','★★★★★','磁気ポケット付き・ネジ・小物を手放さず持てる職人用リストバンド (4,018 upvotes)'],
  [3374,'テクノロジー・ガジェット','Soaq Ultrasonic Cleaner Professional','Soaq','The Grommet','https://thegrommet.com/product/tech/soaq-ultrasonic-cleaner','https://soaqclean.com','info@soaqclean.com','★★★★★','40kHz超音波・宝石/眼鏡/時計洗浄対応・自宅用超音波洗浄器 (5,376 upvotes)'],
  [3375,'テクノロジー・ガジェット','VaBroom Sweep Vacuum Combo','VaBroom','The Grommet','https://thegrommet.com/product/tech/vabroom-sweep-vacuum-combo','https://vabroom.com','support@vabroom.com','★★★★','ほうき先端に掃除機内蔵・かがまず吸引・2-in-1床掃除ツール (928 upvotes)'],
  [3376,'テクノロジー・ガジェット','Lavender Fields Bee Eco Wool Dryer Balls','Lavender Fields Bee','The Grommet','https://thegrommet.com/product/home/lavender-fields-bee-eco-wool-dryer-balls','https://lavenderfields.co','hello@lavenderfields.co','★★★★','純ウール・柔軟剤不要・乾燥時間20%短縮エコドライヤーボール (1,646 upvotes)'],
  [3377,'テクノロジー・ガジェット','Lit Mobile Solar Charger Panel','Lit Mobile','The Grommet','https://thegrommet.com/product/tech/lit-mobile-solar-charger-panel','https://litmobile.com','support@litmobile.com','★★★★','折りたたみ・USB-C/A対応・バックパック掛け可能ポータブルソーラー充電器'],
  [3378,'テクノロジー・ガジェット','Genie by Rocketbook Smart Notebook','Rocketbook','The Grommet','https://thegrommet.com/product/tech/rocketbook-genie-smart-notebook','https://getrocketbook.com','hello@getrocketbook.com','★★★★★','水で消えて繰り返し使えるスマートノート・Evernote/GDrive自動同期'],
  [3379,'テクノロジー・ガジェット','Joby GorillaPod 3K Flexible Tripod','Joby','The Grommet','https://thegrommet.com/product/tech/joby-gorillapod-3k-flexible-tripod','https://joby.com','info@joby.com','★★★★','3kg耐荷重・どこにでも巻き付く・柔軟足3脚トライポッド'],
  [3380,'テクノロジー・ガジェット','PopSockets PopGrip for MagSafe','PopSockets LLC','The Grommet','https://thegrommet.com/product/tech/popsockets-popgrip-magsafe','https://popsockets.com','support@popsockets.com','★★★★','MagSafe対応・交換可能トップ・スマホスタンド&グリップ'],
  [3381,'テクノロジー・ガジェット','Mighty Bright XtraFlex2 Book Light','Mighty Bright','The Grommet','https://thegrommet.com/product/tech/mighty-bright-xtraflex2-book-light','https://mightybright.com','hello@mightybright.com','★★★','2段階輝度・360°フレキシブルアーム・USB充電式読書ライト'],
  [3382,'テクノロジー・ガジェット','Quirky Pivot Power Genius Smart Strip','Quirky','The Grommet','https://thegrommet.com/product/tech/quirky-pivot-power-genius-smart-strip','https://quirky.com','support@quirky.com','★★★★','関節が曲がる・スマホ制御・大型プラグ対応スマート電源タップ'],
  [3383,'テクノロジー・ガジェット','Tile Pro Bluetooth Tracker','Tile','The Grommet','https://thegrommet.com/product/tech/tile-pro-bluetooth-tracker','https://thetileapp.com','hello@thetileapp.com','★★★★★','400m電波・音量大・IP67防水・探し物発見Bluetoothトラッカー'],
  [3384,'テクノロジー・ガジェット','Nomad Base One Max MagSafe Charger','Nomad Goods','The Grommet','https://thegrommet.com/product/tech/nomad-base-one-max-magsafe-charger','https://nomadgoods.com','support@nomadgoods.com','★★★★','アルミ+本革・15W・iPhone/AirPods同時充電MagSafeパッド'],
  [3385,'テクノロジー・ガジェット','Insta360 GO 3S Action Camera','Insta360','The Grommet','https://thegrommet.com/product/tech/insta360-go-3s-action-camera','https://insta360.com','hello@insta360.com','★★★★★','35g超軽量・4K・マウント不要どこでも撮れる小型アクションカメラ'],
  [3386,'テクノロジー・ガジェット','Anker Magnetic Portable Charger 10000mAh','Anker','The Grommet','https://thegrommet.com/product/tech/anker-magnetic-portable-charger-10000mah','https://anker.com','support@anker.com','★★★★','MagSafe対応・10000mAh・ワイヤレスマグネット式モバイルバッテリー'],
  [3387,'テクノロジー・ガジェット','Backbone One Mobile Gaming Controller','Backbone','The Grommet','https://thegrommet.com/product/tech/backbone-one-mobile-gaming-controller','https://playbackbone.com','hello@playbackbone.com','★★★★★','コンソール品質・iPhone/Android対応・パススルー充電ゲームコントローラー'],
  [3388,'テクノロジー・ガジェット','Blink Mini 2 Plug-in Smart Security Camera','Blink','The Grommet','https://thegrommet.com/product/tech/blink-mini-2-smart-security-camera','https://blinkforhome.com','support@blinkforhome.com','★★★★','1080p・カラーナイトビジョン・プラグイン小型スマートセキュリティカメラ'],
  [3389,'テクノロジー・ガジェット','Mophie Snap+ Juice Pack Mini','Mophie','The Grommet','https://thegrommet.com/product/tech/mophie-snap-juice-pack-mini','https://mophie.com','info@mophie.com','★★★★','MagSafe・5000mAh・ポケットに収まる磁気吸着型モバイルバッテリー'],
  [3390,'テクノロジー・ガジェット','Withings ScanWatch Nova Watch','Withings','The Grommet','https://thegrommet.com/product/tech/withings-scanwatch-nova','https://withings.com','support@withings.com','★★★★★','ECG/SpO2/睡眠・30日電池持ち・コネクテッドアナログスマートウォッチ'],
  [3391,'テクノロジー・ガジェット','Light Phone 3 Minimal Smartphone','Light','The Grommet','https://thegrommet.com/product/tech/light-phone-3','https://thelightphone.com','hello@thelightphone.com','★★★★','SNSなし・最小限機能・デジタルデトックスのためのシンプルスマートフォン'],

  // ── 美容・スキンケア (20件) ──
  [3392,'美容・スキンケア','The Best Hair Clip Twist Clip','The Hair Edit','The Grommet','https://thegrommet.com/product/beauty/the-best-hair-clip','https://thehairedit.com','support@thehairedit.com','★★★★★','ひねって留める・頭痛知らず・厚い髪もまとまるヘアクリップ (12,810 upvotes)'],
  [3393,'美容・スキンケア','Purely White Deluxe Teeth Whitening Kit','Purely White','The Grommet','https://thegrommet.com/product/beauty/purely-white-teeth-whitening-kit','https://purelywhite.com','hello@purelywhite.com','★★★★','LED光照射・ゼロ知覚過敏・3段階濃度ホワイトニングキット'],
  [3394,'美容・スキンケア','TwirlyT T-Shirt Style Clip','TwirlyT','The Grommet','https://thegrommet.com/product/fashion/twirly-t-shirt-clip','https://twirlyt.com','info@twirlyt.com','★★★★★','Tシャツを秒でおしゃれに変形・サイドノット風スタイルクリップ (5,930 upvotes)'],
  [3395,'美容・スキンケア','Pluxy Epil Pro Hair Remover','Pluxy','The Grommet','https://thegrommet.com/product/beauty/pluxy-epil-pro-hair-remover','https://pluxy.co','support@pluxy.co','★★★★','光脱毛+電気脱毛・2in1・痛みなし家庭用ムダ毛処理器'],
  [3396,'美容・スキンケア','Jolie Filtered Showerhead Skin Care','Jolie Skin Co.','The Grommet','https://thegrommet.com/product/beauty/jolie-filtered-showerhead','https://jolieskin.com','hello@jolieskin.com','★★★★★','塩素・重金属・34種除去・肌と髪に優しいフィルター付きシャワーヘッド'],
  [3397,'美容・スキンケア','Act+Acre Cold Processed Hair Treatment','Act+Acre','The Grommet','https://thegrommet.com/product/beauty/act-and-acre-cold-processed-hair-treatment','https://actandacre.com','info@actandacre.com','★★★★','コールドプロセス・頭皮ケア専門スカルプオイルトリートメント'],
  [3398,'美容・スキンケア','Kitsch Rice Water Shampoo Bar','Kitsch','The Grommet','https://thegrommet.com/product/beauty/kitsch-rice-water-shampoo-bar','https://mykitsch.com','support@mykitsch.com','★★★★','米水・ケラチン配合・1本で80回分プラスチックフリーシャンプーバー'],
  [3399,'美容・スキンケア','Moon Juice Hair Dust Supplement','Moon Juice','The Grommet','https://thegrommet.com/product/beauty/moon-juice-hair-dust','https://moonjuice.com','hello@moonjuice.com','★★★★','アシュワガンダ・サビシア配合・内側からケアするヘアサプリパウダー'],
  [3400,'美容・スキンケア','Hims Thick Fix Volumizing Shampoo','Hims','The Grommet','https://thegrommet.com/product/beauty/hims-thick-fix-volumizing-shampoo','https://forhims.com','support@forhims.com','★★★','ビオチン・ソウパルメット配合・抜け毛ケア男性向けボリュームシャンプー'],
  [3401,'美容・スキンケア','GoodLight Lash Lift Kit','GoodLight','The Grommet','https://thegrommet.com/product/beauty/goodlight-lash-lift-kit','https://goodlightco.com','info@goodlightco.com','★★★★','サロン品質・まつげパーマ・ビューラー不要DIYラッシュリフトキット'],
  [3402,'美容・スキンケア','Pai Skincare Light Work Rosehip Cleansing Oil','Pai Skincare','The Grommet','https://thegrommet.com/product/beauty/pai-skincare-rosehip-cleansing-oil','https://paiskincare.com','hello@paiskincare.com','★★★★','オーガニック認証・ローズヒップ・敏感肌向けクレンジングオイル'],
  [3403,'美容・スキンケア','Briogeo Don\'t Despair Repair Hair Mask','Briogeo','The Grommet','https://thegrommet.com/product/beauty/briogeo-dont-despair-repair-hair-mask','https://briogeohair.com','support@briogeohair.com','★★★★★','ローズヒップ+アルガン+ビタミンB5・深部補修ヘアマスク'],
  [3404,'美容・スキンケア','Crystal Deodorant Mineral Salt Stone','Crystal Deodorant','The Grommet','https://thegrommet.com/product/beauty/crystal-deodorant-mineral-salt-stone','https://thecrystal.com','info@thecrystal.com','★★★','天然ミネラル岩塩・アルミニウムフリー・肌に優しいデオドラントストーン'],
  [3405,'美容・スキンケア','LANEIGE Lip Sleeping Mask Berry','LANEIGE','The Grommet','https://thegrommet.com/product/beauty/laneige-lip-sleeping-mask-berry','https://laneige.com','support@laneige.com','★★★★★','睡眠中に唇をケア・ビタミンC誘導体・韓国発大人気リップマスク'],
  [3406,'美容・スキンケア','Sunday Riley Good Genes All-In-One Lactic Acid','Sunday Riley','The Grommet','https://thegrommet.com/product/beauty/sunday-riley-good-genes-lactic-acid','https://sundayriley.com','hello@sundayriley.com','★★★★','乳酸・精製リコリス・毛穴・くすみ即効改善ラクティックアシッドセラム'],
  [3407,'美容・スキンケア','Glow Recipe Watermelon Glow Sleeping Mask','Glow Recipe','The Grommet','https://thegrommet.com/product/beauty/glow-recipe-watermelon-sleeping-mask','https://glowrecipe.com','support@glowrecipe.com','★★★★★','スイカ抽出物・AHA・ヒアルロン酸・寝ている間に透明感UP'],
  [3408,'美容・スキンケア','Touchland Power Mist Hydrating Hand Sanitizer','Touchland','The Grommet','https://thegrommet.com/product/beauty/touchland-power-mist-hand-sanitizer','https://touchland.com','info@touchland.com','★★★★','ヒアルロン酸・アロエベラ・保湿しながら除菌するミスト型消毒液'],
  [3409,'美容・スキンケア','Slip Pure Silk Scrunchie Set','Slip','The Grommet','https://thegrommet.com/product/beauty/slip-pure-silk-scrunchie-set','https://slipsilkpillowcase.com','hello@slipsilkpillowcase.com','★★★★','22匁シルク・髪へのダメージゼロ・6色セットシルクシュシュ'],
  [3410,'美容・スキンケア','Revlon One-Step Volumizer Hair Dryer','Revlon','The Grommet','https://thegrommet.com/product/beauty/revlon-one-step-volumizer-hair-dryer','https://revlon.com','support@revlon.com','★★★★★','ドライ+スタイリング+ボリューム・3工程を1台で完結するヘアドライヤーブラシ'],
  [3411,'美容・スキンケア','Farmacy Green Clean Makeup Meltaway Cleansing Balm','Farmacy Beauty','The Grommet','https://thegrommet.com/product/beauty/farmacy-green-clean-cleansing-balm','https://farmacybeauty.com','hello@farmacybeauty.com','★★★★','ひまわりワックス・プロバイオティクス配合・W洗顔不要クレンジングバーム'],

  // ── 子供・教育 (20件) ──
  [3412,'子供・教育','Fold & Scoop Toddler Feeding Tool','Fold & Scoop','The Grommet','https://thegrommet.com/product/family-kids/fold-and-scoop','https://foldandscoop.com','support@foldandscoop.com','★★★','皿を折り曲げてスプーンでスコープしやすくする幼児食事サポートプレート'],
  [3413,'子供・教育','Learning Resources Botley Coding Robot','Learning Resources','The Grommet','https://thegrommet.com/product/family-kids/learning-resources-botley-coding-robot','https://learningresources.com','hello@learningresources.com','★★★★★','画面なし・コーディングカード操作・5歳からのプログラミングロボット'],
  [3414,'子供・教育','National Geographic Mega Fossil Dig Kit','National Geographic','The Grommet','https://thegrommet.com/product/family-kids/national-geographic-mega-fossil-dig-kit','https://natgeotoys.com','support@natgeotoys.com','★★★★','石膏ブロックを掘って恐竜の化石を発掘・本物化石入り発掘キット'],
  [3415,'子供・教育','Crayola Scribble Scrubbie Pets Toy','Crayola','The Grommet','https://thegrommet.com/product/family-kids/crayola-scribble-scrubbie-pets','https://crayola.com','info@crayola.com','★★★','ぬりぬり→洗って→またぬれる・水で落とせる繰り返し塗り絵ペットトイ'],
  [3416,'子供・教育','MindWare Keva Contraptions Plank Set','MindWare','The Grommet','https://thegrommet.com/product/family-kids/mindware-keva-contraptions-plank-set','https://mindware.com','hello@mindware.com','★★★★','200枚の木製プランクと球を組み合わせてビー玉コースを作るSTEM玩具'],
  [3417,'子供・教育','Banagrams Word Tile Game','Bananagrams Inc.','The Grommet','https://thegrommet.com/product/family-kids/bananagrams-word-tile-game','https://bananagrams.com','support@bananagrams.com','★★★★★','バナナ形ポーチ・ボード不要・速攻系クロスワードタイルゲーム'],
  [3418,'子供・教育','Stomp Rocket Ultra Rocket Launcher','Stomp Rocket','The Grommet','https://thegrommet.com/product/family-kids/stomp-rocket-ultra-rocket-launcher','https://stomprocket.com','info@stomprocket.com','★★★★','踏むだけで100フィート飛ぶ電池不要エアロケットランチャー'],
  [3419,'子供・教育','SmartLab Toys Smart Circuits Electronics Kit','SmartLab Toys','The Grommet','https://thegrommet.com/product/family-kids/smartlab-smart-circuits-electronics-kit','https://smartlabtoys.com','hello@smartlabtoys.com','★★★★','30回路・光/音/動き実験・子供向けエレクトロニクス入門キット'],
  [3420,'子供・教育','Moonlite Gift Pack Storybook Projector','Moonlite','The Grommet','https://thegrommet.com/product/family-kids/moonlite-gift-pack-storybook-projector','https://moonlite.io','support@moonlite.io','★★★★','スマホのライトで壁に絵本を投影・音声読み聞かせ対応ストーリープロジェクター'],
  [3421,'子供・教育','GeoMag Mechanics Gravity Motor Set','GeoMag','The Grommet','https://thegrommet.com/product/family-kids/geomag-mechanics-gravity-motor-set','https://geomag.com','info@geomag.com','★★★★','磁石+スチールバー・引力で動くモーター内蔵マグネット建設セット'],
  [3422,'子供・教育','Highlights Puzzlemania Subscription Box','Highlights for Children','The Grommet','https://thegrommet.com/product/family-kids/highlights-puzzlemania-subscription-box','https://highlights.com','hello@highlights.com','★★★','月1冊・クロスワード・迷路・なぞなぞ入り子供向けパズル雑誌定期便'],
  [3423,'子供・教育','Crazy Forts Fort Building Kit','Crazy Forts','The Grommet','https://thegrommet.com/product/family-kids/crazy-forts-fort-building-kit','https://crazyforts.com','support@crazyforts.com','★★★★','ボール&スティック連結式・秘密基地・室内テント簡単組み立てキット'],
  [3424,'子供・教育','Melissa & Doug Wooden Abacus','Melissa & Doug','The Grommet','https://thegrommet.com/product/family-kids/melissa-and-doug-wooden-abacus','https://melissaanddoug.com','info@melissaanddoug.com','★★★★','10行×10玉・堅牢木製フレーム・1〜100が学べる算数用そろばん'],
  [3425,'子供・教育','Educational Insights Kanoodle Puzzle Game','Educational Insights','The Grommet','https://thegrommet.com/product/family-kids/educational-insights-kanoodle-puzzle-game','https://educationalinsights.com','support@educationalinsights.com','★★★★★','1人用200問収録・3Dソリティアパズル・携帯ゲーム'],
  [3426,'子供・教育','Fat Brain Toys Tobbles Neo Stacking Toy','Fat Brain Toys','The Grommet','https://thegrommet.com/product/family-kids/fat-brain-toys-tobbles-neo','https://fatbraintoys.com','hello@fatbraintoys.com','★★★★','球面バランス・スタッキングトイ・赤ちゃんから大人まで楽しめる知育玩具'],
  [3427,'子供・教育','Osmo Pizza Co. Math Game','Tangible Play','The Grommet','https://thegrommet.com/product/family-kids/osmo-pizza-co-math-game','https://playosmo.com','support@playosmo.com','★★★★★','タブレット+物理コイン・ピザ屋経営で算数・お金の概念を学ぶゲーム'],
  [3428,'子供・教育','MBI Big Kids Coloring Book Set','My Big Intellect','The Grommet','https://thegrommet.com/product/family-kids/my-big-intellect-coloring-book-set','https://mybigintellect.com','info@mybigintellect.com','★★★','解説付き・世界の名画・建築をなぞる大人も楽しめる教育ぬりえセット'],
  [3429,'子供・教育','Jacks Games Magnetic Fishing Pond Game','Jacks Games','The Grommet','https://thegrommet.com/product/family-kids/jacks-games-magnetic-fishing-pond','https://jacksgames.com','hello@jacksgames.com','★★★','磁石釣り竿・カラフル木製魚・2〜5歳向け知育フィッシングゲーム'],
  [3430,'子供・教育','Playmonster My First Chores Kit','PlayMonster','The Grommet','https://thegrommet.com/product/family-kids/playmonster-my-first-chores-kit','https://playmonster.com','support@playmonster.com','★★★','マグネット式お手伝いチャートで責任感と達成感を育む知育グッズ'],
  [3431,'子供・教育','Snap Circuits Jr. Electronics Kit','Elenco Electronics','The Grommet','https://thegrommet.com/product/family-kids/snap-circuits-jr-electronics-kit','https://elenco.com','info@elenco.com','★★★★★','100プロジェクト・スナップ式部品・電子回路の仕組みを学ぶSTEMキット'],

  // ── ファッション・アクセサリー (20件) ──
  [3432,'ファッション・アクセサリー','Vista Vail RFID Crossbody Phone Purse','Vista Vail','The Grommet','https://thegrommet.com/product/fashion/vista-vail-rfid-crossbody-purse','https://vistavail.com','hello@vistavail.com','★★★','RFIDブロック・スマホ収納特化・片手で支払いできるクロスボディバッグ'],
  [3433,'ファッション・アクセサリー','Origami Owl Living Locket Necklace','Origami Owl','The Grommet','https://thegrommet.com/product/fashion/origami-owl-living-locket-necklace','https://origamiowl.com','support@origamiowl.com','★★★★','チャームを入れ替えてカスタムできる・想い出を閉じ込めるロケットネックレス'],
  [3434,'ファッション・アクセサリー','Birdies Flat Shoes The Starling','Birdies','The Grommet','https://thegrommet.com/product/fashion/birdies-flat-shoes-starling','https://birdiesshoes.com','info@birdiesshoes.com','★★★★★','整形外科クッション・おしゃれフラット・1日中快適に歩けるシューズ'],
  [3435,'ファッション・アクセサリー','Quay High Key Mini Sunglasses','Quay Australia','The Grommet','https://thegrommet.com/product/fashion/quay-high-key-mini-sunglasses','https://quayaustralia.com','hello@quayaustralia.com','★★★★','UV400・ポリカーボネートレンズ・トレンド感満点スモールフレームサングラス'],
  [3436,'ファッション・アクセサリー','Kizik Hands-Free Slip-On Shoes','Kizik','The Grommet','https://thegrommet.com/product/fashion/kizik-hands-free-slip-on-shoes','https://kizik.com','support@kizik.com','★★★★★','かがまず履ける・特許機構・スニーカー感覚スリッポンシューズ'],
  [3437,'ファッション・アクセサリー','Tommy John Second Skin Underwear','Tommy John','The Grommet','https://thegrommet.com/product/fashion/tommy-john-second-skin-underwear','https://tommyjohn.com','info@tommyjohn.com','★★★★','マイクロモダール・360°ストレッチ・ずり上がらないセカンドスキン下着'],
  [3438,'ファッション・アクセサリー','Corkor Cork Leather Wallet','Corkor','The Grommet','https://thegrommet.com/product/fashion/corkor-cork-leather-wallet','https://corkor.com','hello@corkor.com','★★★★','天然コルク素材・ヴィーガンレザー・軽量RFIDブロック財布'],
  [3439,'ファッション・アクセサリー','Lucky Brand 2-in-1 Convertible Puffer Vest','Lucky Brand','The Grommet','https://thegrommet.com/product/fashion/lucky-brand-convertible-puffer-vest','https://luckybrand.com','support@luckybrand.com','★★★','ベスト⇔ジャケット変換・軽量防風・2-in-1コンバーチブルダウン'],
  [3440,'ファッション・アクセサリー','Harvest Label Roll Top Backpack','Harvest Label','The Grommet','https://thegrommet.com/product/fashion/harvest-label-roll-top-backpack','https://harvestlabel.com','info@harvestlabel.com','★★★★','バリスティックナイロン・ロールトップ・米国アウトドアメーカー製バックパック'],
  [3441,'ファッション・アクセサリー','NOMATIC 20L Travel Pack','NOMATIC','The Grommet','https://thegrommet.com/product/fashion/nomatic-20l-travel-pack','https://nomatic.com','hello@nomatic.com','★★★★★','20L・40+機能ポケット・デジタルノマド向け最高峰旅行バックパック'],
  [3442,'ファッション・アクセサリー','Topo Designs Rover Pack Mini','Topo Designs','The Grommet','https://thegrommet.com/product/fashion/topo-designs-rover-pack-mini','https://topodesigns.com','support@topodesigns.com','★★★★','コロラド発・レトロデザイン・テックウェビング仕様の日常使いデイパック'],
  [3443,'ファッション・アクセサリー','Dagne Dover Landon Carryall Tote','Dagne Dover','The Grommet','https://thegrommet.com/product/fashion/dagne-dover-landon-carryall-tote','https://dagnedover.com','info@dagnedover.com','★★★★','ネオプレン素材・整理しやすい・MacBook対応・女性向け多機能トートバッグ'],
  [3444,'ファッション・アクセサリー','Bellroy Hide & Seek Leather Wallet','Bellroy','The Grommet','https://thegrommet.com/product/fashion/bellroy-hide-seek-leather-wallet','https://bellroy.com','support@bellroy.com','★★★★★','薄型・隠しキャッシュポケット・高品質レザースリムウォレット'],
  [3445,'ファッション・アクセサリー','Cotopaxi Fuego 18L Hooded Down Jacket','Cotopaxi','The Grommet','https://thegrommet.com/product/fashion/cotopaxi-fuego-18l-hooded-down-jacket','https://cotopaxi.com','hello@cotopaxi.com','★★★★','余り布使用・各個体ユニーク・責任ある製造アウトドアダウンジャケット'],
  [3446,'ファッション・アクセサリー','Rothy\'s The Point Ballet Flat','Rothy\'s','The Grommet','https://thegrommet.com/product/fashion/rothys-the-point-ballet-flat','https://rothys.com','support@rothys.com','★★★★★','ペットボトル再生糸・洗濯機OK・サステナブルバレエフラットシューズ'],
  [3447,'ファッション・アクセサリー','Allbirds Wool Runner Go','Allbirds','The Grommet','https://thegrommet.com/product/fashion/allbirds-wool-runner-go','https://allbirds.com','info@allbirds.com','★★★★★','メリノウール・カーボンニュートラル・サステナブルスニーカー'],
  [3448,'ファッション・アクセサリー','Manduka PRO Yoga Mat 6mm','Manduka','The Grommet','https://thegrommet.com/product/fashion/manduka-pro-yoga-mat','https://manduka.com','hello@manduka.com','★★★★','密度保証・生涯保証・プロ品質6mm厚グリップヨガマット'],
  [3449,'ファッション・アクセサリー','Sunday Afternoons Adventure Hat UPF 50+','Sunday Afternoons','The Grommet','https://thegrommet.com/product/fashion/sunday-afternoons-adventure-hat','https://sundayafternoons.com','support@sundayafternoons.com','★★★★','UPF50+・ネック日よけ付き・アウトドア最強紫外線カットハット'],
  [3450,'ファッション・アクセサリー','Darn Tough Vermont Merino Wool Socks','Darn Tough Vermont','The Grommet','https://thegrommet.com/product/fashion/darn-tough-vermont-merino-wool-socks','https://darntough.com','info@darntough.com','★★★★★','生涯保証・メリノウール・バーモント製最強耐久性ハイキングソックス'],
  [3451,'ファッション・アクセサリー','Anew Climate Recycled Polyester Fleece Jacket','Anew Climate','The Grommet','https://thegrommet.com/product/fashion/anew-climate-recycled-fleece-jacket','https://anewclimate.com','hello@anewclimate.com','★★★','再生ポリエステル100%・カーボンネガティブ認証フリースジャケット'],

  // ── クリーニング・収納・整理 (20件) ──
  [3452,'クリーニング・収納・整理','Uproot Lint Changer Pet Hair Remover Roller','Uproot Clean','The Grommet','https://thegrommet.com/product/home/uproot-lint-changer-pet-hair-remover','https://uprootclean.com','support@uprootclean.com','★★★★★','詰め替えテープ不要・逆方向でリセット・繰り返し使えるペット毛ローラー (505 upvotes)'],
  [3453,'クリーニング・収納・整理','Zulay Kitchen Handheld Milk Frother','Zulay Kitchen','The Grommet','https://thegrommet.com/product/kitchen/zulay-handheld-milk-frother','https://zulaykitchen.com','hello@zulaykitchen.com','★★★★','高速回転・19秒でラテ泡立て・コードレス電動ミルクフォーマー'],
  [3454,'クリーニング・収納・整理','Full Circle Suds Up Liquid Soap Dispenser','Full Circle Home','The Grommet','https://thegrommet.com/product/home/full-circle-suds-up-soap-dispenser','https://fullcirclehome.com','support@fullcirclehome.com','★★★','内蔵ポンプ+スポンジホルダー一体型・シンク周り整理ソープディスペンサー'],
  [3455,'クリーニング・収納・整理','Foaming Hand Soap Starter Kit','Blueland','The Grommet','https://thegrommet.com/product/home/blueland-foaming-hand-soap-starter-kit','https://blueland.com','info@blueland.com','★★★★','錠剤を水に溶かすだけ・プラスチック削減・エコフォームハンドソープキット'],
  [3456,'クリーニング・収納・整理','The Pink Stuff Cleaning Paste','Star Brands','The Grommet','https://thegrommet.com/product/home/the-pink-stuff-cleaning-paste','https://thepinkstuffcleaner.com','hello@thepinkstuffcleaner.com','★★★★★','SNS話題・頑固汚れを驚異的に落とす英国発万能ピンク洗浄ペースト'],
  [3457,'クリーニング・収納・整理','Grove Co. Cleaning Kit Starter Set','Grove Collaborative','The Grommet','https://thegrommet.com/product/home/grove-co-cleaning-starter-kit','https://grove.co','support@grove.co','★★★★','天然由来成分・詰め替え対応・プラスチックフリー掃除スターターセット'],
  [3458,'クリーニング・収納・整理','Scrubbing Bubbles Continuous Clean Toilet System','SC Johnson','The Grommet','https://thegrommet.com/product/home/scrubbing-bubbles-continuous-clean-toilet','https://scrabbingbubbles.com','info@scjohnson.com','★★★','毎回流すたびに洗浄・スタンプ式タンク取付型トイレ継続洗浄システム'],
  [3459,'クリーニング・収納・整理','Neat Freak Collapsible Dish Rack','Neat Freak','The Grommet','https://thegrommet.com/product/home/neat-freak-collapsible-dish-rack','https://neat-freak.com','hello@neat-freak.com','★★★','使わない時は薄く折りたたむ・シリコン脚・省スペース折りたたみ水切りラック'],
  [3460,'クリーニング・収納・整理','Casabella Perfecta Spin Mop System','Casabella','The Grommet','https://thegrommet.com/product/home/casabella-perfecta-spin-mop','https://casabella.com','support@casabella.com','★★★★','遠心分離脱水・360°ヘッド・フットペダル操作スピンモップシステム'],
  [3461,'クリーニング・収納・整理','Bamboozle Food Composter Indoor Bin','Bamboozle','The Grommet','https://thegrommet.com/product/home/bamboozle-food-composter-indoor-bin','https://bamboozleproduct.com','info@bamboozleproduct.com','★★★★','竹繊維製・カーボンフィルター内蔵・臭いが漏れないキッチン生ゴミコンポストビン'],
  [3462,'クリーニング・収納・整理','Simple Human Sensor Pump with Soap','simplehuman','The Grommet','https://thegrommet.com/product/home/simplehuman-sensor-pump-with-soap','https://simplehuman.com','hello@simplehuman.com','★★★★★','タッチフリー・USB充電・高精度センサー式ソープディスペンサー'],
  [3463,'クリーニング・収納・整理','Joseph Joseph Tota Laundry Separator','Joseph Joseph','The Grommet','https://thegrommet.com/product/home/joseph-joseph-tota-laundry-separator','https://josephjoseph.com','support@josephjoseph.com','★★★★','3仕切り分類・回転ローラー付き・スタイリッシュランドリーバスケット'],
  [3464,'クリーニング・収納・整理','Bees Wrap Reusable Food Wrap','Bee\'s Wrap','The Grommet','https://thegrommet.com/product/home/bees-wrap-reusable-food-wrap','https://beeswrap.com','hello@beeswrap.com','★★★★','ミツロウコーティング・食品保存ラップ代替・洗って繰り返し使えるエコラップ'],
  [3465,'クリーニング・収納・整理','Brabantia Bo Pedal Bin','Brabantia','The Grommet','https://thegrommet.com/product/home/brabantia-bo-pedal-bin','https://brabantia.com','info@brabantia.com','★★★★','ソフトクローズ・10年保証・フィンガープリントフリーステンレスペダルゴミ箱'],
  [3466,'クリーニング・収納・整理','Rubbermaid Brilliance Food Storage Set','Rubbermaid','The Grommet','https://thegrommet.com/product/home/rubbermaid-brilliance-food-storage-set','https://rubbermaid.com','support@rubbermaid.com','★★★★','100%液漏れなし・電子レンジOK・スタック収納トライタン製フードコンテナセット'],
  [3467,'クリーニング・収納・整理','Umbra Trigg Pegboard Organizer','Umbra','The Grommet','https://thegrommet.com/product/home/umbra-trigg-pegboard-organizer','https://umbra.com','hello@umbra.com','★★★','カスタムレイアウト・ウォールナット/白2色・壁掛けペグボードオーガナイザー'],
  [3468,'クリーニング・収納・整理','OXO Good Grips Clip-on Strainer','OXO','The Grommet','https://thegrommet.com/product/home/oxo-good-grips-clip-on-strainer','https://oxo.com','info@oxo.com','★★★★','シリコンクリップ・折りたたみハンドル・鍋に直接挟むクリップオンストレーナー'],
  [3469,'クリーニング・収納・整理','SpaceSaver Vacuum Storage Bags','SpaceSaver','The Grommet','https://thegrommet.com/product/home/spacesaver-vacuum-storage-bags','https://spacemax.com','support@spacemax.com','★★★★★','ポンプ不要・ロールアップで圧縮・衣類/布団収納バキュームバッグセット'],
  [3470,'クリーニング・収納・整理','Full Circle Neat Nut Walnut Shell Scrubber','Full Circle Home','The Grommet','https://thegrommet.com/product/home/full-circle-neat-nut-walnut-shell-scrubber','https://fullcirclehome.com','support@fullcirclehome.com','★★★','クルミ殻粉砕研磨・傷つけない・サステナブル素材ノンスクラッチスポンジ'],
  [3471,'クリーニング・収納・整理','Squatty Potty Original Toilet Stool','Squatty Potty','The Grommet','https://thegrommet.com/product/home/squatty-potty-original-toilet-stool','https://squattypotty.com','hello@squattypotty.com','★★★★★','排便姿勢を最適化・医師推奨・白色7/9インチトイレ用足台 (TV Shark Tank出演)'],
];

// ── ユーティリティ ──────────────────────────────────────────────────────────
function buildGmailUrl(toEmail, maker, product) {
  const enc = s => encodeURIComponent(s);
  const subject = `Potential Distribution Partnership for Japan - ${product}`;
  const body =
`Dear ${maker} Team,

I hope this message finds you well.

My name is Yutorin Ino, and I represent LEAGUE, a Japanese company specializing in bringing innovative overseas products to the Japanese market.

We recently discovered your product "${product}" on The Grommet and were truly impressed by its innovation and quality. We believe it has significant potential in Japan, where consumers are enthusiastic about cutting-edge products from abroad.

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
const wb_in = XLSX.readFile(INPUT);
const ws_in = wb_in.Sheets[wb_in.SheetNames[0]];
const existing = XLSX.utils.sheet_to_json(ws_in, { header: 1 });
console.log(`既存行数: ${existing.length - 1} 件`);

const allRows = [...existing];
for (const p of NEW_PRODUCTS) allRows.push(p);

const wb_out = XLSX.utils.book_new();
const ws_out = XLSX.utils.aoa_to_sheet(allRows);

ws_out['!cols'] = [
  { wch: 6 }, { wch: 18 }, { wch: 40 }, { wch: 20 },
  { wch: 14 }, { wch: 50 }, { wch: 35 }, { wch: 35 },
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
const catCount = {};
for (const p of NEW_PRODUCTS) catCount[p[1]] = (catCount[p[1]] || 0) + 1;

console.log('\n====== 追加完了レポート ======');
console.log(`追加件数    : ${NEW_PRODUCTS.length} 件`);
console.log(`メールあり  : ${mailCount} 件`);
console.log('\nカテゴリ別:');
Object.entries(catCount).forEach(([c,n]) => console.log(`  ${c}: ${n}件`));
console.log(`\n総収録件数  : ${existing.length - 1 + NEW_PRODUCTS.length} 件`);
console.log(`出力ファイル: ${OUTPUT}`);
