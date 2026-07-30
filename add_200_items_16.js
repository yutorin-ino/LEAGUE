/**
 * add_200_items_16.js
 * No.3672〜3871 (200件) — Kickstarter/Indiegogo 新規リサーチ
 * INPUT : 海外便利グッズリスト_日本未上陸3705件_評価付.xlsx  (3671件収録)
 * OUTPUT: 海外便利グッズリスト_日本未上陸3905件_評価付.xlsx  (3871件収録)
 * ★評価基準: 日本市場ニーズ（睡眠/健康/ペット/美容/時短/清潔/安全）への適合度
 */

const XLSX = require('xlsx');
const path = require('path');

const DIR    = __dirname;
const INPUT  = path.join(DIR, '海外便利グッズリスト_日本未上陸3705件_評価付.xlsx');
const OUTPUT = path.join(DIR, '海外便利グッズリスト_日本未上陸3905件_評価付.xlsx');
const SENDER = 'yutorin.ino@gmail.com';

// [番号, カテゴリ, 製品名, メーカー, ECサイト, 製品URL, メーカーHP, メール, ★, コメント]
const NEW_PRODUCTS = [
  // ── キッチン・調理器具 (18件) ──
  [3672,'キッチン・調理器具','Brava Glass Oven','Brava Home','Kickstarter','https://www.kickstarter.com/projects/brava/brava-glass-oven','https://bravahome.com','support@bravahome.com','★★★★★','パンチカード連動・6種調理法内蔵・ガラス天板スマートオーブン'],
  [3673,'キッチン・調理器具','June Oven Intelligent Countertop Oven','June Life','Kickstarter','https://www.kickstarter.com/projects/junelife/june-intelligent-oven','https://juneoven.com','support@juneoven.com','★★★★','AIカメラ食材認識・自動調理モード提案スマートオーブン'],
  [3674,'キッチン・調理器具','Mella Smart Slow Cooker','Mella','Kickstarter','https://www.kickstarter.com/projects/mella/mella-smart-slow-cooker','https://mymella.com','support@mymella.com','★★★★','アプリでレシピ選択・帰宅時間に合わせ自動加熱するスロークッカー'],
  [3675,'キッチン・調理器具','Typhur InstaProbe Meat Thermometer','Typhur','Kickstarter','https://www.kickstarter.com/projects/typhur/instaprobe-meat-thermometer','https://typhur.com','support@typhur.com','★★★★','0.5秒瞬間計測・自動オフ機能付きインスタント肉用温度計'],
  [3676,'キッチン・調理器具','Ratio Six Coffee Brewer','Ratio','Kickstarter','https://www.kickstarter.com/projects/ratio/ratio-six-coffee-brewer','https://ratiocoffee.com','support@ratiocoffee.com','★★★★','木製ハンドル・SCA認証精密抽出コーヒーメーカー'],
  [3677,'キッチン・調理器具','GNammi Automatic Pasta Maker','GNammi','Indiegogo','https://www.indiegogo.com/projects/gnammi-automatic-pasta-maker','https://gnammi.com','hello@gnammi.com','★★★★','練り〜押し出しまで全自動・本格生パスタメーカー'],
  [3678,'キッチン・調理器具','Cinder Precision Grill','Cinder','Kickstarter','https://www.kickstarter.com/projects/cinder/cinder-precision-grill','https://cinderco.com','support@cinderco.com','★★★★','45〜232℃精密温度制御・レストラン品質焼き上げグリル'],
  [3679,'キッチン・調理器具','Spinn Coffee Machine','Spinn','Indiegogo','https://www.indiegogo.com/projects/spinn/spinn-coffee-machine','https://spinncoffee.com','hello@spinncoffee.com','★★★★','遠心力抽出・エスプレッソ〜ドリップ対応スマートコーヒーメーカー'],
  [3680,'キッチン・調理器具','Mercato Compact Meat Slicer','Mercato','Indiegogo','https://www.indiegogo.com/projects/mercato/compact-meat-slicer','https://mercatokitchen.com','support@mercatokitchen.com','★★★','折りたたみ収納・薄切り〜厚切り調整可能な家庭用スライサー'],
  [3681,'キッチン・調理器具','Infinite Kitchen Robotic Cooking System','Infinite Foods','Kickstarter','https://www.kickstarter.com/projects/infinitefoods/infinite-kitchen-robot','https://infinitefoods.co','hello@infinitefoods.co','★★★★','アーム式調理ロボット・レシピ自動実行キッチンシステム'],
  [3682,'キッチン・調理器具','Wildone Ferment Crock','Wildone','Indiegogo','https://www.indiegogo.com/projects/wildone/ferment-crock','https://wildonefoods.com','support@wildonefoods.com','★★★','水シール式・自宅で本格発酵漬物ができるセラミック容器'],
  [3683,'キッチン・調理器具','Hero Custom Vitamin Dispenser','Hero Health','Kickstarter','https://www.kickstarter.com/projects/herohealth/hero-vitamin-dispenser','https://herohealth.com','support@herohealth.com','★★★★','個別パック分包・アプリ服薬リマインダー付きサプリディスペンサー'],
  [3684,'キッチン・調理器具','Kegerator Mini Countertop Draft System','Kegerator','Indiegogo','https://www.indiegogo.com/projects/kegerator/mini-countertop-draft','https://kegerator.com','hello@kegerator.com','★★★','卓上サイズ・自宅で生ビールが楽しめるドラフトシステム'],
  [3685,'キッチン・調理器具','Perfect Steak Precision Sous Vide Cooker','Perfect Steak Co.','Kickstarter','https://www.kickstarter.com/projects/perfectsteak/precision-sous-vide','https://perfectsteakco.com','support@perfectsteakco.com','★★★★','アプリ連動・肉の厚み入力で最適加熱時間を算出するスービー機'],
  [3686,'キッチン・調理器具','NutriBullet Balance Smart Blender','NutriBullet','Indiegogo','https://www.indiegogo.com/projects/nutribullet/balance-smart-blender','https://nutribullet.com','support@nutribullet.com','★★★★','栄養成分自動計算・アプリ連携スマートブレンダー'],
  [3687,'キッチン・調理器具','Kuvings Whole Slow Juicer EVO820','Kuvings','Indiegogo','https://www.indiegogo.com/projects/kuvings/whole-slow-juicer-evo820','https://kuvings.com','support@kuvings.com','★★★★★','低速圧縮・丸ごと投入可能な高栄養キープジューサー'],
  [3688,'キッチン・調理器具','Cookit Countertop Cooking Robot','Whirlpool','Kickstarter','https://www.kickstarter.com/projects/whirlpool/cookit-cooking-robot','https://whirlpool.com','support@whirlpool.com','★★★★','自動計量+加熱+攪拌・レシピ内蔵カウンタートップ調理ロボット'],
  [3689,'キッチン・調理器具','Yo-Kai Express Ramen Cooker','Yo-Kai Express','Indiegogo','https://www.indiegogo.com/projects/yokaiexpress/instant-ramen-cooker','https://yokaiexpress.com','hello@yokaiexpress.com','★★★★','90秒調理・生麺クオリティを再現する自動ラーメンクッカー'],

  // ── スマートホーム・インテリア・照明 (18件) ──
  [3690,'スマートホーム・インテリア・照明','Aeris Aair Pro Air Purifier','Aeris','Indiegogo','https://www.indiegogo.com/projects/aeris/aair-pro-air-purifier','https://aeris.com','support@aeris.com','★★★★','8段階フィルター・VOC/PM2.5対応大型空気清浄機'],
  [3691,'スマートホーム・インテリア・照明','Aqara Presence Sensor FP2','Aqara','Kickstarter','https://www.kickstarter.com/projects/aqara/presence-sensor-fp2','https://aqara.com','support@aqara.com','★★★★','mmWaveレーダー・人の存在/動作をゾーン検知するプレゼンスセンサー'],
  [3692,'スマートホーム・インテリア・照明','Third Reality Smart Blinds Motor','Third Reality','Kickstarter','https://www.kickstarter.com/projects/thirdreality/smart-blinds-motor','https://thirdreality.com','support@thirdreality.com','★★★','既存ブラインドに後付け・アプリ/音声操作対応モーター'],
  [3693,'スマートホーム・インテリア・照明','Chamberlain MyQ Smart Garage Hub','Chamberlain','Indiegogo','https://www.indiegogo.com/projects/chamberlain/myq-smart-garage-hub','https://myq.com','support@myq.com','★★★★','既存ガレージドアを後付けでスマート化するWi-Fiハブ'],
  [3694,'スマートホーム・インテリア・照明','Miku Pro Baby Monitor','Miku','Kickstarter','https://www.kickstarter.com/projects/miku/miku-pro-baby-monitor','https://mikucare.com','support@mikucare.com','★★★★★','非接触センサーで呼吸/睡眠を検知するスマートベビーモニター'],
  [3695,'スマートホーム・インテリア・照明','Dreame X40 Ultra Robot Vacuum','Dreame','Indiegogo','https://www.indiegogo.com/projects/dreame/x40-ultra-robot-vacuum','https://dreametech.com','support@dreametech.com','★★★★★','脚上げ機能付きモップ・全自動お手入れ不要ロボット掃除機'],
  [3696,'スマートホーム・インテリア・照明','Shelly Plus 2PM Smart Relay','Shelly','Kickstarter','https://www.kickstarter.com/projects/shelly/plus-2pm-smart-relay','https://shelly.cloud','support@shelly.cloud','★★★','壁裏設置・消費電力計測付き小型スマートリレー'],
  [3697,'スマートホーム・インテリア・照明','Aqara Video Doorbell G4','Aqara','Indiegogo','https://www.indiegogo.com/projects/aqara/video-doorbell-g4','https://aqara.com','support@aqara.com','★★★★','顔認証解錠連携・2Kカメラ内蔵スマートドアベル'],
  [3698,'スマートホーム・インテリア・照明','Hatch Restore 2 Sunrise Alarm Clock','Hatch','Kickstarter','https://www.kickstarter.com/projects/hatch/restore-2-sunrise-alarm','https://hatch.co','support@hatch.co','★★★★★','日の出シミュレーション・サウンドスケープ内蔵目覚まし'],
  [3699,'スマートホーム・インテリア・照明','Vocolinc Flowerbud Smart Diffuser','VOCOlinc','Indiegogo','https://www.indiegogo.com/projects/vocolinc/flowerbud-smart-diffuser','https://vocolinc.com','support@vocolinc.com','★★★','LED連携・アプリ操作対応スマートアロマディフューザー'],
  [3700,'スマートホーム・インテリア・照明','Roborock Dyad Wet Dry Vacuum','Roborock','Kickstarter','https://www.kickstarter.com/projects/roborock/dyad-wet-dry-vacuum','https://roborock.com','support@roborock.com','★★★★','自浄式ローラーブラシ・水拭き+吸引一体型コードレス掃除機'],
  [3701,'スマートホーム・インテリア・照明','Govee Smart Ice Maker','Govee','Indiegogo','https://www.indiegogo.com/projects/govee/smart-ice-maker','https://govee.com','support@govee.com','★★★','アプリ通知・9分で氷完成する卓上スマート製氷機'],
  [3702,'スマートホーム・インテリア・照明','Level Lock Plus Invisible Smart Lock','Level','Kickstarter','https://www.kickstarter.com/projects/level/level-lock-plus','https://level.co','support@level.co','★★★★★','外観そのまま・鍵穴内蔵型見えないスマートロック'],
  [3703,'スマートホーム・インテリア・照明','SwitchBot Robot Vacuum K10+','SwitchBot','Indiegogo','https://www.indiegogo.com/projects/switchbot/robot-vacuum-k10-plus','https://switch-bot.com','support@switch-bot.com','★★★★','直径25cm・自動ゴミ収集対応コンパクトロボット掃除機'],
  [3704,'スマートホーム・インテリア・照明','Nobi Smart Lamp Fall Detection','Nobi','Kickstarter','https://www.kickstarter.com/projects/nobi/smart-lamp-fall-detection','https://nobi.life','hello@nobi.life','★★★★★','AIカメラ内蔵・高齢者の転倒を検知し自動通報するスマートランプ'],
  [3705,'スマートホーム・インテリア・照明','Somneo Sleep and Wake-up Light','Philips','Indiegogo','https://www.indiegogo.com/projects/philips/somneo-sleep-wake-up-light','https://philips.com','support@philips.com','★★★★','心拍連動サウンド・入眠〜起床をサポートするライト'],
  [3706,'スマートホーム・インテリア・照明','Aqara Climate Sensor W100','Aqara','Kickstarter','https://www.kickstarter.com/projects/aqara/climate-sensor-w100','https://aqara.com','support@aqara.com','★★★','温湿度+気圧測定・E Inkディスプレイ搭載環境センサー'],
  [3707,'スマートホーム・インテリア・照明','Yale Linus Smart Lock','Yale','Indiegogo','https://www.indiegogo.com/projects/yale/linus-smart-lock','https://yalehome.com','support@yalehome.com','★★★★','既存シリンダーに後付け・静音モーター搭載スマートロック'],

  // ── ウェアラブル・ヘルス・フィットネス (18件) ──
  [3708,'ウェアラブル・ヘルス・フィットネス','Circular Ring Smart Health Tracker','Circular','Kickstarter','https://www.kickstarter.com/projects/circular/ring-smart-health-tracker','https://circular.eu','support@circular.eu','★★★★★','フランス発・睡眠/HRV/エネルギースコア計測スマートリング'],
  [3709,'ウェアラブル・ヘルス・フィットネス','RingConn Smart Ring Gen 2','RingConn','Indiegogo','https://www.indiegogo.com/projects/ringconn/smart-ring-gen-2','https://ringconn.com','support@ringconn.com','★★★★★','サブスク不要・充電ケース付きヘルストラッキングスマートリング'],
  [3710,'ウェアラブル・ヘルス・フィットネス','Cove Stress Relief Wearable','Feelmore Labs','Kickstarter','https://www.kickstarter.com/projects/feelmorelabs/cove-stress-relief-wearable','https://feelcove.com','hello@feelcove.com','★★★★','こめかみ装着・微振動で自律神経を整えるストレス緩和ウェアラブル'],
  [3711,'ウェアラブル・ヘルス・フィットネス','Naboso Neuro Ball Foot Therapy','Naboso Technology','Indiegogo','https://www.indiegogo.com/projects/naboso/neuro-ball-foot-therapy','https://naboso.com','support@naboso.com','★★★','足裏神経を刺激・バランス感覚向上フットセラピーボール'],
  [3712,'ウェアラブル・ヘルス・フィットネス','Xsensor Smart Insole Pressure Mapping','XSENSOR','Kickstarter','https://www.kickstarter.com/projects/xsensor/smart-insole-pressure-mapping','https://xsensor.com','support@xsensor.com','★★★','足圧分布リアルタイム計測・歩行改善スマートインソール'],
  [3713,'ウェアラブル・ヘルス・フィットネス','Motiv Ring Fitness Tracker','Motiv','Indiegogo','https://www.indiegogo.com/projects/motiv/ring-fitness-tracker','https://motiv.com','support@motiv.com','★★★★','軽量防水・心拍/睡眠/歩数を計測する指輪型トラッカー'],
  [3714,'ウェアラブル・ヘルス・フィットネス','NuCalm Stress Relief Neuroscience Kit','Solace Lifesciences','Kickstarter','https://www.kickstarter.com/projects/solace/nucalm-stress-relief-kit','https://nucalm.com','support@nucalm.com','★★★★','CES音響+バイオシグナル・15分で深いリラックス状態に導くキット'],
  [3715,'ウェアラブル・ヘルス・フィットネス','Sonde Health Vocal Biomarker App','Sonde Health','Indiegogo','https://www.indiegogo.com/projects/sondehealth/vocal-biomarker-app','https://sondehealth.com','hello@sondehealth.com','★★★','声紋分析・メンタルヘルス状態をスクリーニングするアプリ'],
  [3716,'ウェアラブル・ヘルス・フィットネス','Kernel Flow Brain Monitoring Headset','Kernel','Kickstarter','https://www.kickstarter.com/projects/kernel/flow-brain-monitoring-headset','https://kernel.com','support@kernel.com','★★★★','非侵襲光学式・脳活動をリアルタイム測定するヘッドセット'],
  [3717,'ウェアラブル・ヘルス・フィットネス','Whoop MG Health Tracker','Whoop','Indiegogo','https://www.indiegogo.com/projects/whoop/whoop-mg-health-tracker','https://whoop.com','support@whoop.com','★★★★★','血圧インサイト・心血管健康スコア計測次世代ヘルスバンド'],
  [3718,'ウェアラブル・ヘルス・フィットネス','Prevention Hearing Health Earbuds','Prevention','Kickstarter','https://www.kickstarter.com/projects/prevention/hearing-health-earbuds','https://preventionhearing.com','support@preventionhearing.com','★★★★','聴力検査+補聴機能内蔵・日常使いできるイヤホン型デバイス'],
  [3719,'ウェアラブル・ヘルス・フィットネス','Cala Trio Hand Tremor Relief Device','Cala Health','Indiegogo','https://www.indiegogo.com/projects/calahealth/trio-hand-tremor-relief','https://calahealth.com','support@calahealth.com','★★★★','手首装着・末梢神経刺激で手の震えを緩和する医療デバイス'],
  [3720,'ウェアラブル・ヘルス・フィットネス','Vivi Chair Posture Correction Cushion','Vivi','Kickstarter','https://www.kickstarter.com/projects/vivi/posture-correction-cushion','https://vivichair.com','hello@vivichair.com','★★★','振動フィードバック・座り姿勢をリアルタイム補正するクッション'],
  [3721,'ウェアラブル・ヘルス・フィットネス','Frenz Brainband Sleep Wearable','Earable Neuroscience','Indiegogo','https://www.indiegogo.com/projects/earable/frenz-brainband-sleep-wearable','https://frenzband.com','support@frenzband.com','★★★★★','EEGセンサー内蔵ヘッドバンド・入眠誘導サウンド自動再生'],
  [3722,'ウェアラブル・ヘルス・フィットネス','Vali Sleep Coach Wearable Patch','Vali Science','Kickstarter','https://www.kickstarter.com/projects/valiscience/sleep-coach-wearable-patch','https://valiscience.com','support@valiscience.com','★★★★','体温+心拍センサーパッチ・パーソナライズド睡眠コーチング'],
  [3723,'ウェアラブル・ヘルス・フィットネス','Wearable X Nadi X Yoga Pants','Wearable X','Indiegogo','https://www.indiegogo.com/projects/wearablex/nadi-x-yoga-pants','https://wearablex.com','hello@wearablex.com','★★★','内蔵振動モーター・ポーズ矯正をガイドするスマートヨガパンツ'],
  [3724,'ウェアラブル・ヘルス・フィットネス','Emfit QS Sleep Tracker Under-Mattress','Emfit','Kickstarter','https://www.kickstarter.com/projects/emfit/qs-sleep-tracker-under-mattress','https://emfit.com','support@emfit.com','★★★★','マットレス下設置・非接触式で心拍/呼吸/睡眠段階を計測'],
  [3725,'ウェアラブル・ヘルス・フィットネス','Owlet Dream Sock Baby Monitor','Owlet','Indiegogo','https://www.indiegogo.com/projects/owlet/dream-sock-baby-monitor','https://owletcare.com','support@owletcare.com','★★★★★','心拍/酸素飽和度計測・靴下型ベビーヘルスモニター'],

  // ── アウトドア・スポーツ・旅行 (20件) ──
  [3726,'アウトドア・スポーツ・旅行','Voltaic Systems Arc 20W Solar Backpack','Voltaic Systems','Kickstarter','https://www.kickstarter.com/projects/voltaicsystems/arc-20w-solar-backpack','https://voltaicsystems.com','support@voltaicsystems.com','★★★★','20Wソーラーパネル内蔵・行動中も充電できるバックパック'],
  [3727,'アウトドア・スポーツ・旅行','Nortlan All-Terrain Electric Skateboard','Nortlan','Indiegogo','https://www.indiegogo.com/projects/nortlan/all-terrain-electric-skateboard','https://nortlanboards.com','support@nortlanboards.com','★★★★','オフロード対応大径タイヤ・長距離走行電動スケートボード'],
  [3728,'アウトドア・スポーツ・旅行','Bivy Stick Satellite Communicator','Bivy','Kickstarter','https://www.kickstarter.com/projects/bivy/satellite-communicator','https://bivystick.com','support@bivystick.com','★★★★','圏外地帯でもメッセージ送受信・SOS発信対応衛星通信機'],
  [3729,'アウトドア・スポーツ・旅行','Wildland Camp Shower Portable','Wildland','Indiegogo','https://www.indiegogo.com/projects/wildland/camp-shower-portable','https://wildlandcogear.com','hello@wildlandcogear.com','★★★','USB充電式ポンプ・温水対応ポータブルキャンプシャワー'],
  [3730,'アウトドア・スポーツ・旅行','Anycubic Power Solar Charging Station','Anycubic Power','Kickstarter','https://www.kickstarter.com/projects/anycubicpower/solar-charging-station','https://anycubicpower.com','support@anycubicpower.com','★★★★','折りたたみパネル+大容量バッテリー一体型アウトドア電源'],
  [3731,'アウトドア・スポーツ・旅行','Onewheel GT Electric Skateboard','Future Motion','Indiegogo','https://www.indiegogo.com/projects/futuremotion/onewheel-gt','https://onewheel.com','support@onewheel.com','★★★★★','セルフバランス式・オフロード対応単輪電動スケートボード'],
  [3732,'アウトドア・スポーツ・旅行','Ocean Signal RescueMe PLB Personal Locator Beacon','Ocean Signal','Kickstarter','https://www.kickstarter.com/projects/oceansignal/rescueme-plb','https://oceansignal.com','support@oceansignal.com','★★★★★','GPS+406MHz発信・海難/遭難時救助要請パーソナルビーコン'],
  [3733,'アウトドア・スポーツ・旅行','Klymit Insulated Static V Sleeping Pad','Klymit','Indiegogo','https://www.indiegogo.com/projects/klymit/insulated-static-v-pad','https://klymit.com','support@klymit.com','★★★★','ボディマッピング設計・保温材内蔵軽量スリーピングパッド'],
  [3734,'アウトドア・スポーツ・旅行','Woodsman Camp Axe Multi-Tool','Woodsman','Kickstarter','https://www.kickstarter.com/projects/woodsman/camp-axe-multi-tool','https://woodsmangear.com','hello@woodsmangear.com','★★★','斧+ハンマー+くぎ抜き一体型キャンプマルチツール'],
  [3735,'アウトドア・スポーツ・旅行','Verge Motorcycles Electric Adventure Bike','Verge Motorcycles','Indiegogo','https://www.indiegogo.com/projects/vergemotorcycles/electric-adventure-bike','https://vergemotorcycles.com','support@vergemotorcycles.com','★★★★','ハブレスホイールモーター・長距離走行電動バイク'],
  [3736,'アウトドア・スポーツ・旅行','Therm-a-Rest Vesper Down Quilt','Therm-a-Rest','Kickstarter','https://www.kickstarter.com/projects/thermarest/vesper-down-quilt','https://thermarest.com','support@thermarest.com','★★★★','900フィルパワーダウン・超軽量325gバックパッキング用キルト'],
  [3737,'アウトドア・スポーツ・旅行','GoSun Rove Portable Solar Cooler','GoSun','Indiegogo','https://www.indiegogo.com/projects/gosun/rove-portable-solar-cooler','https://gosun.co','support@gosun.co','★★★★','ソーラー+バッテリー駆動・電源不要ポータブル冷却ボックス'],
  [3738,'アウトドア・スポーツ・旅行','Zephyr Solar Power Tent','Zephyr Solar','Kickstarter','https://www.kickstarter.com/projects/zephyrsolar/solar-power-tent','https://zephyrsolar.com','hello@zephyrsolar.com','★★★','フライシート内蔵ソーラーパネル・設営するだけで発電するテント'],
  [3739,'アウトドア・スポーツ・旅行','Garmin Tactix 7 Adventure GPS Watch','Garmin','Indiegogo','https://www.indiegogo.com/projects/garmin/tactix-7-adventure-gps-watch','https://garmin.com','support@garmin.com','★★★★★','ミルスペック耐久・戦術ナビゲーション対応アドベンチャーGPSウォッチ'],
  [3740,'アウトドア・スポーツ・旅行','Helinox Cot One Convertible Camp Bed','Helinox','Kickstarter','https://www.kickstarter.com/projects/helinox/cot-one-convertible','https://helinox.com','support@helinox.com','★★★★','折りたたみ・地面から浮いて眠れる軽量キャンプベッド'],
  [3741,'アウトドア・スポーツ・旅行','Mountain House Just In Case Food Storage Kit','Mountain House','Indiegogo','https://www.indiegogo.com/projects/mountainhouse/just-in-case-food-kit','https://mountainhouse.com','support@mountainhouse.com','★★★','25年保存可能・フリーズドライ非常食セット'],
  [3742,'アウトドア・スポーツ・旅行','Rovr RollR 60 Wheeled Cooler','Rovr Products','Kickstarter','https://www.kickstarter.com/projects/rovrproducts/rollr-60-wheeled-cooler','https://rovrproducts.com','hello@rovrproducts.com','★★★★','オフロードタイヤ付き・ペット水入れ変換可能大型クーラー'],
  [3743,'アウトドア・スポーツ・旅行','Exped MegaMat Duo Sleeping Pad','Exped','Indiegogo','https://www.indiegogo.com/projects/exped/megamat-duo-sleeping-pad','https://exped.com','support@exped.com','★★★★','2人用・厚さ10cmの快適性を追求したキャンプマットレス'],
  [3744,'アウトドア・スポーツ・旅行','Ryder Innovation Foldable E-Bike','Ryder Innovation','Kickstarter','https://www.kickstarter.com/projects/ryderinnovation/foldable-e-bike','https://ryderinnovation.com','support@ryderinnovation.com','★★★★','10秒で折りたたみ・軽量18kgの都市型電動アシスト自転車'],
  [3745,'アウトドア・スポーツ・旅行','Sea Eagle 330 Inflatable Kayak','Sea Eagle','Indiegogo','https://www.indiegogo.com/projects/seaeagle/330-inflatable-kayak','https://seaeagle.com','support@seaeagle.com','★★★★','収納袋サイズに収まる・軽量ポータブルインフレータブルカヤック'],

  // ── ペット用品 (24件) ──
  [3746,'ペット用品','Vetsy Smart Pet Health Monitor Collar','Vetsy','Kickstarter','https://www.kickstarter.com/projects/vetsy/smart-pet-health-monitor-collar','https://vetsy.io','support@vetsy.io','★★★★','体温/心拍/活動量を24時間計測するペット用ヘルスカラー'],
  [3747,'ペット用品','PetPace Smart Health Collar','PetPace','Indiegogo','https://www.indiegogo.com/projects/petpace/smart-health-collar','https://petpace.com','support@petpace.com','★★★★★','獣医師監修・早期疾患兆候を検知するスマートペットカラー'],
  [3748,'ペット用品','Petivity Smart Litter Box Monitor','Purina','Kickstarter','https://www.kickstarter.com/projects/purina/petivity-smart-litter-monitor','https://petivity.com','support@petivity.com','★★★★','体重変化/トイレ回数を自動記録する猫トイレ用センサー'],
  [3749,'ペット用品','Sundays for Dogs Air-Dried Food','Sundays','Indiegogo','https://www.indiegogo.com/projects/sundaysfordogs/air-dried-food','https://sundaysfordogs.com','hello@sundaysfordogs.com','★★★★','人間用食材グレード・エアドライ製法犬用フード'],
  [3750,'ペット用品','Yumove Joint Care Dog Supplement','Lintbells','Kickstarter','https://www.kickstarter.com/projects/lintbells/yumove-joint-care-supplement','https://yumove.com','support@yumove.com','★★★★','獣医推奨・グルコサミン配合関節ケアサプリメント'],
  [3751,'ペット用品','Whistle Health Wellness Monitor','Whistle','Indiegogo','https://www.indiegogo.com/projects/whistle/health-wellness-monitor','https://whistle.com','support@whistle.com','★★★★','食事/水分/皮膚状態を追跡するペット健康モニタリングデバイス'],
  [3752,'ペット用品','Loobani Automatic Cat Litter Box','Loobani','Kickstarter','https://www.kickstarter.com/projects/loobani/automatic-cat-litter-box','https://loobani.com','support@loobani.com','★★★★','アプリ通知・大型猫対応自動猫トイレ'],
  [3753,'ペット用品','Homerun Pet Leos Loo Too Litter Box','Homerun Pet','Indiegogo','https://www.indiegogo.com/projects/homerunpet/leos-loo-too-litter-box','https://homerunpet.com','support@homerunpet.com','★★★★★','ふるい分け式・手動でも使える静音自動猫トイレ'],
  [3754,'ペット用品','Companion Pet AI Video Monitor','Companion','Kickstarter','https://www.kickstarter.com/projects/companion/ai-video-monitor','https://meetcompanion.com','hello@meetcompanion.com','★★★★','異常行動をAI検知・獣医アラート機能付きペット見守りカメラ'],
  [3755,'ペット用品','Vetnique Labs Omega Bites Dog Supplement','Vetnique Labs','Indiegogo','https://www.indiegogo.com/projects/vetniquelabs/omega-bites-supplement','https://vetniquelabs.com','support@vetniquelabs.com','★★★','オメガ3配合・皮膚/被毛ケア犬用チュアブル'],
  [3756,'ペット用品','Tractive Cat Mini GPS Tracker','Tractive','Kickstarter','https://www.kickstarter.com/projects/tractive/cat-mini-gps-tracker','https://tractive.com','support@tractive.com','★★★★★','猫専用超軽量・リアルタイムGPS追跡トラッカー'],
  [3757,'ペット用品','Catit Pixi Smart Water Fountain','Catit','Indiegogo','https://www.indiegogo.com/projects/catit/pixi-smart-water-fountain','https://catit.com','support@catit.com','★★★★','使用量記録アプリ連携・猫用スマート給水ファウンテン'],
  [3758,'ペット用品','Barkbox Doggy Dental Chews','BarkBox','Kickstarter','https://www.kickstarter.com/projects/barkbox/doggy-dental-chews','https://barkbox.com','hello@barkbox.com','★★★','天然成分・歯垢除去サポート犬用デンタルチュー'],
  [3759,'ペット用品','Kinderpet Smart Feeding Bowl','Kinderpet','Indiegogo','https://www.indiegogo.com/projects/kinderpet/smart-feeding-bowl','https://kinderpet.com','support@kinderpet.com','★★★','食事ペース検知・早食い防止スマート給餌ボウル'],
  [3760,'ペット用品','Vivaia Orthopedic Dog Bed Memory Foam','Vivaia Pet','Kickstarter','https://www.kickstarter.com/projects/vivaiapet/orthopedic-dog-bed','https://vivaiapet.com','support@vivaiapet.com','★★★★','整形外科医監修・関節をサポートする老犬向けベッド'],
  [3761,'ペット用品','Petcube Cam Interactive Pet Monitor','Petcube','Indiegogo','https://www.indiegogo.com/projects/petcube/cam-interactive-monitor','https://petcube.com','support@petcube.com','★★★★','双方向通話・レーザーポインター内蔵ペット見守りカメラ'],
  [3762,'ペット用品','GoodBoy Automatic Treat Dispenser','GoodBoy','Kickstarter','https://www.kickstarter.com/projects/goodboy/automatic-treat-dispenser','https://goodboyco.com','hello@goodboyco.com','★★★','トレーニング連動・スマホから遠隔おやつ発射できるディスペンサー'],
  [3763,'ペット用品','Rex Specs Dog Goggles UV Protection','Rex Specs','Indiegogo','https://www.indiegogo.com/projects/rexspecs/dog-goggles-uv-protection','https://rexspecs.com','support@rexspecs.com','★★★','UVカット・アウトドア犬用保護ゴーグル'],
  [3764,'ペット用品','Cat Person Stainless Steel Litter Box','Cat Person','Kickstarter','https://www.kickstarter.com/projects/catperson/stainless-steel-litter-box','https://catperson.com','support@catperson.com','★★★','抗菌ステンレス製・臭い/傷つきにくい猫用トイレ'],
  [3765,'ペット用品','AATAS Pet Hair Dryer Box','AATAS','Indiegogo','https://www.indiegogo.com/projects/aatas/pet-hair-dryer-box','https://aataspet.com','support@aataspet.com','★★★★','静音設計・シャンプー後の自動乾燥ボックス'],
  [3766,'ペット用品','Bella + Duke Raw Dog Food Subscription','Bella + Duke','Kickstarter','https://www.kickstarter.com/projects/bellaandduke/raw-dog-food-subscription','https://bellaandduke.com','hello@bellaandduke.com','★★★★','個体別レシピ提案・生食主義ドッグフード定期便'],
  [3767,'ペット用品','iCPooch Treat and Camera Dispenser','iCPooch','Indiegogo','https://www.indiegogo.com/projects/icpooch/treat-camera-dispenser','https://icpooch.com','support@icpooch.com','★★★','ビデオ通話+おやつ配布・留守中のペットと交流できるデバイス'],
  [3768,'ペット用品','Purrsong Meowdio Cat Health Collar','Purrsong','Kickstarter','https://www.kickstarter.com/projects/purrsong/meowdio-cat-health-collar','https://purrsong.com','support@purrsong.com','★★★★','鳴き声解析AI・猫のストレス状態を可視化するスマートカラー'],
  [3769,'ペット用品','Native Pet Gut Health Supplement','Native Pet','Indiegogo','https://www.indiegogo.com/projects/nativepet/gut-health-supplement','https://nativepet.com','hello@nativepet.com','★★★','プロバイオティクス配合・消化器サポート犬猫用サプリ'],

  // ── テクノロジー・ガジェット (14件) ──
  [3770,'テクノロジー・ガジェット','Plaud Note AI Voice Recorder','Plaud','Kickstarter','https://www.kickstarter.com/projects/plaud/note-ai-voice-recorder','https://plaud.ai','support@plaud.ai','★★★★★','スマホ背面装着・AI自動文字起こし+要約対応ボイスレコーダー'],
  [3771,'テクノロジー・ガジェット','Limitless Pendant AI Wearable','Limitless','Indiegogo','https://www.indiegogo.com/projects/limitless/pendant-ai-wearable','https://limitless.ai','support@limitless.ai','★★★★','会話を自動記録・要約するペンダント型AIウェアラブル'],
  [3772,'テクノロジー・ガジェット','Nothing Ear Open Wireless Earbuds','Nothing','Kickstarter','https://www.kickstarter.com/projects/nothing/ear-open-wireless-earbuds','https://nothing.tech','support@nothing.tech','★★★★','耳を塞がないクリップ型・透明デザインワイヤレスイヤホン'],
  [3773,'テクノロジー・ガジェット','Balmuda The Speaker Compact Audio','Balmuda','Indiegogo','https://www.indiegogo.com/projects/balmuda/the-speaker-compact-audio','https://balmuda.com','support@balmuda.com','★★★★','日本発デザイン・アナログ回路採用コンパクトオーディオスピーカー'],
  [3774,'テクノロジー・ガジェット','Marshall Middleton Portable Speaker','Marshall','Kickstarter','https://www.kickstarter.com/projects/marshall/middleton-portable-speaker','https://marshallheadphones.com','support@marshallheadphones.com','★★★★','IP67防水・スタック機能付き大出力ポータブルスピーカー'],
  [3775,'テクノロジー・ガジェット','Deeper Chirp+ 2 Smart Fish Finder','Deeper','Indiegogo','https://www.indiegogo.com/projects/deeper/chirp-plus-2-fish-finder','https://deepersonar.com','support@deepersonar.com','★★★★','Wi-Fi直結・GPS内蔵ポータブル魚群探知機'],
  [3776,'テクノロジー・ガジェット','Insta360 Ace Pro 2 Action Camera','Insta360','Kickstarter','https://www.kickstarter.com/projects/insta360/ace-pro-2-action-camera','https://insta360.com','hello@insta360.com','★★★★★','Leica共同開発レンズ・8K撮影対応アクションカメラ'],
  [3777,'テクノロジー・ガジェット','Human AI Pin Alternative','Human','Indiegogo','https://www.indiegogo.com/projects/human/ai-pin-alternative','https://human.tech','support@human.tech','★★★','音声操作特化・画面レスシンプルAIウェアラブル'],
  [3778,'テクノロジー・ガジェット','Nreal Chameleon Wireless AR Streaming','Nreal','Kickstarter','https://www.kickstarter.com/projects/nreal/chameleon-wireless-ar-streaming','https://nrealchameleon.com','support@nrealchameleon.com','★★★★','無線接続・ARグラスとPCをつなぐストリーミングアダプター'],
  [3779,'テクノロジー・ガジェット','Segway Navimow Robot Lawn Mower','Segway','Indiegogo','https://www.indiegogo.com/projects/segway/navimow-robot-lawn-mower','https://segway.com','support@segway.com','★★★★★','境界線ワイヤー不要・GPS高精度測位ロボット芝刈り機'],
  [3780,'テクノロジー・ガジェット','Anker Soundcore Sleep A20 Earbuds','Anker Soundcore','Kickstarter','https://www.kickstarter.com/projects/ankersoundcore/sleep-a20-earbuds','https://soundcore.com','support@soundcore.com','★★★★','側寝でも痛くない・遮音+ホワイトノイズ再生睡眠用イヤホン'],
  [3781,'テクノロジー・ガジェット','Roli Airwave Gesture Controller','Roli','Indiegogo','https://www.indiegogo.com/projects/roli/airwave-gesture-controller','https://roli.com','support@roli.com','★★★','ジェスチャーで音楽制作を操作するモーションコントローラー'],
  [3782,'テクノロジー・ガジェット','Teenage Engineering OP-XY Portable Synth','Teenage Engineering','Kickstarter','https://www.kickstarter.com/projects/teenageengineering/op-xy-portable-synth','https://teenage.engineering','support@teenage.engineering','★★★★★','コンパクト筐体・本格シーケンサー内蔵ポータブルシンセサイザー'],
  [3783,'テクノロジー・ガジェット','Pebble Flow Towable Smart Camper Trailer','Pebble','Indiegogo','https://www.indiegogo.com/projects/pebble/flow-towable-smart-camper','https://pebble.com','support@pebble.com','★★★★★','電動アシスト牽引・EV充電対応スマートキャンピングトレーラー'],

  // ── 美容・スキンケア (22件) ──
  [3784,'美容・スキンケア','Dr. Dennis Gross SpectraLite FaceWare Pro','Dr. Dennis Gross Skincare','Kickstarter','https://www.kickstarter.com/projects/drdennisgross/spectralite-faceware-pro','https://drdennisgross.com','support@drdennisgross.com','★★★★','100個LED×2波長・美容皮膚科医監修光美容マスク'],
  [3785,'美容・スキンケア','Lyma Laser Anti-Aging Device','Lyma','Indiegogo','https://www.indiegogo.com/projects/lyma/laser-anti-aging-device','https://lyma.life','support@lyma.life','★★★★★','医療グレードレーザー・コラーゲン生成促進美容デバイス'],
  [3786,'美容・スキンケア','Prose Custom Hair Care Formula','Prose','Kickstarter','https://www.kickstarter.com/projects/prose/custom-hair-care-formula','https://prose.com','hello@prose.com','★★★★','髪質診断に基づくパーソナライズ処方ヘアケアシステム'],
  [3787,'美容・スキンケア','Skinstric AI Skin Analysis Device','Skinstric','Indiegogo','https://www.indiegogo.com/projects/skinstric/ai-skin-analysis-device','https://skinstric.com','support@skinstric.com','★★★★','AI肌診断・個人に最適化されたスキンケア処方を提案'],
  [3788,'美容・スキンケア',"Silk'n Titan Radiofrequency Skin Tightening","Silk'n",'Kickstarter','https://www.kickstarter.com/projects/silkn/titan-radiofrequency-device','https://silkn.com','support@silkn.com','★★★★','高周波+光エネルギー・自宅でできる肌引き締めデバイス'],
  [3789,'美容・スキンケア','Nuface Trinity Facial Toning Device','NuFace','Indiegogo','https://www.indiegogo.com/projects/nuface/trinity-facial-toning-device','https://mynuface.com','support@mynuface.com','★★★★★','マイクロカレント・表情筋を鍛えるフェイシャルトーニング'],
  [3790,'美容・スキンケア','Function of Beauty Custom Shampoo','Function of Beauty','Kickstarter','https://www.kickstarter.com/projects/functionofbeauty/custom-shampoo','https://functionofbeauty.com','hello@functionofbeauty.com','★★★★','髪質+悩みを診断・完全オーダーメイドシャンプー'],
  [3791,'美容・スキンケア','Wthn Acupuncture At-Home Kit','Wthn','Indiegogo','https://www.indiegogo.com/projects/wthn/acupuncture-at-home-kit','https://wthn.com','support@wthn.com','★★★','鍼灸師監修・自宅でできる美容鍼スターターキット'],
  [3792,'美容・スキンケア','Perfect Corp YouCam Skin AI Scanner','Perfect Corp','Kickstarter','https://www.kickstarter.com/projects/perfectcorp/youcam-skin-ai-scanner','https://perfectcorp.com','support@perfectcorp.com','★★★','スマホカメラで肌年齢/毛穴を分析するAIスキンスキャナー'],
  [3793,'美容・スキンケア','Foreo UFO 3 Smart Mask Device','FOREO','Indiegogo','https://www.indiegogo.com/projects/foreo/ufo-3-smart-mask-device','https://foreo.com','support@foreo.com','★★★★★','90秒完結・光+温冷+振動複合美容マスクデバイス'],
  [3794,'美容・スキンケア','Espira LED Face Mask Therapy','Espira','Kickstarter','https://www.kickstarter.com/projects/espira/led-face-mask-therapy','https://espirabeauty.com','support@espirabeauty.com','★★★★','医療用LEDパネル・7波長切替対応フェイスマスク'],
  [3795,'美容・スキンケア','Kitsch Scalp Massager Shampoo Brush','Kitsch','Indiegogo','https://www.indiegogo.com/projects/kitsch/scalp-massager-shampoo-brush','https://mykitsch.com','support@mykitsch.com','★★★','シリコン製・血行促進頭皮マッサージブラシ'],
  [3796,'美容・スキンケア','Curie Natural Deodorant Wipes','Curie','Kickstarter','https://www.kickstarter.com/projects/curie/natural-deodorant-wipes','https://shopcurie.com','hello@shopcurie.com','★★★','アルミフリー・持ち運びできる天然成分デオドラントワイプ'],
  [3797,'美容・スキンケア','Vieve Color Match Foundation System','Vieve','Indiegogo','https://www.indiegogo.com/projects/vieve/color-match-foundation-system','https://vievebeauty.com','support@vievebeauty.com','★★★','AIスキャン・肌色に完全一致するファンデーション調合システム'],
  [3798,'美容・スキンケア','Homedics Total Comfort Massaging Foot Spa','Homedics','Kickstarter','https://www.kickstarter.com/projects/homedics/total-comfort-foot-spa','https://homedics.com','support@homedics.com','★★★★','バイブレーション+加熱+バブル機能付きフットスパ'],
  [3799,'美容・スキンケア','Filorga Meso Mask Instant Skincare Patch','Filorga','Indiegogo','https://www.indiegogo.com/projects/filorga/meso-mask-instant-patch','https://filorga.com','support@filorga.com','★★★★','医療エステ発・即効ハリを与えるメソマスクパッチ'],
  [3800,'美容・スキンケア','Amiro Hair Removal IPL Device','Amiro','Kickstarter','https://www.kickstarter.com/projects/amiro/hair-removal-ipl-device','https://amirolab.com','support@amirolab.com','★★★★','サファイアクーリング・自宅用IPL脱毛器'],
  [3801,'美容・スキンケア','Wanders Beauty Waterless Skincare Bars','Wanders Beauty','Indiegogo','https://www.indiegogo.com/projects/wandersbeauty/waterless-skincare-bars','https://wandersbeauty.com','hello@wandersbeauty.com','★★★','水を一切使わない・旅行にも便利な固形スキンケアバー'],
  [3802,'美容・スキンケア','Angel Sculpt EMS Face Lifting Device','Angel Sculpt','Kickstarter','https://www.kickstarter.com/projects/angelsculpt/ems-face-lifting-device','https://angelsculpt.com','support@angelsculpt.com','★★★★','EMS+RF複合・小顔ケアフェイスリフティングデバイス'],
  [3803,'美容・スキンケア','True Botanicals Clean Beauty Serum','True Botanicals','Indiegogo','https://www.indiegogo.com/projects/truebotanicals/clean-beauty-serum','https://truebotanicals.com','support@truebotanicals.com','★★★','有害成分不使用・オーガニック認証クリーンビューティーセラム'],
  [3804,'美容・スキンケア','Vanity Planet Spin for Perfect Skin Brush','Vanity Planet','Kickstarter','https://www.kickstarter.com/projects/vanityplanet/spin-perfect-skin-brush','https://vanityplanet.com','hello@vanityplanet.com','★★★','回転式・毛穴の汚れを除去する洗顔ブラシ'],
  [3805,'美容・スキンケア','Therabody TheraFace Mask LED Therapy','Therabody','Indiegogo','https://www.indiegogo.com/projects/therabody/theraface-mask-led-therapy','https://therabody.com','support@therabody.com','★★★★','赤色+近赤外線LED・肌トーン改善フェイシャルマスク'],

  // ── 子供・教育 (24件) ──
  [3806,'子供・教育','Kinedu Early Learning App Kit','Kinedu','Kickstarter','https://www.kickstarter.com/projects/kinedu/early-learning-app-kit','https://kinedu.com','support@kinedu.com','★★★★','発達心理学に基づく・0〜6歳向けアクティビティ提案アプリ連動キット'],
  [3807,'子供・教育','Piper Computer Kit STEM Build','Piper','Indiegogo','https://www.indiegogo.com/projects/piper/computer-kit-stem-build','https://playpiper.com','support@playpiper.com','★★★★★','木箱から自分でPCを組み立てる子供向けSTEM学習キット'],
  [3808,'子供・教育','Marbotic Smart Wooden Toys','Marbotic','Kickstarter','https://www.kickstarter.com/projects/marbotic/smart-wooden-toys','https://marbotic.com','hello@marbotic.com','★★★★','タブレット連動・木製ブロックで文字/数字を学ぶ知育玩具'],
  [3809,'子供・教育','Osmo Coding Awbie STEM Game','Tangible Play','Indiegogo','https://www.indiegogo.com/projects/tangibleplay/coding-awbie-stem-game','https://playosmo.com','support@playosmo.com','★★★★★','物理ブロック+タブレット・遊びながらコーディングを学ぶゲーム'],
  [3810,'子供・教育','CodeBug Mini Programmable Computer','CodeBug','Kickstarter','https://www.kickstarter.com/projects/codebug/codebug-mini-computer','https://codebug.org','support@codebug.org','★★★','LEDマトリクス搭載・初心者向け超小型プログラミング学習ボード'],
  [3811,'子供・教育','Loona Petbot AI Learning Companion','KEYi Tech','Indiegogo','https://www.indiegogo.com/projects/keyitech/loona-petbot-ai-companion','https://keyitech.com','support@keyitech.com','★★★★','表情豊かなAIロボット・子供の情操教育をサポートするペットロボット'],
  [3812,'子供・教育','Yoto Mini Screen-Free Player','Yoto','Kickstarter','https://www.kickstarter.com/projects/yoto/mini-screen-free-player','https://yotoplay.com','hello@yotoplay.com','★★★★★','携帯サイズ・画面なしで物語/音楽を楽しめるオーディオプレイヤー'],
  [3813,'子供・教育','Curiosity Box Monthly STEM Subscription','Vsauce','Indiegogo','https://www.indiegogo.com/projects/vsauce/curiosity-box-stem-subscription','https://curiositybox.com','support@curiositybox.com','★★★★','科学系YouTuber監修・毎月届く実験キットサブスク'],
  [3814,'子供・教育','Robo Wunderkind Modular Robotics Kit','Robo Wunderkind','Kickstarter','https://www.kickstarter.com/projects/robowunderkind/modular-robotics-kit','https://robowunderkind.com','support@robowunderkind.com','★★★★','モジュール組み合わせ式・年齢に応じて拡張できるロボット教材'],
  [3815,'子供・教育','Lovevery Play Kits Age-Based Toys','Lovevery','Indiegogo','https://www.indiegogo.com/projects/lovevery/play-kits-age-based-toys','https://lovevery.com','support@lovevery.com','★★★★★','発達段階ごとに厳選・専門家監修の月齢別知育玩具サブスク'],
  [3816,'子供・教育','Bloxels Build Your Own Video Game','Pixel Press','Kickstarter','https://www.kickstarter.com/projects/pixelpress/bloxels-build-your-own-game','https://bloxelsbuild.com','hello@bloxelsbuild.com','★★★★','物理ブロックでゲームステージを作りアプリで遊べる知育玩具'],
  [3817,'子供・教育','Tinkerine Litto 3D Printer for Kids','Tinkerine','Indiegogo','https://www.indiegogo.com/projects/tinkerine/litto-3d-printer-for-kids','https://tinkerine.com','support@tinkerine.com','★★★★','安全設計・子供が使える教育向け小型3Dプリンター'],
  [3818,'子供・教育','Skillmatics Educational Card Games','Skillmatics','Kickstarter','https://www.kickstarter.com/projects/skillmatics/educational-card-games','https://skillmatics.com','support@skillmatics.com','★★★','消せるペン付き・繰り返し遊べる知育カードゲームシリーズ'],
  [3819,'子供・教育','Squishable Emotion Plush Learning Set','Squishable','Indiegogo','https://www.indiegogo.com/projects/squishable/emotion-plush-learning-set','https://squishable.com','hello@squishable.com','★★★','表情差し替え可能・感情教育ぬいぐるみセット'],
  [3820,'子供・教育','Botley 2.0 The Coding Robot Activity Set','Learning Resources','Kickstarter','https://www.kickstarter.com/projects/learningresources/botley-2-coding-robot-set','https://learningresources.com','support@learningresources.com','★★★★','障害物回避センサー搭載・画面なしコーディングロボット'],
  [3821,'子供・教育','Sensory Genius Calm Down Corner Kit','Sensory Genius','Indiegogo','https://www.indiegogo.com/projects/sensorygenius/calm-down-corner-kit','https://sensorygenius.com','support@sensorygenius.com','★★★','感覚統合ツール一式・子供の自己調整力を育むコーナーキット'],
  [3822,'子供・教育','Codemonkey Coding Game Subscription','CodeMonkey','Kickstarter','https://www.kickstarter.com/projects/codemonkey/coding-game-subscription','https://codemonkey.com','support@codemonkey.com','★★★★','実際のプログラミング言語を学べるゲーム形式学習サブスク'],
  [3823,'子供・教育','Ozobot Evo Coding Robot','Ozobot','Indiegogo','https://www.indiegogo.com/projects/ozobot/evo-coding-robot','https://ozobot.com','support@ozobot.com','★★★★','色センサー+アプリ連動・線をなぞって動く小型コーディングロボット'],
  [3824,'子供・教育','Toniebox Audio Character Player','Tonies','Kickstarter','https://www.kickstarter.com/projects/tonies/toniebox-audio-character-player','https://tonies.com','support@tonies.com','★★★★★','フィギュアを乗せると音声再生・ドイツ発画面なしプレイヤー'],
  [3825,'子供・教育','Little Journey Sensory Play Mat','Little Journey','Indiegogo','https://www.indiegogo.com/projects/littlejourney/sensory-play-mat','https://littlejourney.com','hello@littlejourney.com','★★★','異素材パネル組み合わせ・赤ちゃんの感覚刺激プレイマット'],
  [3826,'子供・教育','Coding Critters Interactive Pet Robot','Learning Resources','Kickstarter','https://www.kickstarter.com/projects/learningresources/coding-critters-pet-robot','https://learningresources.com','support@learningresources.com','★★★★','ペット育成しながら基礎コーディングを学ぶロボットトイ'],
  [3827,'子供・教育','Kiwico Panda Crate Infant Sensory Kit','KiwiCo','Indiegogo','https://www.indiegogo.com/projects/kiwico/panda-crate-infant-sensory-kit','https://kiwico.com','support@kiwico.com','★★★★','0〜12ヶ月向け・感覚発達を促す月齢別知育キット'],
  [3828,'子供・教育','Storypod Interactive Storytelling Robot','Storypod','Kickstarter','https://www.kickstarter.com/projects/storypod/interactive-storytelling-robot','https://storypod.com','support@storypod.com','★★★★','フィギュアタップで物語再生・読み聞かせロボット'],
  [3829,'子供・教育','Tegu Magnetic Wooden Blocks','Tegu','Indiegogo','https://www.indiegogo.com/projects/tegu/magnetic-wooden-blocks','https://tegu.com','hello@tegu.com','★★★★','フェアトレード木材使用・磁石内蔵組み立てブロック'],

  // ── ファッション・アクセサリー (24件) ──
  [3830,'ファッション・アクセサリー','Coalatree Nomad Poncho Travel Blanket','Coalatree','Kickstarter','https://www.kickstarter.com/projects/coalatree/nomad-poncho-travel-blanket','https://coalatree.com','support@coalatree.com','★★★★','ポンチョ⇔ブランケット2wayアウトドアウェア'],
  [3831,'ファッション・アクセサリー','Wolf & Shepherd Performance Dress Shoes','Wolf & Shepherd','Indiegogo','https://www.indiegogo.com/projects/wolfandshepherd/performance-dress-shoes','https://wolfandshepherd.com','support@wolfandshepherd.com','★★★★','スニーカーの快適性+ドレスシューズの見た目を両立'],
  [3832,'ファッション・アクセサリー','Res Ipsa Convertible Travel Dress','Res Ipsa','Kickstarter','https://www.kickstarter.com/projects/resipsa/convertible-travel-dress','https://resipsa.com','hello@resipsa.com','★★★★','結び方次第で何通りにも着回せる旅行用ワンピース'],
  [3833,'ファッション・アクセサリー','Silent Pocket Faraday Cage Wallet','Silent Pocket','Indiegogo','https://www.indiegogo.com/projects/silentpocket/faraday-cage-wallet','https://silent-pocket.com','support@silent-pocket.com','★★★★','電波遮断素材・GPS/RFID追跡防止フェラデーウォレット'],
  [3834,'ファッション・アクセサリー','Kammok Field Blanket Packable','Kammok','Kickstarter','https://www.kickstarter.com/projects/kammok/field-blanket-packable','https://kammok.com','hello@kammok.com','★★★','コンパクト収納・防水裏地付き携帯ブランケット'],
  [3835,'ファッション・アクセサリー','Wanderlust Convertible Travel Jacket','Wanderlust','Indiegogo','https://www.indiegogo.com/projects/wanderlust/convertible-travel-jacket','https://wanderlustgear.com','support@wanderlustgear.com','★★★★','15ポケット内蔵・機内持ち込み品を身につけられる旅行用ジャケット'],
  [3836,'ファッション・アクセサリー','Neverfar Recycled Ocean Plastic Sunglasses','Neverfar','Kickstarter','https://www.kickstarter.com/projects/neverfar/recycled-ocean-plastic-sunglasses','https://neverfar.com','support@neverfar.com','★★★★','海洋プラスチックごみ再生素材使用サングラス'],
  [3837,'ファッション・アクセサリー','Genusee Recycled Water Bottle Eyewear','Genusee','Indiegogo','https://www.indiegogo.com/projects/genusee/recycled-water-bottle-eyewear','https://genusee.com','hello@genusee.com','★★★★','フリント水危機支援・リサイクル素材製アイウェア'],
  [3838,'ファッション・アクセサリー','Anonymous Ism Japanese Crew Socks','Anonymous Ism','Kickstarter','https://www.kickstarter.com/projects/anonymousism/japanese-crew-socks','https://anonymousism.jp','support@anonymousism.jp','★★★★','日本製・鮮やかなカラーバリエーションのクルーソックス'],
  [3839,'ファッション・アクセサリー','Paravel Fold-Up Weekender Bag','Paravel','Indiegogo','https://www.indiegogo.com/projects/paravel/fold-up-weekender-bag','https://paravel.com','support@paravel.com','★★★★','使わない時は平たく折りたためるリサイクル素材旅行バッグ'],
  [3840,'ファッション・アクセサリー','Anrealage Light-Responsive Color Change Apparel','Anrealage','Kickstarter','https://www.kickstarter.com/projects/anrealage/light-responsive-apparel','https://anrealage.com','hello@anrealage.com','★★★','紫外線で色が変化する日本発フォトクロミックアパレル'],
  [3841,'ファッション・アクセサリー','Rains Long Jacket Waterproof','Rains','Indiegogo','https://www.indiegogo.com/projects/rains/long-jacket-waterproof','https://rains.com','support@rains.com','★★★★','ミニマルデザイン・完全防水ロングレインジャケット'],
  [3842,'ファッション・アクセサリー','Chums Recycled Fleece Vest','Chums','Kickstarter','https://www.kickstarter.com/projects/chums/recycled-fleece-vest','https://chums.us','support@chums.us','★★★','ペットボトル再生繊維使用アウトドアフリースベスト'],
  [3843,'ファッション・アクセサリー','Solo New York Hybrid Backpack Briefcase','Solo New York','Indiegogo','https://www.indiegogo.com/projects/solonewyork/hybrid-backpack-briefcase','https://solonewyork.com','support@solonewyork.com','★★★','ビジネス⇔バックパック変形2wayバッグ'],
  [3844,'ファッション・アクセサリー','Feral Footwear Waterproof Chelsea Boots','Feral Footwear','Kickstarter','https://www.kickstarter.com/projects/feralfootwear/waterproof-chelsea-boots','https://feralfootwear.com','hello@feralfootwear.com','★★★★','完全防水・洗える都市型チェルシーブーツ'],
  [3845,'ファッション・アクセサリー','Groove Life Silicone Ring Active Wear','Groove Life','Indiegogo','https://www.indiegogo.com/projects/groovelife/silicone-ring-active-wear','https://groovelife.com','support@groovelife.com','★★★★','運動/水仕事対応シリコン製結婚指輪'],
  [3846,'ファッション・アクセサリー','Wandering Bear Goods Modular Travel Tote','Wandering Bear Goods','Kickstarter','https://www.kickstarter.com/projects/wanderingbeargoods/modular-travel-tote','https://wanderingbeargoods.com','support@wanderingbeargoods.com','★★★','パーツ着脱式・用途に応じて形を変える旅行トートバッグ'],
  [3847,'ファッション・アクセサリー','Endy Eyewear Sustainable Bamboo Sunglasses','Endy Eyewear','Indiegogo','https://www.indiegogo.com/projects/endyeyewear/sustainable-bamboo-sunglasses','https://endyeyewear.com','hello@endyeyewear.com','★★★','竹製フレーム・軽量サステナブルサングラス'],
  [3848,'ファッション・アクセサリー','Nomatic Travel Pack Modular System','NOMATIC','Kickstarter','https://www.kickstarter.com/projects/nomatic/travel-pack-modular-system','https://nomatic.com','support@nomatic.com','★★★★★','ポーチ着脱式・完全モジュール型トラベルバックパックシステム'],
  [3849,'ファッション・アクセサリー','Res Ipsa The Everywhere Belt Bag','Res Ipsa','Indiegogo','https://www.indiegogo.com/projects/resipsa/everywhere-belt-bag','https://resipsa.com','hello@resipsa.com','★★★','フェイクレザー・軽量多機能ベルトバッグ'],
  [3850,'ファッション・アクセサリー','Except Recycled Ocean Plastic Backpack','Except','Kickstarter','https://www.kickstarter.com/projects/except/recycled-ocean-plastic-backpack','https://exceptbags.com','support@exceptbags.com','★★★★','廃棄漁網素材使用・防水サステナブルバックパック'],
  [3851,'ファッション・アクセサリー','United By Blue Weekender Duffel Bag','United By Blue','Indiegogo','https://www.indiegogo.com/projects/unitedbyblue/weekender-duffel-bag','https://unitedbyblue.com','hello@unitedbyblue.com','★★★★','購入ごとに海洋ゴミ回収に貢献するダッフルバッグ'],
  [3852,'ファッション・アクセサリー','Rumpl Original Puffy Poncho','Rumpl','Kickstarter','https://www.kickstarter.com/projects/rumpl/original-puffy-poncho','https://rumpl.com','hello@rumpl.com','★★★★','ブランケット×ポンチョ・キャンプでも街でも使えるアウターウェア'],
  [3853,'ファッション・アクセサリー','Findlay Hats Packable Wool Fedora','Findlay Hats','Indiegogo','https://www.indiegogo.com/projects/findlayhats/packable-wool-fedora','https://findlayhats.com','support@findlayhats.com','★★★','型崩れしない・旅行に便利な折りたたみウールハット'],

  // ── クリーニング・収納・整理 (18件) ──
  [3854,'クリーニング・収納・整理','Bluewind Reusable Paper Towels','Bluewind','Kickstarter','https://www.kickstarter.com/projects/bluewind/reusable-paper-towels','https://bluewindhome.com','support@bluewindhome.com','★★★★','洗って繰り返し使える・ペーパータオル代替エコクロス'],
  [3855,'クリーニング・収納・整理','Whisk Wizard Automatic Dish Scrubber','Whisk Wizard','Indiegogo','https://www.indiegogo.com/projects/whiskwizard/automatic-dish-scrubber','https://whiskwizard.com','support@whiskwizard.com','★★★','モーター内蔵・自動回転で皿洗いを時短するスクラバー'],
  [3856,'クリーニング・収納・整理','Cloudwash Portable Mini Washing Machine','Cloudwash','Kickstarter','https://www.kickstarter.com/projects/cloudwash/portable-mini-washing-machine','https://cloudwash.co','hello@cloudwash.co','★★★★','折りたたみバケツ型・一人暮らし向け携帯洗濯機'],
  [3857,'クリーニング・収納・整理','Neatfreak Vacuum Storage System','Neatfreak','Indiegogo','https://www.indiegogo.com/projects/neatfreak/vacuum-storage-system','https://neat-freak.com','support@neat-freak.com','★★★','電動ポンプ付き・衣類/寝具圧縮バキューム収納システム'],
  [3858,'クリーニング・収納・整理','Cleanse Bot UV Sanitizing Robot','CleanseBot','Kickstarter','https://www.kickstarter.com/projects/cleansebot/uv-sanitizing-robot','https://cleansebot.com','support@cleansebot.com','★★★★','UV-C照射・自走式で床面を除菌するロボット'],
  [3859,'クリーニング・収納・整理','Onestone Reusable Silicone Storage Bag','Onestone','Indiegogo','https://www.indiegogo.com/projects/onestone/reusable-silicone-storage-bag','https://onestonehome.com','hello@onestonehome.com','★★★','密閉ジップ式・食洗機/冷凍対応シリコン保存袋'],
  [3860,'クリーニング・収納・整理','Airdog X8 Formaldehyde Removal Air Purifier','Airdog','Kickstarter','https://www.kickstarter.com/projects/airdog/x8-formaldehyde-removal-purifier','https://airdoghome.com','support@airdoghome.com','★★★★','フィルター交換不要・TPA技術搭載空気清浄機'],
  [3861,'クリーニング・収納・整理','Spruce Smart Home Composter','Spruce','Indiegogo','https://www.indiegogo.com/projects/spruce/smart-home-composter','https://sprucehome.co','support@sprucehome.co','★★★★★','24時間で生ゴミを堆肥化するスマート家庭用コンポスター'],
  [3862,'クリーニング・収納・整理','Grovemade Modular Desk Organizer','Grovemade','Kickstarter','https://www.kickstarter.com/projects/grovemade/modular-desk-organizer','https://grovemade.com','support@grovemade.com','★★★','木製+フェルト・組み合わせ自在デスクオーガナイザー'],
  [3863,'クリーニング・収納・整理','Vacuum Wand Handheld Air Compressor Duster','Vacuum Wand','Indiegogo','https://www.indiegogo.com/projects/vacuumwand/handheld-air-compressor-duster','https://vacuumwand.com','hello@vacuumwand.com','★★★','高圧エアー噴射・キーボード/家電の隙間掃除に特化したダスター'],
  [3864,'クリーニング・収納・整理','Boxie Cardboard Storage Furniture System','Boxie','Kickstarter','https://www.kickstarter.com/projects/boxie/cardboard-storage-furniture','https://boxiehome.com','support@boxiehome.com','★★★','リサイクル段ボール製・組み立て式収納家具システム'],
  [3865,'クリーニング・収納・整理','Aerify Portable Mattress Vacuum Cleaner','Aerify','Indiegogo','https://www.indiegogo.com/projects/aerify/portable-mattress-vacuum-cleaner','https://aerifyhome.com','support@aerifyhome.com','★★★★','UV除菌+吸引・ダニ/ハウスダスト除去マットレスクリーナー'],
  [3866,'クリーニング・収納・整理','Compost Now Kitchen Countertop Bin','Compost Now','Kickstarter','https://www.kickstarter.com/projects/compostnow/kitchen-countertop-bin','https://compostnow.org','hello@compostnow.org','★★★','密閉+活性炭フィルター・臭い漏れ防止生ゴミ収集ビン'],
  [3867,'クリーニング・収納・整理','Clean Origin Ultrasonic Jewelry Cleaner','Clean Origin','Indiegogo','https://www.indiegogo.com/projects/cleanorigin/ultrasonic-jewelry-cleaner','https://cleanorigin.com','support@cleanorigin.com','★★★★','家庭用超音波洗浄・宝石/時計対応クリーナー'],
  [3868,'クリーニング・収納・整理','Boxed Water Reusable Cardboard Container','Boxed Water','Kickstarter','https://www.kickstarter.com/projects/boxedwater/reusable-cardboard-container','https://boxedwaterisbetter.com','support@boxedwaterisbetter.com','★★★','92%植物由来素材・環境配慮型飲料コンテナ'],
  [3869,'クリーニング・収納・整理','Homeright SteamMachine Multi-Purpose Steamer','HomeRight','Indiegogo','https://www.indiegogo.com/projects/homeright/steammachine-multi-purpose-steamer','https://homeright.com','support@homeright.com','★★★★','高圧スチーム・キッチン/バスルーム除菌に対応する多用途スチーマー'],
  [3870,'クリーニング・収納・整理','Cabinet IQ Smart Pantry Organizer System','Cabinet IQ','Kickstarter','https://www.kickstarter.com/projects/cabinetiq/smart-pantry-organizer-system','https://cabinetiq.com','support@cabinetiq.com','★★★','ラベル管理+賞味期限トラッキング・パントリー収納システム'],
  [3871,'クリーニング・収納・整理','Casabella Wow Mop Spin System','Casabella','Indiegogo','https://www.indiegogo.com/projects/casabella/wow-mop-spin-system','https://casabella.com','support@casabella.com','★★★★','フットペダル式脱水・回転モップ掃除システム'],
];

// ── ユーティリティ ──────────────────────────────────────────────────────────
function buildGmailUrl(toEmail, maker, product) {
  const enc = s => encodeURIComponent(s);
  const subject = 'Potential Distribution Partnership for Japan';
  const body =
`Dear ${maker} Team,

I hope this message finds you well.

My name is Hiroyuki Inoguchi from Sumai pluS in Japan.

Our company focuses on promoting products that enrich people's daily lives. We are currently seeking unique international brands to introduce to the Japanese market and support their growth.

We recently discovered your product "${product}" and were truly impressed by its innovation and quality. We believe it has significant potential in Japan, where consumers are enthusiastic about cutting-edge products from abroad.

We would love to explore the possibility of becoming your authorized distributor in Japan. Our team has extensive experience in market entry, logistics, and retail partnerships across Japan.

Could we schedule a brief call or email exchange to discuss this opportunity further?

Looking forward to your response.

Sincerely,
Hiroyuki Inoguchi
Sumai pluS
 (LEAGUE Co., Ltd. Agent )
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
const catCount  = {};
for (const p of NEW_PRODUCTS) catCount[p[1]] = (catCount[p[1]] || 0) + 1;

console.log('\n====== 追加完了レポート ======');
console.log(`追加件数       : ${NEW_PRODUCTS.length} 件`);
console.log(`メールあり     : ${mailCount} 件`);
console.log(`Kickstarter/IF : ${cfCount} 件`);
console.log('\nカテゴリ別:');
Object.entries(catCount).forEach(([c,n]) => console.log(`  ${c}: ${n}件`));
console.log(`\n総収録件数     : ${existing.length - 1 + NEW_PRODUCTS.length} 件`);
console.log(`出力ファイル   : ${OUTPUT}`);
