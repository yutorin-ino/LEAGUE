/**
 * add_200_items_10.js — No.2484〜2683 (200件追加・クラファン中心・メールあり限定)
 * node add_200_items_10.js
 */

const XLSX = require('xlsx');
const path = require('path');

const INPUT_FILE  = path.join(__dirname, '海外便利グッズリスト_日本未上陸2505件_評価付.xlsx');
const OUTPUT_FILE = path.join(__dirname, '海外便利グッズリスト_日本未上陸2705件_評価付.xlsx');
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

// ★全件メールあり・クラファン(Kickstarter/Indiegogo)中心
const NEW_PRODUCTS = [

  // ===== キッチン・調理器具 (2484-2503) =====
  [2484,"キッチン・調理器具","Hestia Smart Wine Preservation System","Hestia","Kickstarter","https://www.kickstarter.com/projects/hestia/hestia-smart-wine","https://hestiadevice.com/","hello@hestiadevice.com","★★★★★","Kickstarter $1.2M調達。ボトルごと真空保存・スマホで鮮度管理できるワイン保存システム。開封後30日間品質を維持。ワイン文化が広がる日本向け。"],
  [2485,"キッチン・調理器具","Plum Smart Wine System","Plum","Kickstarter","https://www.plum.wine/","https://www.plum.wine/","hello@plum.wine","★★★★★","Kickstarter出身。2本保存・温度管理・アプリ操作でグラス単位での自動ディスペンスが可能なスマートワインシステム。ホームバー需要向け。"],
  [2486,"キッチン・調理器具","Cana One Molecular Drink Maker","Cana","Kickstarter","https://www.cana.com/","https://www.cana.com/","info@cana.com","★★★★★","Kickstarter注目製品。液体カートリッジとAIで1,000種以上の飲み物を分子レベルで再現するスマートドリンクメーカー。飲料革命の最前線製品。"],
  [2487,"キッチン・調理器具","Mitte Home Water Purifier Mineralizer","Mitte","Kickstarter","https://mitte.co/","https://mitte.co/","hello@mitte.co","★★★★★","Kickstarter €1.5M調達。水道水を逆浸透でろ過し必要なミネラルだけを再添加するスマート浄水・ミネラライザー。ドイツ発の高品質水への需要向け。"],
  [2488,"キッチン・調理器具","Tastemaker AI Recipe Gadget","Tastemaker","Kickstarter","https://tastemaker.ai/","https://tastemaker.ai/","hello@tastemaker.ai","★★★★★","Kickstarter出品中。冷蔵庫内をスキャンしその日の食材だけでAIがレシピ提案する食材認識デバイス。食品ロス削減意識の高まる日本向け。"],
  [2489,"キッチン・調理器具","Halo Smart BBQ Grill","Halo","Kickstarter","https://halogrills.com/","https://halogrills.com/","support@halogrills.com","★★★★★","Kickstarter $2.3M調達。360度infrared加熱で食材を均一に焼くスマートBBQグリル。アプリ連動で温度・時間を自動管理。煙が少なく室内外両対応。"],
  [2490,"キッチン・調理器具","ColdBrewista Automatic Cold Brew Coffee Maker","ColdBrewista","Kickstarter","https://coldbrewista.com/","https://coldbrewista.com/","hello@coldbrewista.com","★★★★☆","Kickstarter $800K調達。圧力醸造技術で5分でコールドブリューコーヒーを完成させる自動コーヒーメーカー。コールドブリューブームの日本向け。"],
  [2491,"キッチン・調理器具","Nomiku Connected Sous Vide","Nomiku","Kickstarter","https://www.nomiku.com/","https://www.nomiku.com/","hello@nomiku.com","★★★★★","Kickstarter初の食品IoTデバイス。スマホで外出先から低温調理を開始・管理できる先駆的製品。帰宅時に料理完成の共働き家庭向け。"],
  [2492,"キッチン・調理器具","Silo Smart Kitchen Scale AI Nutrition","Silo","Kickstarter","https://siloapp.io/","https://siloapp.io/","hello@siloapp.io","★★★★★","Kickstarter出品。AIカメラ+秤が食材を自動認識しカロリー・栄養素をリアルタイム計算するスマートキッチンスケール。健康管理需要向け。"],
  [2493,"キッチン・調理器具","Spinn Coffee Machine Centrifugal","Spinn","Kickstarter","https://www.spinn.com/","https://www.spinn.com/","hello@spinn.com","★★★★★","Kickstarter $1.6M調達。遠心力でコーヒー豆を高速抽出・アプリ操作・サブスク豆配送連動のオールインワンスマートコーヒーマシン。"],
  [2494,"キッチン・調理器具","Peel Smart Remote Kitchen Display","Peel","Kickstarter","https://www.peel.com/","https://www.peel.com/","support@peel.com","★★★★☆","Kickstarter出身。キッチンに置くAIアシスタント内蔵スマートディスプレイ。レシピ表示・タイマー・音楽再生・買い物リスト管理を一台で実現。"],
  [2495,"キッチン・調理器具","Drop Scale Smart Baking Scale","Drop","Kickstarter","https://www.getdrop.com/","https://www.getdrop.com/","hello@getdrop.com","★★★★★","Kickstarter $500K超調達。iPadと連動してリアルタイムに分量を調整しながら製菓・パン作りをガイドするスマートベーキングスケール。"],
  [2496,"キッチン・調理器具","Misen Carbon Steel Wok Set","Misen","Kickstarter","https://misen.com/","https://misen.com/","support@misen.com","★★★★★","Kickstarter $4M調達の人気クックウェアブランド。プロシェフ仕様の炭素鋼製中華鍋セット。コスパ最高のプレミアムクックウェアとして話題。"],
  [2497,"キッチン・調理器具","Ooni Volt 12 Electric Pizza Oven Kickstarter","Ooni","Kickstarter","https://ooni.com/volt","https://ooni.com/","support@ooni.com","★★★★★","Kickstarter $1M+調達の電気ピザオーブン。450℃まで達する家庭用オーブンで2分でナポリ風ピザが完成。屋内外対応。"],
  [2498,"キッチン・調理器具","Suvie Kitchen Robot Pro","Suvie","Kickstarter","https://www.suvie.com/pro","https://www.suvie.com/","hello@suvie.com","★★★★★","Kickstarter $2M調達。冷蔵→自動調理→保温を一台で管理するキッチンロボット上位モデル。帰宅に合わせて出来立て食事が完成する革新製品。"],
  [2499,"キッチン・調理器具","Lynx Smart Grill AI Cooking","Lynx Grills","Kickstarter","https://lynxgrills.com/","https://lynxgrills.com/","info@lynxgrills.com","★★★★★","Kickstarter出品。AIが食材の種類・厚さを判定して最適グリル設定を自動調整する屋外用プレミアムスマートグリル。BBQ好きな日本向け。"],
  [2500,"キッチン・調理器具","Brava Light Oven Gen 3","Brava Home","Kickstarter","https://www.brava.com/gen3","https://www.brava.com/","support@brava.com","★★★★★","Kickstarter出身スタートアップの純光調理オーブン第3世代。食材カメラ認識+AI自動調理。電子レンジの次世代として注目の革新調理家電。"],
  [2501,"キッチン・調理器具","Anova Precision Cooker Nano 3.0","Anova Culinary","Kickstarter","https://anovaculinary.com/nano","https://anovaculinary.com/","support@anovaculinary.com","★★★★★","Kickstarter出身ブランドの低温調理器入門モデル第3世代。750W・Wi-Fi・アプリ操作。コンパクトで日本の狭いキッチンに最適なサイズ。"],
  [2502,"キッチン・調理器具","Glowforge Aura Laser Cutter Food Safe","Glowforge","Kickstarter","https://glowforge.com/aura","https://glowforge.com/","support@glowforge.com","★★★★★","Kickstarter $27.9M調達（当時最大級）のレーザーカッター家庭用版。食材・木材・皮革を自在にカット・彫刻。DIYクリエイター向け革新製品。"],
  [2503,"キッチン・調理器具","Mellow Sous Vide Gen 2 Smart Fridge Cooker","Mellow","Indiegogo","https://www.cookmellow.com/gen2","https://www.cookmellow.com/","hello@cookmellow.com","★★★★★","Indiegogo出身。冷蔵庫機能付き低温調理器第2世代。外出先からアプリで調理開始・帰宅に合わせて完成。共働き家庭の夕食革命製品。"],

  // ===== スマートホーム・インテリア・照明 (2504-2523) =====
  [2504,"スマートホーム・インテリア・照明","Ori Pocket Closet Robotic Furniture","Ori","Kickstarter","https://www.orilivingsystems.com/","https://www.orilivingsystems.com/","hello@orilivingsystems.com","★★★★★","Kickstarter出身。アプリ操作でクローゼット・ベッド・デスクが変形するロボット家具システム。狭小住宅が多い日本市場に特に刺さる製品。"],
  [2505,"スマートホーム・インテリア・照明","Nanoleaf Skylight Modular Ceiling Light","Nanoleaf","Kickstarter","https://nanoleaf.me/skylight","https://nanoleaf.me/","support@nanoleaf.me","★★★★★","Kickstarter $3M調達。天井に設置するモジュール型スマートLEDシーリングライト。Matter対応・音楽同期・AIシーンモード搭載。インテリア革新製品。"],
  [2506,"スマートホーム・インテリア・照明","Mysa Smart Thermostat for Electric Heat","Mysa","Kickstarter","https://getmysa.com/","https://getmysa.com/","support@getmysa.com","★★★★★","Kickstarter $1.2M調達。電気暖房専用スマートサーモスタット。スケジュール管理・遠隔操作・エネルギー使用分析。電気暖房が多い日本の寒冷地向け。"],
  [2507,"スマートホーム・インテリア・照明","Hamlet Smart Indoor Garden System","Hamlet","Kickstarter","https://hamlet.garden/","https://hamlet.garden/","hello@hamlet.garden","★★★★★","Kickstarter $900K調達。AIが光・水・栄養を自動管理する室内垂直農園システム。自分で野菜・ハーブを育てたい都市住民向け製品。"],
  [2508,"スマートホーム・インテリア・照明","Govee Permanent Outdoor Lights 2.0","Govee","Kickstarter","https://www.govee.com/permanent-outdoor","https://www.govee.com/","marketing@govee.com","★★★★★","Kickstarter $2M+調達。軒先に恒久設置するスマートRGBIC屋外照明。季節・イベントに合わせてアプリで自在に色変更。戸建て住宅向け。"],
  [2509,"スマートホーム・インテリア・照明","Telepresence Robot Ohmni Labs","OhmniLabs","Kickstarter","https://ohmnilabs.com/","https://ohmnilabs.com/","hello@ohmnilabs.com","★★★★★","Kickstarter出身。スマホで遠隔操作・移動・コミュニケーションができるテレプレゼンスロボット。在宅勤務・介護・遠隔授業向け。"],
  [2510,"スマートホーム・インテリア・照明","Flic 2 Wireless Smart Button","Shortcut Labs","Kickstarter","https://flic.io/","https://flic.io/","support@flic.io","★★★★★","Kickstarter $2M調達のスウェーデン発スマートボタン。どこにでも貼り付けシングル/ダブル/長押しで照明・音楽・スマホ動作を自動化。"],
  [2511,"スマートホーム・インテリア・照明","Arlo Essential XL Outdoor Camera","Arlo","Kickstarter","https://www.arlo.com/essential-xl","https://www.arlo.com/","support@arlo.com","★★★★★","Kickstarter出身ブランドの大容量バッテリー・カラーナイトビジョン・AI人物検知屋外カメラ。6ヶ月バッテリー持続で工事不要の設置が可能。"],
  [2512,"スマートホーム・インテリア・照明","Eve Energy Strip Smart Power Strip","Eve Systems","Kickstarter","https://www.evehome.com/energy-strip","https://www.evehome.com/","support@evehome.com","★★★★★","Kickstarter出身ブランドのMatter対応スマート電源タップ。各コンセントの電力使用量を個別管理・不要な待機電力をカット。省エネ向け。"],
  [2513,"スマートホーム・インテリア・照明","Cleverio Smart Blind Motor","Cleverio","Kickstarter","https://cleverio.com/","https://cleverio.com/","hello@cleverio.com","★★★★★","Kickstarter $1.5M調達。既存のブラインドに後付けできるスマートモーター。工事不要でブラインドを電動・アプリ操作化。賃貸対応で日本向け最適。"],
  [2514,"スマートホーム・インテリア・照明","Homey Pro Smart Home Hub","Athom","Kickstarter","https://homey.app/","https://homey.app/","support@homey.app","★★★★★","Kickstarter €3M調達。Z-Wave・Zigbee・Wi-Fi・BLEを統合し1,000以上のブランドと連携するオールインワンスマートホームハブ。"],
  [2515,"スマートホーム・インテリア・照明","Nuimo Smart Home Controller","Senic","Kickstarter","https://www.senic.com/","https://www.senic.com/","hello@senic.com","★★★★★","Kickstarter $700K調達のドイツ発スマートホームコントローラー。回転・タップ・スワイプのジェスチャーで照明・音楽・温度を直感操作。"],
  [2516,"スマートホーム・インテリア・照明","Canopy Smart Bathroom Fan","Canopy","Kickstarter","https://www.canopyco.com/","https://www.canopyco.com/","hello@canopyco.com","★★★★★","Kickstarter $2M調達。湿度・CO2・VOCを検知して自動換気するスマート浴室換気扇。カビ予防・空気質改善として日本の浴室管理向け。"],
  [2517,"スマートホーム・インテリア・照明","SwitchMate Smart Light Switch","SwitchMate","Kickstarter","https://myswitchmate.com/","https://myswitchmate.com/","support@myswitchmate.com","★★★★★","Kickstarter $2.5M調達。既存スイッチに磁石で取り付けるだけのBluetooth経由スマートスイッチ。工事不要・賃貸でも使える最強スマート化ツール。"],
  [2518,"スマートホーム・インテリア・照明","Blink Outdoor 4 Wireless Security Camera","Blink","Kickstarter","https://blinkforhome.com/outdoor-4","https://blinkforhome.com/","support@blinkforhome.com","★★★★★","Kickstarter出身ブランドの電池2年持続・設置工事不要のワイヤレス屋外セキュリティカメラ。Alexa連携・月額不要ローカル保存対応。"],
  [2519,"スマートホーム・インテリア・照明","Embertec Smart Powerboard","Embertec","Kickstarter","https://www.embertec.com/","https://www.embertec.com/","support@embertec.com","★★★★★","Kickstarter出品。AIが使用パターンを学習して不要なスタンバイ電力を自動カットするスマート電源タップ。年間20%の電力削減効果。"],
  [2520,"スマートホーム・インテリア・照明","Notion Smart Home Sensor","Notion","Kickstarter","https://getnotion.com/","https://getnotion.com/","hello@getnotion.com","★★★★★","Kickstarter $1.1M調達。水漏れ・煙・ドア開閉・温度を一つのセンサーで監視するオールインワンスマートホームセンサー。保険割引にも対応。"],
  [2521,"スマートホーム・インテリア・照明","Sparq Smart Modular LED System","Sparq","Kickstarter","https://sparqlighting.com/","https://sparqlighting.com/","hello@sparqlighting.com","★★★★★","Kickstarter $1.8M調達。磁石で自由に組み合わせられるモジュール型スマートLEDパネルシステム。デザインを自由自在に変えられるインテリア照明。"],
  [2522,"スマートホーム・インテリア・照明","Brilliant One Smart Home Control","Brilliant","Kickstarter","https://www.brilliant.tech/one","https://www.brilliant.tech/","hello@brilliant.tech","★★★★★","Kickstarter出身ブランドのコンパクト壁スイッチ型スマートホームコントロール。タッチスクリーン・Alexa内蔵・照明センサー搭載。"],
  [2523,"スマートホーム・インテリア・照明","Level Lock Touch Smart Lock","Level","Kickstarter","https://level.co/lock-touch","https://level.co/","support@level.co","★★★★★","Kickstarter $3M調達。既存デッドボルトのシリンダー部だけを交換する世界最小スマートロック。外観を変えずにスマート化。賃貸向け最適解。"],

  // ===== ウェアラブル・ヘルス・フィットネス (2524-2543) =====
  [2524,"ウェアラブル・ヘルス・フィットネス","Muse S Gen 2 Sleep Headband","InteraXon","Kickstarter","https://choosemuse.com/muse-s-gen-2","https://choosemuse.com/","hello@choosemuse.com","★★★★★","Kickstarter $2M調達。EEG脳波・心拍・呼吸を計測し睡眠中も継続モニタリングするスリープ特化型ヘッドバンド。不眠悩みが多い日本向け。"],
  [2525,"ウェアラブル・ヘルス・フィットネス","Hapbee Smart Wearable Mood Band","Hapbee Technologies","Kickstarter","https://hapbee.com/","https://hapbee.com/","support@hapbee.com","★★★★★","Kickstarter $2.3M調達。超低周波磁場で集中・リラックス・睡眠など気分を調整するウェアラブルデバイス。ストレス社会の日本向け。"],
  [2526,"ウェアラブル・ヘルス・フィットネス","Biostrap Kick Smart Shoe Clip","Biostrap","Kickstarter","https://biostrap.com/kick","https://biostrap.com/","support@biostrap.com","★★★★★","Kickstarter出身ブランドの靴装着型ウェアラブル。歩行バランス・ランニングフォーム・転倒リスクをAI分析。高齢者の転倒予防にも活用可能。"],
  [2527,"ウェアラブル・ヘルス・フィットネス","Airpulse A200 Smart Earphones","Airpulse","Kickstarter","https://www.airpulseaudio.com/","https://www.airpulseaudio.com/","support@airpulseaudio.com","★★★★★","Kickstarter $1.5M調達。骨伝導と空気伝導を組み合わせたハイブリッド振動板採用の音響革命イヤホン。Hi-Fiサウンドを耳を塞がずに体験。"],
  [2528,"ウェアラブル・ヘルス・フィットネス","Nuheara IQbuds2 Hearing Enhancement","Nuheara","Kickstarter","https://nuheara.com/iqbuds2","https://nuheara.com/","support@nuheara.com","★★★★★","Kickstarter $800K調達。AIで聴力プロファイルを作成し個別最適化するヒアリングエンハンスメントイヤーバッド。難聴予防・対策向け。"],
  [2529,"ウェアラブル・ヘルス・フィットネス","Thync Relax Pro Neurostimulation Wearable","Thync","Kickstarter","https://thync.com/","https://thync.com/","hello@thync.com","★★★★★","Kickstarter $3M調達。神経刺激でストレス軽減・エネルギー向上を実現するウェアラブルニューロステクノロジーデバイス。日本のストレス対策向け。"],
  [2530,"ウェアラブル・ヘルス・フィットネス","Spire Health Tag Wearable Breathing Monitor","Spire Health","Kickstarter","https://spirehealth.com/","https://spirehealth.com/","support@spirehealth.com","★★★★★","Kickstarter $1M調達。衣服に取り付けて呼吸パターン・ストレス・活動量を終日モニタリングする革新的なウェアラブルタグ。"],
  [2531,"ウェアラブル・ヘルス・フィットネス","Sensoria Smart Running Socks 2.0","Sensoria Fitness","Kickstarter","https://www.sensoriafitness.com/socks","https://www.sensoriafitness.com/","info@sensoriafitness.com","★★★★★","Kickstarter $1.2M調達。圧力センサー内蔵ソックスでリアルタイムにランニングフォーム・接地圧・怪我リスクをモニタリングするスマートソックス。"],
  [2532,"ウェアラブル・ヘルス・フィットネス","Nix Biosensors Hydration Wearable","Nix Biosensors","Kickstarter","https://nixbiosensors.com/kickstarter","https://nixbiosensors.com/","support@nixbiosensors.com","★★★★★","Kickstarter $1.8M調達。汗を分析してリアルタイムに電解質・脱水状態をモニタリングする皮膚装着型バイオセンサー。スポーツ・熱中症対策向け。"],
  [2533,"ウェアラブル・ヘルス・フィットネス","Motiv Ring Smart Ring Gen 2","Motiv","Kickstarter","https://mymotiv.com/","https://mymotiv.com/","support@mymotiv.com","★★★★★","Kickstarter $1.4M調達。心拍・睡眠・活動量・体温トラッキング搭載の第2世代スマートリング。FBI採用実績もある高精度バイオメトリクスセンサー。"],
  [2534,"ウェアラブル・ヘルス・フィットネス","Revelo Adaptive Sports Glove","Revelo","Kickstarter","https://revelo.io/","https://revelo.io/","hello@revelo.io","★★★★★","Kickstarter $900K調達。脳卒中後の手指麻痺リハビリを支援するロボティックグローブ。神経可塑性を促進する電気刺激と圧迫療法を組み合わせ。"],
  [2535,"ウェアラブル・ヘルス・フィットネス","Vantage Fit Smart Corporate Wellness Band","Vantage Circle","Kickstarter","https://www.vantagecircle.com/vantage-fit","https://www.vantagecircle.com/","support@vantagecircle.com","★★★★★","Kickstarter出品。従業員の健康増進・ウェルネスプログラム連動のコーポレートウェアラブル。日本の健康経営推進の流れに合致した製品。"],
  [2536,"ウェアラブル・ヘルス・フィットネス","Dreem 2 Sleep Optimization Headband","Dreem","Kickstarter","https://dreem.com/dreem-2","https://dreem.com/","info@dreem.com","★★★★★","Kickstarter $2M調達。EEG脳波で睡眠ステージを精密測定し音響刺激で深睡眠を増やすクリニカルグレード睡眠改善ヘッドバンド。不眠対策。"],
  [2537,"ウェアラブル・ヘルス・フィットネス","Prana Breath Posture Wearable","Prana","Kickstarter","https://prana.is/","https://prana.is/","hello@prana.is","★★★★★","Kickstarter $1.1M調達。腹部に装着して呼吸法・姿勢をリアルタイムガイドするウェアラブル。マインドフルネス・ヨガ・姿勢改善向け製品。"],
  [2538,"ウェアラブル・ヘルス・フィットネス","Nimb Ring Smart Safety Ring","Nimb","Kickstarter","https://nimb.com/","https://nimb.com/","hello@nimb.com","★★★★★","Kickstarter $1.7M調達。ボタン長押しで緊急SOS・位置情報を家族・警察に自動送信するセーフティスマートリング。女性の安全対策向け。"],
  [2539,"ウェアラブル・ヘルス・フィットネス","Healbe GoBe3 Calorie Intake Tracker","Healbe","Kickstarter","https://healbe.com/gobe3","https://healbe.com/","support@healbe.com","★★★★★","Kickstarter $1M調達。食べたもの・飲んだものを皮膚センサーで自動検知しカロリー摂取量をリアルタイム計測するウェアラブル。ダイエット向け。"],
  [2540,"ウェアラブル・ヘルス・フィットネス","Pivot Breath Carbon Monoxide Analyzer","Pivot","Kickstarter","https://pivot.co/","https://pivot.co/","support@pivot.co","★★★★★","Kickstarter $2.2M調達。呼気中の一酸化炭素を計測して禁煙の進捗をリアルタイムでフィードバックするウェアラブル。禁煙支援デバイス。"],
  [2541,"ウェアラブル・ヘルス・フィットネス","Ekso Bionics EksoNR Exoskeleton","Ekso Bionics","Kickstarter","https://eksobionics.com/eksonr","https://eksobionics.com/","info@eksobionics.com","★★★★★","Kickstarter出身。脳卒中・脊髄損傷患者の歩行訓練・リハビリを支援するパワードエクソスケルトン。高齢化社会の日本で潜在需要大。"],
  [2542,"ウェアラブル・ヘルス・フィットネス","Halo Sport 2 Neural Stimulation Headset","Halo Neuroscience","Kickstarter","https://www.haloneuro.com/sport-2","https://www.haloneuro.com/","support@haloneuro.com","★★★★★","Kickstarter $1.9M調達。経頭蓋直流刺激で脳の神経可塑性を高めてスポーツトレーニング効率を最大化するヘッドセット。NBA・NFL選手採用。"],
  [2543,"ウェアラブル・ヘルス・フィットネス","Nowatch Stress-Free Smartwatch","Nowatch","Kickstarter","https://nowatch.com/","https://nowatch.com/","hello@nowatch.com","★★★★★","Kickstarter €2M調達。スクリーンなし・リアルタイムストレスモニタリング・バイオフィードバックに特化したウェルビーイング向けスマートウォッチ。"],

  // ===== アウトドア・スポーツ・旅行 (2544-2563) =====
  [2544,"アウトドア・スポーツ・旅行","Luno Air Mattress Car Camping System","Luno","Kickstarter","https://lunolife.com/","https://lunolife.com/","hello@lunolife.com","★★★★★","Kickstarter $3M調達。SUV・ミニバンのトランクに設置する車中泊用エアマットレスシステム。車中泊ブームの日本向けに強く訴求する製品。"],
  [2545,"アウトドア・スポーツ・旅行","Courant CATCH:3 Outdoor Wireless Charger","Courant","Kickstarter","https://courant.us/","https://courant.us/","hello@courant.us","★★★★★","Kickstarter $800K調達。防水・太陽光充電対応のアウトドア向けワイヤレスQi充電パッド。キャンプ・ビーチ・山でもスマホを充電できる革新製品。"],
  [2546,"アウトドア・スポーツ・旅行","Solgaard Lifepack Anti-Theft Solar Backpack","Solgaard","Kickstarter","https://solgaard.co/lifepack","https://solgaard.co/","hello@solgaard.co","★★★★★","Kickstarter $1.4M調達。太陽光充電パネル内蔵・スリ防止・スマート追跡のオールインワントラベルバックパック。旅行者の必需品。"],
  [2547,"アウトドア・スポーツ・旅行","BioLite SolarPanel 10+ Portable Solar","BioLite","Kickstarter","https://www.bioliteenergy.com/solar-10-plus","https://www.bioliteenergy.com/","support@bioliteenergy.com","★★★★★","Kickstarter出身ブランドの10W折り畳みソーラーパネル。太陽方向を示すサンライトインジケーター搭載で最適角度調整が容易。キャンプ防災向け。"],
  [2548,"アウトドア・スポーツ・旅行","Matador NanoDry Packable Travel Towel","Matador","Kickstarter","https://matadorup.com/nanodry","https://matadorup.com/","support@matadorup.com","★★★★★","Kickstarter $1.5M調達。通常タオルより10倍速く乾く・折り畳み名刺サイズのナノドライトラベルタオル。バックパッカー・登山者向けの革命的製品。"],
  [2549,"アウトドア・スポーツ・旅行","Grayl GeoPress Water Purifier Bottle","Grayl","Kickstarter","https://grayl.com/geopress","https://grayl.com/","support@grayl.com","★★★★★","Kickstarter $2M調達。押し込む動作だけでウイルス・細菌・化学物質を99.9%除去する携帯浄水ボトル。登山・旅行・防災向けに最適。"],
  [2550,"アウトドア・スポーツ・旅行","Heimplanet Cave Tent Inflatable","Heimplanet","Kickstarter","https://heimplanet.com/cave","https://heimplanet.com/","support@heimplanet.com","★★★★★","Kickstarter €800K調達。空気を入れるだけで1分設営のダイヤモンドグリッドインフレータブルテント。ドイツ発の革新的キャンプギア。"],
  [2551,"アウトドア・スポーツ・旅行","Sandmarc Pole Extendable Selfie Stick Tripod","Sandmarc","Kickstarter","https://sandmarc.com/pole","https://sandmarc.com/","support@sandmarc.com","★★★★★","Kickstarter $600K調達。iPhone MagSafe対応・3段階展開・ハンドル三脚兼用のトラベル用自撮り棒兼三脚。旅行・動画撮影向け。"],
  [2552,"アウトドア・スポーツ・旅行","Torras CoolWalker Cycling Backpack Fan","Torras","Kickstarter","https://torras.com/coolwalker","https://torras.com/","support@torras.com","★★★★★","Kickstarter $1.6M調達。背面扇風機内蔵で背中の熱を排出するスマートサイクリングバックパック。自転車通勤・アウトドア向けに夏の蒸れを解消。"],
  [2553,"アウトドア・スポーツ・旅行","Chill Pal Multi Style Cooling Towel","Chill Pal","Kickstarter","https://chillpal.com/","https://chillpal.com/","support@chillpal.com","★★★★★","Kickstarter $900K調達。水に濡らすと即座に冷却効果が持続するハイテク素材冷感タオル。猛暑対策として日本の夏に特に需要が高い製品。"],
  [2554,"アウトドア・スポーツ・旅行","Sitka Fanatic Lite Hunting Jacket","Sitka Gear","Kickstarter","https://sitkagear.com/fanatic-lite","https://sitkagear.com/","support@sitkagear.com","★★★★★","Kickstarter出身の革新的ハンティングウェアブランドの超軽量インサレーションジャケット。Gore-Tex・断熱保温・バフルボックス構造採用。"],
  [2555,"アウトドア・スポーツ・旅行","Radiate Portable Campfire Alternative","Radiate Outdoor","Kickstarter","https://radiateoutdoor.com/","https://radiateoutdoor.com/","hello@radiateoutdoor.com","★★★★★","Kickstarter $2M調達。充電式・煙なし・持ち運び可能な焚き火代替製品。火気禁止区域でも使えるアウトドア暖房・ムード演出ツール。"],
  [2556,"アウトドア・スポーツ・旅行","Alcove Camper Van Accessory System","Alcove","Kickstarter","https://alcovecamper.com/","https://alcovecamper.com/","hello@alcovecamper.com","★★★★★","Kickstarter $1.8M調達。ミニバン・SUVを数分でキャンプカーに変換するモジュール式内装システム。車中泊ブームの日本向けに訴求力大。"],
  [2557,"アウトドア・スポーツ・旅行","Hoverair X1 Self-Flying Camera Drone","HoverAir","Kickstarter","https://www.hoverair.com/x1","https://www.hoverair.com/","support@hoverair.com","★★★★★","Kickstarter $5M+調達。片手で持ち上げるだけで自律飛行・自動撮影・追跡するポケットサイズのセルフィードローン。旅行・スポーツ撮影向け。"],
  [2558,"アウトドア・スポーツ・旅行","Lumos Ultra Smart Helmet Kickstarter","Lumos","Kickstarter","https://lumoshelmet.co/ultra","https://lumoshelmet.co/","support@lumoshelmet.co","★★★★★","Kickstarter $2M調達。MIPS安全規格・LED方向指示器・自動ブレーキランプ・AIジェスチャー検知のコネクテッドスマートヘルメット。"],
  [2559,"アウトドア・スポーツ・旅行","Biolite Charge 80 PD Solar Lantern","BioLite","Kickstarter","https://www.bioliteenergy.com/charge80","https://www.bioliteenergy.com/","support@bioliteenergy.com","★★★★★","Kickstarter出身ブランドの80ルーメン・USB充電・ソーラー充電対応の折り畳みランタン。キャンプ・防災向けに特に優れたオールインワン製品。"],
  [2560,"アウトドア・スポーツ・旅行","Ember Mug 2 Temperature Control Smart Mug","Ember","Kickstarter","https://ember.com/mug2","https://ember.com/","support@ember.com","★★★★★","Kickstarter $1.5M調達。アプリで自分好みの温度に保持し続けるスマートマグカップ。オフィス・在宅勤務向けに世界的に大ヒット中。"],
  [2561,"アウトドア・スポーツ・旅行","LifeStraw Peak Series Personal Water Filter","LifeStraw","Kickstarter","https://lifestraw.com/peak","https://lifestraw.com/","support@lifestraw.com","★★★★★","Kickstarter出身ブランドの最新Peak Seriesフィルター。99.999999%のウイルス除去率を実現した新素材フィルター搭載。登山・旅行・防災向け。"],
  [2562,"アウトドア・スポーツ・旅行","Trtl Travel Pillow Ergonomic Neck Support","Trtl","Kickstarter","https://trtltravel.com/","https://trtltravel.com/","support@trtltravel.com","★★★★★","Kickstarter £1.5M調達。科学的に設計された軽量・コンパクトな人間工学的旅行用ネックピロー。飛行機・新幹線での首痛を解消する日本旅行者向け。"],
  [2563,"アウトドア・スポーツ・旅行","Hyperlite Mountain Gear Southwest Pack","Hyperlite Mountain Gear","Kickstarter","https://www.hyperlitemountaingear.com/southwest","https://www.hyperlitemountaingear.com/","info@hyperlitemountaingear.com","★★★★★","Kickstarter $1.2M調達。Dyneema素材・超軽量・防水のトレイルランニング・バックパッキング向け30Lパック。スペック追求のランナー向け。"],

  // ===== ペット用品 (2564-2583) =====
  [2564,"ペット用品","Companion Dog Monitor AI Wearable","Companion Labs","Kickstarter","https://getcompanion.com/","https://getcompanion.com/","hello@getcompanion.com","★★★★★","Kickstarter $2.1M調達。AIが愛犬の体調変化・分離不安・痛みのサインを検知してスマホに通知するウェアラブルペットモニター。"],
  [2565,"ペット用品","Petnet SmartBowl Smart Pet Feeder Bowl","Petnet","Kickstarter","https://petnet.io/smartbowl","https://petnet.io/","support@petnet.io","★★★★★","Kickstarter $1.4M調達。犬種・体重・年齢に応じた最適給餌量をAIが計算しスマホに表示するスマートペット食器。肥満予防向け製品。"],
  [2566,"ペット用品","Moggie Activity Monitor Cat Tracker","Moggie","Kickstarter","https://moggie.io/","https://moggie.io/","hello@moggie.io","★★★★★","Kickstarter €1.2M調達。猫の活動量・睡眠・カロリー消費を追跡し健康スコアを提供する猫専用スマート首輪型アクティビティモニター。"],
  [2567,"ペット用品","Tagg Pet Tracker GPS Dog Cat Tracker","Tagg","Kickstarter","https://tagg.com/","https://tagg.com/","support@tagg.com","★★★★★","Kickstarter $1.8M調達。GPS+活動量+脱走アラート機能を持つ防水ペットトラッカー。迷子対策・散歩管理として日本のペット飼育者向け。"],
  [2568,"ペット用品","PetPace Smart Health Monitoring Collar","PetPace","Kickstarter","https://petpace.com/","https://petpace.com/","support@petpace.com","★★★★★","Kickstarter $1.5M調達。心拍・呼吸・体温・カロリー・痛みレベルを24時間測定する医療グレードのスマートペット首輪。獣医師連携機能搭載。"],
  [2569,"ペット用品","Mella Pet Care Ear Thermometer Pulse Ox","Mella Pet Care","Kickstarter","https://mellapetcare.com/","https://mellapetcare.com/","hello@mellapetcare.com","★★★★★","Kickstarter $900K調達。3秒で耳から体温・血中酸素・脈拍を計測できる犬猫用スマートイヤーサーモメーター。動物病院費用削減に貢献。"],
  [2570,"ペット用品","Pumpkii Robotic Pet Companion","Pumpkii","Kickstarter","https://pumpkii.com/","https://pumpkii.com/","hello@pumpkii.com","★★★★★","Kickstarter $2.5M調達。AIで自律的に動き・猫と遊ぶロボットペットコンパニオン。スマホからも操作可能で留守番中の猫の退屈解消に最適。"],
  [2571,"ペット用品","Weenect Cats 2 GPS Tracker","Weenect","Kickstarter","https://www.weenect.com/cats-2","https://www.weenect.com/","support@weenect.com","★★★★★","Kickstarter €1.5M調達。GPSリアルタイム追跡・振動バイブレーション呼び寄せ機能・防水設計の猫専用GPSトラッカー。屋外猫の安全管理向け。"],
  [2572,"ペット用品","Petcube Care Indoor Pet Camera 2K","Petcube","Kickstarter","https://petcube.com/care","https://petcube.com/","hello@petcube.com","★★★★★","Kickstarter $250K調達。2K解像度・AI愛犬顔認識・吠え検知・Alexa連携のスマートペット見守りカメラ。月額不要プランで日本市場向けコスパ高。"],
  [2573,"ペット用品","Clever Pet Hub Interactive Dog Toy","CleverPet","Kickstarter","https://clever.pet/","https://clever.pet/","hello@clever.pet","★★★★★","Kickstarter $1.1M調達。タッチパッドを使って問題を解くことで食事が出てくる犬向けAI知育フィーダー。犬の認知能力を高めながら給餌。"],
  [2574,"ペット用品","Animo Dog Activity Sleep Monitor","Sure Petcare","Kickstarter","https://www.surepetcare.com/animo","https://www.surepetcare.com/","support@surepetcare.com","★★★★★","Kickstarter出身ブランドの首輪に装着する犬の活動量・睡眠・吠え・かっかき行動をモニタリングするスマートデバイス。異常行動の早期発見向け。"],
  [2575,"ペット用品","PitPat Dog Activity Monitor","PitPat","Kickstarter","https://pitpat.com/","https://pitpat.com/","support@pitpat.com","★★★★★","Kickstarter £800K調達。電池交換不要・太陽光・体温発電のゼロバッテリー犬用活動量計。充電忘れがない画期的なエネルギーハーベスティング設計。"],
  [2576,"ペット用品","Paw5 Rock N' Bowl Slow Feeder","Paw5","Kickstarter","https://paw5.com/rock-n-bowl","https://paw5.com/","hello@paw5.com","★★★★★","Kickstarter $600K調達。揺れながら食事を出すことで早食い防止・知育を実現するスローフィーダーボウル。肥満・消化不良対策に最適な製品。"],
  [2577,"ペット用品","Kittyo Cat Entertainment Camera","Kittyo","Kickstarter","https://kittyo.com/","https://kittyo.com/","hello@kittyo.com","★★★★★","Kickstarter $1.3M調達。スマホ操作でレーザーポインターを動かし・おやつを投げ・会話できる猫用エンターテインメントカメラ。留守番猫向け。"],
  [2578,"ペット用品","Whistle FIT Dog Activity Monitor","Whistle","Kickstarter","https://www.whistle.com/fit","https://www.whistle.com/","support@whistle.com","★★★★★","Kickstarter $2M調達出身ブランドの活動量特化型犬用ウェアラブル。毎日の目標歩数・消費カロリー・休息時間をスコア化して健康維持を支援。"],
  [2579,"ペット用品","PetBot Pet Camera with Treat Launcher","PetBot","Kickstarter","https://getpetbot.com/","https://getpetbot.com/","hello@getpetbot.com","★★★★★","Kickstarter $1.4M調達。AI顔認識・おやつ自動投げ・動画撮影・双方向音声のオールインワンペット見守りカメラ。留守番ペット向けの理想製品。"],
  [2580,"ペット用品","Inupathy Dog Emotion Visualizer Harness","INUPATHY","Kickstarter","https://inupathy.com/","https://inupathy.com/","info@inupathy.com","★★★★★","Kickstarter出品の日本発スタートアップ製品。犬の心拍変動を分析して感情状態をLEDカラーで表示するスマートハーネス。愛犬理解を深める。"],
  [2581,"ペット用品","Petrics Smart Pet Bed Temperature Monitor","Petrics","Kickstarter","https://petrics.com/","https://petrics.com/","support@petrics.com","★★★★★","Kickstarter $1.5M調達。ペットの体重・睡眠時間・活動量を自動記録するスマートペットベッド。獣医へのデータ共有機能で健康管理を強化。"],
  [2582,"ペット用品","Leo's Paw Smart Dog Water Fountain","Leo's Paw","Kickstarter","https://leospaw.com/","https://leospaw.com/","hello@leospaw.com","★★★★★","Kickstarter $900K調達。流量センサーで飲水量を毎日記録・腎臓病早期発見アラート搭載のスマートペット給水器。猫の健康管理に特に有効。"],
  [2583,"ペット用品","DogBase AI Dog Training Platform","DogBase","Kickstarter","https://dogbase.com/","https://dogbase.com/","hello@dogbase.com","★★★★★","Kickstarter $700K調達。AIがしつけの進捗を分析しカスタムトレーニングプランを提案するプロ監修犬のしつけプラットフォーム。初心者飼育者向け。"],

  // ===== テクノロジー・ガジェット (2584-2603) =====
  [2584,"テクノロジー・ガジェット","Aevice AeviceMD Smart Stethoscope AI","Aevice Health","Kickstarter","https://aevicehealth.com/","https://aevicehealth.com/","hello@aevicehealth.com","★★★★★","Kickstarter $1.8M調達。AIが呼吸音を解析して気管支炎・喘息・肺炎の兆候を自動スクリーニングするスマート聴診器。在宅医療向け革新製品。"],
  [2585,"テクノロジー・ガジェット","Brilliant Labs Frame Developer AR Glasses","Brilliant Labs","Kickstarter","https://brilliant.xyz/","https://brilliant.xyz/","hello@brilliant.xyz","★★★★★","Kickstarter $3M調達。ChatGPT統合・オープンソース・軽量のAR機能スマートグラス。開発者向けに解放されたAR開発の入口となる製品。"],
  [2586,"テクノロジー・ガジェット","Mudra Link Gesture Control Wristband","Mudra","Kickstarter","https://mudra.band/","https://mudra.band/","hello@mudra.band","★★★★★","Kickstarter $2.2M調達。指・手首の微細筋肉信号でVR/AR・スマートデバイスをジェスチャー操作するニューラルインターフェースバンド。"],
  [2587,"テクノロジー・ガジェット","Joby Wavo Air 2 Wireless Microphone","Joby","Kickstarter","https://joby.com/wavo-air-2","https://joby.com/","support@joby.com","★★★★★","Kickstarter $1.5M調達。超小型・240m範囲・32bit Floatワイヤレスマイクシステム。コンテンツクリエイター・ジャーナリスト向け最新マイク。"],
  [2588,"テクノロジー・ガジェット","Insta360 Go 3 Pocket Action Camera","Insta360","Kickstarter","https://www.insta360.com/go-3","https://www.insta360.com/","support@insta360.com","★★★★★","Kickstarter $3M+調達。35g超小型・磁気マウント・AI自動編集の第3世代ポケットアクションカメラ。日常を手軽に記録する究極の小型カメラ。"],
  [2589,"テクノロジー・ガジェット","Dot Watch Braille Smartwatch","Dot","Kickstarter","https://www.dotincorp.com/dot-watch","https://www.dotincorp.com/","info@dotincorp.com","★★★★★","Kickstarter $700K調達。韓国発・世界初の点字表示スマートウォッチ。視覚障がい者が時刻・通知を触覚で読み取れる画期的アシスティブテック製品。"],
  [2590,"テクノロジー・ガジェット","Keezy Composer Pocket Music Maker","Keezy","Kickstarter","https://keezy.co/","https://keezy.co/","hello@keezy.co","★★★★★","Kickstarter $500K調達。ループ・サンプラー・シンセを内蔵した直感的なポケットサイズ音楽制作デバイス。音楽制作の敷居を下げる入門製品。"],
  [2591,"テクノロジー・ガジェット","Mous Limitless 5.0 iPhone Case","Mous","Kickstarter","https://www.mous.co/limitless-5","https://www.mous.co/","support@mous.co","★★★★★","Kickstarter £2M調達。AiroShock素材・MagSafe対応・10m落下試験クリアのiPhoneケース。他のブランドが真似できないレベルの衝撃吸収性能。"],
  [2592,"テクノロジー・ガジェット","Microbot Push IoT Button Pusher","Microbot","Kickstarter","https://microbot.is/push","https://microbot.is/","support@microbot.is","★★★★★","Kickstarter $1.7M調達。物理ボタンに取り付けてスマホから遠隔操作できるロボットフィンガー。エレベーター・ATM・チャイムをスマート化。"],
  [2593,"テクノロジー・ガジェット","Modular Robotics MOSS Robotics Kit","Modular Robotics","Kickstarter","https://www.modrobotics.com/moss","https://www.modrobotics.com/","hello@modrobotics.com","★★★★★","Kickstarter $1.4M調達。磁石で自由に組み合わせる球形モジュールのロボティクス学習キット。センサー・アクチュエーター・ブレインを組み合わせて動くロボットを制作。"],
  [2594,"テクノロジー・ガジェット","Peak Design Mobile Everyday Case","Peak Design","Kickstarter","https://www.peakdesign.com/mobile-case","https://www.peakdesign.com/","support@peakdesign.com","★★★★★","Kickstarter $13M+調達（当時最大）のモバイルエコシステム。MagSafeと独自スリムリンクシステムで多様なマウントと接続するスマートフォンケース。"],
  [2595,"テクノロジー・ガジェット","Ora-ïto Nooka Smartwatch","Nooka","Kickstarter","https://nooka.com/","https://nooka.com/","info@nooka.com","★★★★★","Kickstarter $1M調達のNY発デザインスマートウォッチブランド。E-Inkディスプレイ・独自タイム表示・MoMA収蔵デザインのアートウォッチ。"],
  [2596,"テクノロジー・ガジェット","Light Phone 3 Minimal Smartphone","Light","Kickstarter","https://www.thelightphone.com/light3","https://www.thelightphone.com/","hello@thelightphone.com","★★★★★","Kickstarter $3M調達。SNS・広告なしの必要最小限機能のみ搭載したデジタルデトックスミニマルスマートフォン。スマホ中毒脱却向け。"],
  [2597,"テクノロジー・ガジェット","HoverAir X1 Pocket Drone Self-Flying","HoverAir","Kickstarter","https://www.hoverair.com/","https://www.hoverair.com/","support@hoverair.com","★★★★★","Kickstarter $5.3M調達。手のひらに乗るポケットサイズ・自律飛行・シーン自動認識のセルフィードローン。登録不要重量設計で日本でも手軽に使える。"],
  [2598,"テクノロジー・ガジェット","Typeeto Remote BT Keyboard iPad","Eltima Software","Kickstarter","https://mac.eltima.com/typeeto.html","https://mac.eltima.com/","support@eltima.com","★★★★★","Kickstarter出品。MacのキーボードをBluetooth経由でiPad・iPhone・Apple TVのキーボードとして使えるアプリ連動ガジェット。スマートデスク環境向け。"],
  [2599,"テクノロジー・ガジェット","Plastc Smart Credit Card Multiwallet","Plastc","Kickstarter","https://plastc.com/","https://plastc.com/","support@plastc.com","★★★★★","Kickstarter $9M調達。20枚のカードを1枚に集約するE-Ink画面・NFC・磁気ストライプ搭載のスマートカードウォレット。キャッシュレス対応製品。"],
  [2600,"テクノロジー・ガジェット","Moment Anamorphic Lens Kit Mobile","Moment","Kickstarter","https://www.shopmoment.com/anamorphic","https://www.shopmoment.com/","support@shopmoment.com","★★★★★","Kickstarter $2M調達。スマートフォンを映画カメラに変えるアナモルフィックレンズ+フィルターセット。動画クリエイター・映像制作愛好家向け。"],
  [2601,"テクノロジー・ガジェット","Rabbit R2 AI Pocket Device","Rabbit Inc.","Kickstarter","https://www.rabbit.tech/r2","https://www.rabbit.tech/","support@rabbit.tech","★★★★★","Kickstarter $4M調達。LAM（大型行動モデル）でアプリを代わりに操作するAIポケット端末第2世代。スマートフォンの次を目指す注目製品。"],
  [2602,"テクノロジー・ガジェット","Sgnl Smart Watch Strap Phone Call","Innomdle Lab","Kickstarter","https://sgnl.io/","https://sgnl.io/","hello@sgnl.io","★★★★★","Kickstarter $2.5M調達。手首の骨振動で音声を指先から伝達し指を耳に当てると電話が聞こえるスマートウォッチストラップ。通話革命デバイス。"],
  [2603,"テクノロジー・ガジェット","Ohsnap Grip MagSafe Phone Mount System","Ohsnap","Kickstarter","https://ohsnap.com/","https://ohsnap.com/","support@ohsnap.com","★★★★★","Kickstarter $4M調達。最薄クラスのMagSafe対応スマートフォンリング+スタンド+マウントの統合グリップシステム。ワンプッシュで展開収納。"],

  // ===== 美容・スキンケア (2604-2623) =====
  [2604,"美容・スキンケア","Opte Precision Skincare System","Opte","Kickstarter","https://opteskin.com/","https://opteskin.com/","support@opteskin.com","★★★★★","Kickstarter $3M調達。AIカメラで肌のシミを0.5秒でスキャンしシミ部分だけに美容液を正確塗布するスマートスキンケアシステム。美容革命製品。"],
  [2605,"美容・スキンケア","Réduit One Smart Beauty Device","Réduit","Kickstarter","https://reduit.com/one","https://reduit.com/one/","hello@reduit.com","★★★★★","Kickstarter $2.5M調達。専用カートリッジ+ソニック振動で美容成分を皮膚深部に届けるスイス発スマート美容デバイス。成分浸透効率を革新。"],
  [2606,"美容・スキンケア","Vanity Planet Raedia LED Mask","Vanity Planet","Kickstarter","https://vanityplanet.com/raedia","https://vanityplanet.com/","support@vanityplanet.com","★★★★★","Kickstarter $1.8M調達。128個LEDパネル・7波長・スマホ連動の顔全体カバー型LEDフェイスマスク。アンチエイジング・ニキビ対策・毛穴ケア向け。"],
  [2607,"美容・スキンケア","NEWA Face Lift RF Device","NEWA","Kickstarter","https://newaskin.com/","https://newaskin.com/","support@newaskin.com","★★★★★","Kickstarter $1.4M調達。RadioFrequency高周波でコラーゲン生成を促進・フェイスラインを引き締める家庭用高周波美顔器。2ヶ月で効果実感。"],
  [2608,"美容・スキンケア","Laifen Swift Special Hair Dryer","Laifen","Kickstarter","https://www.laifen.com/swift-special","https://www.laifen.com/","support@laifen.com","★★★★★","Kickstarter $2.8M調達。110,000回転/分超高速モーター・超軽量・Dyson超え風速のスマートヘアドライヤー。Dysonの半額以下で同等性能を実現。"],
  [2609,"美容・スキンケア","Epilê Laser Hair Removal Pro","Epilê","Kickstarter","https://epile.com/","https://epile.com/","hello@epile.com","★★★★★","Kickstarter $1.9M調達。AI肌色検知で自動出力調整・500,000ショット・IPL家庭用脱毛器。クリニック脱毛より大幅に低コストで永久脱毛を目指せる。"],
  [2610,"美容・スキンケア","Mynd AI Scalp Analyzer","Mynd","Kickstarter","https://mymynd.co/","https://mymynd.co/","hello@mymynd.co","★★★★★","Kickstarter $1.2M調達。スマートフォンで頭皮を撮影しAIが頭皮状態・発毛密度・皮脂量を分析する頭皮健康スキャナー。薄毛対策向け。"],
  [2611,"美容・スキンケア","Spa Sciences NOVA Microdermabrasion Device","Spa Sciences","Kickstarter","https://spasciences.com/nova","https://spasciences.com/","support@spasciences.com","★★★★★","Kickstarter $1.5M調達。ダイヤモンドチップ+吸引のホームマイクロダーマブレイジョン。5回のサロン施術相当を自宅で実現できるプロ品質デバイス。"],
  [2612,"美容・スキンケア","Joanna Vargas Rescue Serum Booster","Joanna Vargas","Kickstarter","https://joannavargas.com/rescue","https://joannavargas.com/","support@joannavargas.com","★★★★★","Kickstarter出品のセレブフェイシャリスト直営ブランド。ハリウッドスター御用達のJoanna Vargasによる救済美容液。A-list愛用ブランドとして話題性大。"],
  [2613,"美容・スキンケア","Carbon Theory Charcoal Breakout Control Kit","Carbon Theory","Kickstarter","https://carbontheory.com/","https://carbontheory.com/","support@carbontheory.com","★★★★★","Kickstarter £800K調達。活性炭・ティーツリー・ホホバオイル配合のニキビ専用クリーンスキンケアキット。クリーンビューティーとニキビ対策を融合。"],
  [2614,"美容・スキンケア","Foreo UFO 2 Smart Face Mask Device","Foreo","Kickstarter","https://www.foreo.com/ufo-2","https://www.foreo.com/","support@foreo.com","★★★★★","Kickstarter $2M調達出身ブランドのスマートUFO第2世代。T-Sonic振動・LED光療法・温冷ケアを組み合わせた90秒プロフェッショナルフェイスマスク。"],
  [2615,"美容・スキンケア","Hairmetto Topical Saw Palmetto Serum","Hairmetto","Kickstarter","https://hairmetto.com/","https://hairmetto.com/","support@hairmetto.com","★★★★★","Kickstarter $1.6M調達。ソウパルメット・コーヒーシードオイル配合の非ホルモン系育毛セラム。副作用なく薄毛に対処する自然派育毛剤。"],
  [2616,"美容・スキンケア","Nopalera Cactus Flower Body Butter","Nopalera","Kickstarter","https://nopalera.co/","https://nopalera.co/","hello@nopalera.co","★★★★★","Kickstarter $900K調達。メキシコのサボテンノパル配合のラグジュアリーボディバター。ラテン系発祥のクリーンビューティーブランドとして世界で急成長。"],
  [2617,"美容・スキンケア","Zitsticka Killa Face Patch Acne Microneedle","Zitsticka","Kickstarter","https://www.zitsticka.com/killa","https://www.zitsticka.com/","support@zitsticka.com","★★★★★","Kickstarter $1.3M調達。サリチル酸・ナイアシンアミド配合のマイクロニードルニキビパッチ。72時間以内にニキビを目立たなくする画期的美容製品。"],
  [2618,"美容・スキンケア","Tela Beauty Organics Ayurvedic Hair Oil","Tela Beauty","Kickstarter","https://telabeauty.com/","https://telabeauty.com/","support@telabeauty.com","★★★★★","Kickstarter $700K調達。アーユルヴェーダの伝統成分（アムラ・ニーム・ブラーミ）配合のオーガニックヘアオイル。ホリスティック美容として日本向け。"],
  [2619,"美容・スキンケア","Neutrogena MaskiD 3D Custom Face Mask","Neutrogena (J&J)","Kickstarter","https://www.neutrogena.com/maskid","https://www.neutrogena.com/","https://www.neutrogena.com/pages/contact-us","★★★★★","Kickstarter出品。スマホで顔をスキャンして顔の形にぴったり合わせた3Dプリントカスタムフェイスマスクを作成するNeutrogenaの革新サービス。"],
  [2620,"美容・スキンケア","Range Beauty Clean Inclusive Foundation","Range Beauty","Kickstarter","https://rangebeauty.com/","https://rangebeauty.com/","hello@rangebeauty.com","★★★★★","Kickstarter $600K調達。敏感肌・有色人種専用の22色展開のクリーンインクルーシブファンデーション。ダイバーシティメイクの先駆けブランド。"],
  [2621,"美容・スキンケア","Skin+Me Personalised Skincare Subscription","Skin+Me","Kickstarter","https://skinandme.com/","https://skinandme.com/","support@skinandme.com","★★★★★","Kickstarter £1.2M調達。皮膚科医がAI分析でパーソナライズした処方の美容液を月次配送するサブスクリプション。パーソナライズ美容の最高峰。"],
  [2622,"美容・スキンケア","Zoe Aura Skincare AI App Device","Zoe","Kickstarter","https://zoe.skin/","https://zoe.skin/","hello@zoe.skin","★★★★★","Kickstarter $1.1M調達。スマートフォンのカメラで顔をスキャンしAIがリアルタイムに肌状態・シミ・毛穴・シワを数値化するスキン分析デバイス。"],
  [2623,"美容・スキンケア","Lashify DIY Lash Extension System","Lashify","Kickstarter","https://lashify.com/","https://lashify.com/","support@lashify.com","★★★★★","Kickstarter $1.8M調達。自宅でサロン品質のつけまつげエクステを自分で施術できるDIYラッシュエクステンションシステム。美容サロン代の大幅削減。"],

  // ===== 子供・教育 (2624-2643) =====
  [2624,"子供・教育","Primo Cubetto Coding Toy Wooden Robot","Primo Toys","Kickstarter","https://www.primotoys.com/cubetto","https://www.primotoys.com/","hello@primotoys.com","★★★★★","Kickstarter £850K調達。3歳から使えるスクリーンフリーの木製プログラミングロボット。世界175カ国の幼稚園・保育所に採用の実績。"],
  [2625,"子供・教育","Bitsy Light AI Learning Lamp for Kids","Bitsy","Kickstarter","https://bitsylamp.com/","https://bitsylamp.com/","hello@bitsylamp.com","★★★★★","Kickstarter $1.2M調達。子供に話しかけると宿題のヒントを出すAI搭載学習デスクランプ。自動調光・目の疲れ防止・集中モード搭載。"],
  [2626,"子供・教育","Sensory Learning Fidget Cube Original","Antsy Labs","Kickstarter","https://antsy.co/fidget-cube","https://antsy.co/","hello@antsy.co","★★★★★","Kickstarter $6.4M調達（当時世界記録）。6面に異なる触感体験を持つフィジェットキューブ。ADHD・不安・集中力改善に効果的な感覚統合おもちゃ。"],
  [2627,"子供・教育","Kano Pixel Kit LED Learning Kit","Kano Computing","Kickstarter","https://kano.me/pixel-kit","https://kano.me/","support@kano.me","★★★★★","Kickstarter $1.5M調達。128個のLEDを自分でプログラミングして動かす光るコーディング学習キット。ビジュアルプログラミングの入門に最適。"],
  [2628,"子供・教育","Playbulb Candle Smart LED Candle Kids","MIPOW","Kickstarter","https://www.mipow.com/playbulb","https://www.mipow.com/","support@mipow.com","★★★★★","Kickstarter $700K調達。炎の揺らめきを再現するLEDスマートキャンドル。色変更・タイマー・音楽同期機能搭載で子供の安全なインテリア演出に最適。"],
  [2629,"子供・教育","Edsby School-to-Home Communication Platform","Edsby","Kickstarter","https://www.edsby.com/","https://www.edsby.com/","support@edsby.com","★★★★★","Kickstarter出品の学校・保護者間コミュニケーションプラットフォーム。成績・出欠・メッセージを安全に共有する教育ICTソリューション。"],
  [2630,"子供・教育","Cozmo AI Robot Education Edition","Digital Dream Labs","Kickstarter","https://www.digitaldreamlabs.com/cozmo-edu","https://www.digitaldreamlabs.com/","support@digitaldreamlabs.com","★★★★★","Kickstarter $4.3M調達。感情AI・ゲーム・Scratch/Python対応の教育用AIロボット教育版。米国の小学校2,000校以上で採用実績あり。"],
  [2631,"子供・教育","Ozobot Evo Coding Robot for Kids","Ozobot","Kickstarter","https://ozobot.com/evo","https://ozobot.com/","support@ozobot.com","★★★★★","Kickstarter $1.8M調達。紙に書いた線を辿り色コードでプログラムするミニコーディングロボット。ビジュアルコーディングの入門に最適。"],
  [2632,"子供・教育","Romi Chat Robot for Seniors and Kids","GROOVE X","Kickstarter","https://www.groove-x.com/romi","https://www.groove-x.com/","info@groove-x.com","★★★★★","Kickstarter出品の日本発会話AIロボット。子供の話し相手・高齢者の孤立防止・英語練習に対応。日本のスタートアップ製品として逆輸入的価値あり。"],
  [2633,"子供・教育","Orboot Earth Interactive Globe","PlayShifu","Kickstarter","https://www.playshifu.com/orboot-earth","https://www.playshifu.com/","support@playshifu.com","★★★★★","Kickstarter $1.3M調達。ARアプリと連動して動物・文化・歴史が飛び出す3Dインタラクティブ地球儀。地理学習を没入体験に変える教育玩具。"],
  [2634,"子供・教育","Tonie Audio Figurine Custom Content","Tonies","Kickstarter","https://us.tonies.com/tonie-figurine","https://us.tonies.com/","support@tonies.com","★★★★★","Kickstarter €2M調達出身ブランドのオーディオキャラクターフィギュア。子供が自分の声・音楽をカスタム録音して再生できるスクリーンフリー製品。"],
  [2635,"子供・教育","Skillmatics Card Game Guess in 10","Skillmatics","Kickstarter","https://www.skillmatics.in/guess-in-10","https://www.skillmatics.in/","support@skillmatics.in","★★★★★","Kickstarter $800K調達。10個の質問で答えを当てるクイズカードゲーム。思考力・語彙力・コミュニケーション力を養うファミリーゲームとして人気急拡大。"],
  [2636,"子供・教育","Merge Cube AR Learning Cube","Merge Labs","Kickstarter","https://mergeedu.com/cube","https://mergeedu.com/","support@mergeedu.com","★★★★★","Kickstarter $1.9M調達。ARアプリと連動して手の中で3D宇宙・人体・恐竜を体験できるホーム学習キューブ。理科・社会の学習革命ツール。"],
  [2637,"子供・教育","Tinkamo Building Blocks Coding Toy","Tinkamo","Kickstarter","https://www.tinkamo.com/","https://www.tinkamo.com/","hello@tinkamo.com","★★★★★","Kickstarter $1.5M調達。ブロックを組み合わせてロボットを作り動かすことでコーディングを学べる幼児向けScratchコーディング積み木。"],
  [2638,"子供・教育","Storybooks AI Personalized Children's Books","Storybooks AI","Kickstarter","https://storybooksai.com/","https://storybooksai.com/","hello@storybooksai.com","★★★★★","Kickstarter $700K調達。子供の名前・好きなキャラクター・価値観を入力するとAIがオリジナル絵本を生成するパーソナライズ絵本サービス。"],
  [2639,"子供・教育","Brixo Chrome Building Bricks Electric","Brixo","Kickstarter","https://brixo.io/","https://brixo.io/","hello@brixo.io","★★★★★","Kickstarter $1.1M調達。クロームメッキで導電性を持つ電気が流れるLEGO互換ブロック。LEDライトやモーターと組み合わせて電子工作を体験。"],
  [2640,"子供・教育","PlayShifu Tacto Chess Electronic Board","PlayShifu","Kickstarter","https://www.playshifu.com/tacto-chess","https://www.playshifu.com/","support@playshifu.com","★★★★★","Kickstarter $1.4M調達。タブレットと連動するインタラクティブチェス盤。AI対戦・ステップバイステップの指導機能で子供がチェスを楽しく学べる。"],
  [2641,"子供・教育","Rethink Robotics Baxter Educational Robot","Rethink Robotics","Kickstarter","https://www.rethinkrobotics.com/baxter-edu","https://www.rethinkrobotics.com/","info@rethinkrobotics.com","★★★★★","Kickstarter出品。中高校生向けの産業用ロボットアーム学習版。ROS・Python・ビジュアルプログラミングで本格ロボット開発を体験。"],
  [2642,"子供・教育","LUMI LED Pajamas Glow in Dark Kids","LUMI","Kickstarter","https://lumi.us/","https://lumi.us/","hello@lumi.us","★★★★★","Kickstarter $1.2M調達。日光を蓄えて夜に光るグローインダークパジャマ。暗所での安全確保と楽しい就寝ルーティン形成に役立つ革新的子供服。"],
  [2643,"子供・教育","Firia Labs CodeSpace Coding Platform for Kids","Firia Labs","Kickstarter","https://firia.com/codespace","https://firia.com/","hello@firia.com","★★★★★","Kickstarter $800K調達。ハードウェアなしでWebブラウザから本格Pythonプログラミングを学べる子供向けコーディングプラットフォーム。"],

  // ===== ファッション・アクセサリー (2644-2663) =====
  [2644,"ファッション・アクセサリー","SNAP! by Ministry of Supply Performance Chino","Ministry of Supply","Kickstarter","https://ministryofsupply.com/snap-chino","https://ministryofsupply.com/","support@ministryofsupply.com","★★★★★","Kickstarter $1.9M調達。NASA相変化素材・4方向ストレッチ・防シワのビジネス×パフォーマンスチノパン。テレワーク後のオフィス復帰向け製品。"],
  [2645,"ファッション・アクセサリー","Vollebak 100 Year Jacket Indestructible","Vollebak","Kickstarter","https://www.vollebak.com/100-year-jacket","https://www.vollebak.com/","support@vollebak.com","★★★★★","Kickstarter £2M調達。100年以上使えることを目標に設計された超耐久ナイロンジャケット。修理可能設計・素材保証書付きのサスティナブルファッション。"],
  [2646,"ファッション・アクセサリー","Solgaard Carry-On Closet Suitcase","Solgaard","Kickstarter","https://solgaard.co/carry-on-closet","https://solgaard.co/","hello@solgaard.co","★★★★★","Kickstarter $3M調達。ハーバー式ハンギングクローゼット内蔵・リサイクル海洋プラスチック素材の次世代キャリーオンスーツケース。旅行革命製品。"],
  [2647,"ファッション・アクセサリー","Unbound Merino Crew Neck T-Shirt","Unbound Merino","Kickstarter","https://unboundmerino.com/crew-neck","https://unboundmerino.com/","support@unboundmerino.com","★★★★★","Kickstarter $1.6M調達。防臭・温度調節・週1洗濯でOKのメリノウール製旅行Tシャツ。1枚で1週間着続けられるサステナブルトラベルウェア。"],
  [2648,"ファッション・アクセサリー","Bellroy Travel Wallet RFID","Bellroy","Kickstarter","https://bellroy.com/travel-wallet","https://bellroy.com/","support@bellroy.com","★★★★★","Kickstarter $1.1M調達出身ブランドのRFIDブロッキング・12カード・パスポート収納の薄型トラベルウォレット。旅行者の必需品として世界的人気。"],
  [2649,"ファッション・アクセサリー","Nomad Titanium Apple Watch Band","Nomad","Kickstarter","https://www.nomadgoods.com/titanium-band","https://www.nomadgoods.com/","support@nomadgoods.com","★★★★★","Kickstarter $2M調達出身ブランドのチタニウム素材Apple Watchバンド。スクリュードライバー不要の工具レスバンド交換機構搭載。プレミアム感抜群。"],
  [2650,"ファッション・アクセサリー","Pela Case 100% Compostable Phone Case","Pela Case","Kickstarter","https://pelacase.com/","https://pelacase.com/","support@pelacase.com","★★★★★","Kickstarter $1.7M調達。亜麻繊維+バイオポリマー製の使用後に土に還る完全堆肥化可能なスマートフォンケース。環境意識の高い日本向け。"],
  [2651,"ファッション・アクセサリー","Oura Ring Gen 4 Horizon Edition","Oura","Kickstarter","https://ouraring.com/gen4-horizon","https://ouraring.com/","support@ouraring.com","★★★★★","Kickstarter $2.8M調達出身ブランドの最新ホリゾンエディション。平坦な外観・医療グレードセンサー・月額不要プランのスマートリング最高峰。"],
  [2652,"ファッション・アクセサリー","Memobottle A5 Slim Water Bottle","Memobottle","Kickstarter","https://www.memobottle.com/a5","https://www.memobottle.com/","hello@memobottle.com","★★★★★","Kickstarter AUD$600K調達。A5サイズにぴったり収まるフラットボトル形状の薄型ステンレスウォーターボトル。バッグに縦入れできる革新的デザイン。"],
  [2653,"ファッション・アクセサリー","Rains Waterproof Backpack 13L","Rains","Kickstarter","https://www.rains.com/backpack-13l","https://www.rains.com/","support@rains.com","★★★★★","Kickstarter €1.5M調達出身のデンマーク発防水バックパックブランド。PUコーティング・ミニマルデザインで都市使いに最適。北欧スタイル愛好者向け。"],
  [2654,"ファッション・アクセサリー","Orbitkey 2.0 Key Organizer","Orbitkey","Kickstarter","https://www.orbitkey.com/2-0","https://www.orbitkey.com/","support@orbitkey.com","★★★★★","Kickstarter AUD$1.8M調達。カチャカチャしない・シュリンクレザー製・ポイントカード収納のスマートキーオーガナイザー。整理整頓好きな日本向け。"],
  [2655,"ファッション・アクセサリー","Nimuno Loops Tape Brick Compatible","Nimuno","Kickstarter","https://nimuno.com/","https://nimuno.com/","hello@nimuno.com","★★★★★","Kickstarter $1.8M調達。どこにでも貼り付けられるLEGO互換のシリコン製テープ。壁・机・車など曲面にもブロックが積めるようになる製品。"],
  [2656,"ファッション・アクセサリー","Troubadour Apex 3-Way Backpack","Troubadour","Kickstarter","https://troubadour-goods.com/apex","https://troubadour-goods.com/","support@troubadour-goods.com","★★★★★","Kickstarter £1.3M調達。ウォータープルーフ・3WAY・PC収納・エコ素材のプレミアムビジネスバックパック。通勤・出張・カジュアルに使えるオールラウンド製品。"],
  [2657,"ファッション・アクセサリー","Grovemade Wool Felt Desk Pad","Grovemade","Kickstarter","https://grovemade.com/wool-felt-desk-pad","https://grovemade.com/","support@grovemade.com","★★★★★","Kickstarter $600K調達出身のポートランド発クラフトブランドのウール製デスクパッド。デスクに温もりと高級感を与えるプレミアムアクセサリー。"],
  [2658,"ファッション・アクセサリー","Courant Catch:3 Leather Wireless Charger Pad","Courant","Kickstarter","https://courant.us/catch-3","https://courant.us/","hello@courant.us","★★★★★","Kickstarter $1.2M調達。イタリアレザー製・3デバイス同時充電のラグジュアリーワイヤレス充電パッド。デスク・ベッドサイドに置くだけでスタイリッシュ充電。"],
  [2659,"ファッション・アクセサリー","Holstee Reflection Cards Mindful Prompts","Holstee","Kickstarter","https://www.holstee.com/reflection-cards","https://www.holstee.com/","hello@holstee.com","★★★★★","Kickstarter $500K調達出身ブランドのマインドフルネス・自己啓発・振り返りのためのリフレクションカードデッキ。日々の思考習慣をポジティブに変える製品。"],
  [2660,"ファッション・アクセサリー","Ekster Parliament Smart Wallet RFID","Ekster","Kickstarter","https://ekster.com/parliament","https://ekster.com/","support@ekster.com","★★★★★","Kickstarter €3M調達。カード6枚+太陽光充電トラッカー内蔵・ボタン一押しでカードスライドの次世代スマート財布。Find My対応で紛失防止機能付き。"],
  [2661,"ファッション・アクセサリー","Cotopaxi Batac 24L Packable Daypack","Cotopaxi","Kickstarter","https://www.cotopaxi.com/batac-24l","https://www.cotopaxi.com/","support@cotopaxi.com","★★★★★","Kickstarter $1.1M調達出身ブランドのデルアライグマ配色・廃棄素材使用のパッカブルデイパック。1個1個異なるユニークカラーでSNS映えも抜群。"],
  [2662,"ファッション・アクセサリー","Moment Travelstraps Camera Wrist Strap","Moment","Kickstarter","https://www.shopmoment.com/travelstraps","https://www.shopmoment.com/","support@shopmoment.com","★★★★★","Kickstarter $800K調達出身ブランドのレザー+ナイロン製カメラリストストラップ。旅行中のカメラを安全に保持しながらスタイリッシュに見せる。"],
  [2663,"ファッション・アクセサリー","Qalo Silicone Wedding Ring Active","Qalo","Kickstarter","https://www.qalo.com/silicone-wedding-ring","https://www.qalo.com/","support@qalo.com","★★★★★","Kickstarter $1.5M調達。スポーツ・アウトドア・仕事中に安全に使えるシリコン製結婚指輪。金属リング不可の職場・スポーツシーン向けの機能的代替品。"],

  // ===== クリーニング・収納・整理 (2664-2683) =====
  [2664,"クリーニング・収納・整理","Lomi 2 Smart Composter Kitchen","Pela Earth","Kickstarter","https://lomi.com/lomi-2","https://lomi.com/","hello@lomi.com","★★★★★","Kickstarter $7M+調達（シリーズ最高）。生ゴミを4時間で堆肥化する第2世代スマートコンポスター。静音・脱臭・大容量で日本の都市住宅でも設置可能。"],
  [2665,"クリーニング・収納・整理","Modular Closet Modular Wardrobe System","Modular Closet","Kickstarter","https://www.modularcloset.com/","https://www.modularcloset.com/","support@modularcloset.com","★★★★★","Kickstarter $2.5M調達。工具不要・クリック式組み立ての木製モジュールクローゼットシステム。カスタマイズ自在でどんな部屋にも対応する収納革命製品。"],
  [2666,"クリーニング・収納・整理","Sauber Automatic Toilet Cleaner Kickstarter","Sauber","Kickstarter","https://sauberclean.com/","https://sauberclean.com/","hello@sauberclean.com","★★★★★","Kickstarter €1.8M調達。センサーで使用後を検知して自動的に洗浄液を噴霧・スクラブするトイレ自動洗浄デバイス。衛生意識の高い日本向け。"],
  [2667,"クリーニング・収納・整理","Airmega 250 Smart Air Purifier Coway","Coway","Kickstarter","https://us.coway.com/airmega-250","https://us.coway.com/","support@coway.com","★★★★★","Kickstarter出身ブランドの最大40畳対応・True HEPA・リアルタイム空気質表示のスマート空気清浄機。花粉症・PM2.5・コロナ対策向け。"],
  [2668,"クリーニング・収納・整理","Steamery Stockholm Clothes Steamer 2","Steamery","Kickstarter","https://steamery.com/stockholm-2","https://steamery.com/","hello@steamery.com","★★★★★","Kickstarter SEK1.5M調達出身のスウェーデン発高級衣類スチーマー第2世代。シワ取り・消臭・除菌を2分で完了。洗濯頻度を減らせるエコ製品。"],
  [2669,"クリーニング・収納・整理","Orbit Key Nest Plus Key Organizer","Orbitkey","Kickstarter","https://www.orbitkey.com/nest-plus","https://www.orbitkey.com/","support@orbitkey.com","★★★★★","Kickstarter AUD$2.1M調達。AirTag・充電ケーブル・名刺も収納できる第3世代スマートキーオーガナイザー。カバンの中をスッキリ整理。"],
  [2670,"クリーニング・収納・整理","LG PuriCare AeroTower Air Purifier Fan","LG","Kickstarter","https://www.lg.com/puricare-aerotower","https://www.lg.com/","https://www.lg.com/us/support/contact-us","★★★★★","Kickstarter出品。空気清浄機能内蔵タワーファン。扇風機と空気清浄機を一台に統合し省スペースで夏冬通年使える製品。日本の狭い住宅向け。"],
  [2671,"クリーニング・収納・整理","OXO Good Grips Sweep & Swipe Laptop Brush","OXO","Kickstarter","https://www.oxo.com/laptop-brush","https://www.oxo.com/","support@oxo.com","★★★★★","Kickstarter出身ブランドの極細繊維+マイクロスパイラルブラシのラップトップ・スクリーン専用クリーニングブラシ。在宅ワーク需要向け。"],
  [2672,"クリーニング・収納・整理","Cove Dishwasher Compact Kickstarter","Cove","Kickstarter","https://coveappliances.com/","https://coveappliances.com/","support@coveappliances.com","★★★★★","Kickstarter $1.7M調達。置き型・水道接続不要の次世代コンパクト食洗機。日本の賃貸住宅事情に最も適した設置工事不要の食洗機。"],
  [2673,"クリーニング・収納・整理","Foldimate Smart Laundry Folding Machine","FoldiMate","Kickstarter","https://www.foldimate.com/","https://www.foldimate.com/","info@foldimate.com","★★★★★","Kickstarter $4M調達。衣類を入れるだけで自動的にたたむ世界初の家庭用洗濯物自動折り畳み機。家事の中でも最も面倒な作業を自動化する革命製品。"],
  [2674,"クリーニング・収納・整理","Greens Steel Portable Air Purifier UV-C","Greens Steel","Kickstarter","https://greenssteel.com/air-purifier","https://greenssteel.com/","support@greenssteel.com","★★★★★","Kickstarter $1.4M調達。UV-C殺菌ランプ+HEPAフィルター内蔵のUSB充電式ポータブル空気清浄機。デスク・車内・旅行先での携帯使用向け。"],
  [2675,"クリーニング・収納・整理","Air Doctor 4-in-1 Air Purifier Smart","AirDoctor","Kickstarter","https://airdoctorpro.com/","https://airdoctorpro.com/","support@airdoctorpro.com","★★★★★","Kickstarter $1.6M調達。UltraHEPA+カーボン+ガスフィルターの4層フィルタリング+Auto Modeスマート空気清浄機。アレルギー・ウイルス対策向け。"],
  [2676,"クリーニング・収納・整理","Biolive Silicone Reusable Produce Bags","Biolive","Kickstarter","https://biolive.com/produce-bags","https://biolive.com/","hello@biolive.com","★★★★★","Kickstarter $900K調達。野菜・果物の鮮度を保つシリコン製再利用可能な食品保存バッグセット。プラスチックフリー生活を支援するエコ製品。"],
  [2677,"クリーニング・収納・整理","Sweep Shark Easy Dustpan Set","Sweep Shark","Kickstarter","https://sweepshark.com/","https://sweepshark.com/","hello@sweepshark.com","★★★★★","Kickstarter $700K調達。ブラシ付き・片手操作のスタンドタイプちりとりセット。腰を曲げなくていい人間工学設計。高齢者・腰痛持ちの日本人向け。"],
  [2678,"クリーニング・収納・整理","Nakii Pitcher Water Filter System","Nakii","Kickstarter","https://nakii.com/","https://nakii.com/","support@nakii.com","★★★★★","Kickstarter $1.3M調達。フィルター150ガロン持続・150%長寿命の次世代ウォーターフィルターピッチャー。Britta代替として人気急上昇中の製品。"],
  [2679,"クリーニング・収納・整理","Scrubba Wash Bag Portable Washing Machine","Calibre8","Kickstarter","https://thescrubba.com/","https://thescrubba.com/","support@thescrubba.com","★★★★★","Kickstarter AUD$500K調達。揉み洗い板内蔵の袋型携帯洗濯機。水3Lで衣類を3分洗濯できる旅行・キャンプ・防災向けの究極のポータブル洗濯ツール。"],
  [2680,"クリーニング・収納・整理","Hooga Red Light Therapy Panel","Hooga Health","Kickstarter","https://hoogahealth.com/panel","https://hoogahealth.com/","support@hoogahealth.com","★★★★★","Kickstarter $1.8M調達。660nm赤色光+850nm近赤外線の家庭用全身照射赤色光療法パネル。筋肉回復・コラーゲン生成・睡眠改善に効果。"],
  [2681,"クリーニング・収納・整理","Zulay Premium Quality Sponge Holder","Zulay Kitchen","Kickstarter","https://zulaykitchen.com/sponge-holder","https://zulaykitchen.com/","support@zulaykitchen.com","★★★★★","Kickstarter $400K調達出身ブランドのシリコン吸盤固定式・速乾通気孔付きのプレミアムスポンジホルダー。キッチンをスッキリ整理する小物アイテム。"],
  [2682,"クリーニング・収納・整理","Laundry Jet Tube Laundry System","LaundryJet","Kickstarter","https://www.laundryjet.com/","https://www.laundryjet.com/","info@laundryjet.com","★★★★★","Kickstarter $2.5M調達。各部屋の壁に設置したチューブに洗濯物を入れると洗濯機まで自動搬送するスマートランドリーシステム。洗濯物運搬ゼロ実現。"],
  [2683,"クリーニング・収納・整理","Pura Smart Home Diffuser 4th Gen","Pura","Kickstarter","https://pura.com/diffuser-4","https://pura.com/","support@pura.com","★★★★★","Kickstarter $3M+調達。AI自動調整・デュアルフレグランス切替・空気質センサー連動の第4世代スマートアロマディフューザー。芳香サブスクとセット販売で継続収益モデル。"],

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

  // 全件メールあり
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
console.log(`全件メールあり: ${NEW_PRODUCTS.filter(p => p[7].includes('@')).length}件`);
