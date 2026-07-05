/**
 * add_200_items_9.js — No.2284〜2483 (200件追加)
 * node add_200_items_9.js
 */

const XLSX = require('xlsx');
const path = require('path');

const INPUT_FILE  = path.join(__dirname, '海外便利グッズリスト_日本未上陸2305件_評価付.xlsx');
const OUTPUT_FILE = path.join(__dirname, '海外便利グッズリスト_日本未上陸2505件_評価付.xlsx');
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

  // ===== キッチン・調理器具 (2284-2303) =====
  [2284,"キッチン・調理器具","Mellow Sous Vide Connected Cooker","Mellow","Mellow公式","https://www.cookmellow.com/","https://www.cookmellow.com/","hello@cookmellow.com","★★★★★","冷蔵→低温調理を一台で管理するスマート低温調理器。外出先からアプリで調理開始時刻を設定し帰宅時に料理が完成。共働き家庭向け。"],
  [2285,"キッチン・調理器具","Keurig K-Supreme Plus Smart Coffee Maker","Keurig","Amazon US","https://www.keurig.com/K-Supreme-Plus-SMART/","https://www.keurig.com/","support@keurig.com","★★★★☆","Wi-Fi接続・スマホ予約・温度カスタマイズのシングルサーブコーヒーメーカー。米国市場No.1のKeurigが日本に参入すれば即ヒット必至。"],
  [2286,"キッチン・調理器具","Nespresso Vertuo Pop Coffee Machine","Nespresso","Nespresso公式","https://www.nespresso.com/vertuo-pop","https://www.nespresso.com/","support@nespresso.com","★★★★★","バーコードスキャンで最適抽出を自動設定するヴァーチュオシステムのエントリーモデル。コンパクトで日本の狭いキッチンに最適。"],
  [2287,"キッチン・調理器具","GreenPan Rio Healthy Ceramic Pan Set","GreenPan","Amazon US","https://www.greenpan.us/","https://www.greenpan.us/","support@greenpan.us","★★★★★","PFAS・PFOA完全フリーのセラミックノンスティッククックウェアセット。健康志向のクリーンコーキング需要が高まる日本向け。"],
  [2288,"キッチン・調理器具","Made In Blue Carbon Steel Wok","Made In Cookware","Made In公式","https://madeincookware.com/","https://madeincookware.com/","support@madeincookware.com","★★★★★","プロシェフ愛用の青鋼鉄製本格中華鍋。使い込むほど育つ鉄鍋文化と日本の料理文化の親和性が高い。料理好き向けのプレミアム鍋。"],
  [2289,"キッチン・調理器具","Our Place Always Pan 2.0","Our Place","Our Place公式","https://fromourplace.com/","https://fromourplace.com/","support@fromourplace.com","★★★★★","フライパン・ソテーパン・蒸し器・鍋など8機能を一台で実現するオールインワン万能フライパン。収納スペースが限られる日本の住宅に最適。"],
  [2290,"キッチン・調理器具","Caraway Cookware Set Non-Toxic","Caraway","Caraway公式","https://www.carawayhome.com/","https://www.carawayhome.com/","support@carawayhome.com","★★★★★","セラミックコーティング・PTFE/PFOA/PFAS不使用のクリーンクックウェアセット。インスタグラム映えするパステルカラー展開。"],
  [2291,"キッチン・調理器具","Balmuda The Range Microwave Oven","Balmuda","Balmuda公式","https://www.balmuda.com/range","https://www.balmuda.com/","https://www.balmuda.com/pages/contact","★★★★★","日本発のスタイリッシュなオーブンレンジ。海外市場での日本ブランドの人気を示す逆輸入事例として掲載。デザイン家電の最高峰。"],
  [2292,"キッチン・調理器具","CookUnity Chef-Cooked Meal Delivery","CookUnity","CookUnity公式","https://www.cookunity.com/","https://www.cookunity.com/","support@cookunity.com","★★★★★","プロシェフが調理した冷凍ミールをサブスクで届けるミールデリバリーサービス。調理時間ゼロの生活を実現。日本の時短需要向け。"],
  [2293,"キッチン・調理器具","Staub Cast Iron Round Cocotte 5.5qt","Staub","Amazon US","https://www.staub-online.com/","https://www.staub-online.com/","support@staub.fr","★★★★★","フランス製鋳鉄ホーロー鍋の最高峰。煮込み料理・パン焼きに最適で一生使える投資型調理器具。料理好きな日本人に根強い人気がある。"],
  [2294,"キッチン・調理器具","Lodge 12-inch Cast Iron Skillet","Lodge Cast Iron","Amazon US","https://www.lodgecastiron.com/","https://www.lodgecastiron.com/","support@lodgecastiron.com","★★★★★","米国で100年以上の歴史を持つ鋳鉄スキレットブランド。キャンプ料理・ステーキ焼きに最適。アウトドア料理需要が高まる日本向け。"],
  [2295,"キッチン・調理器具","Dash Mini Waffle Maker","Dash","Amazon US","https://www.bydash.com/","https://www.bydash.com/","support@bydash.com","★★★☆☆","ハート・アニマル・スタースなど形が選べる超小型ミニワッフルメーカー。SNS映えする朝食作りで特にZ世代・子育て家庭向けに人気。"],
  [2296,"キッチン・調理器具","Vitamix A3500 Ascent Series Blender","Vitamix","Vitamix公式","https://www.vitamix.com/a3500","https://www.vitamix.com/","support@vitamix.com","★★★★★","スマートスケール連動・自己洗浄・プログラム5モード搭載のVitamixフラッグシップモデル。プロ・健康志向者向けの究極のブレンダー。"],
  [2297,"キッチン・調理器具","Cuisinart 14-Cup Food Processor","Cuisinart","Amazon US","https://www.cuisinart.com/food-processors","https://www.cuisinart.com/","support@cuisinart.com","★★★★★","14カップ大容量・パルス機能・多彩なアタッチメントのフードプロセッサー。米国キッチン家電の定番ブランドとして日本市場に参入余地大。"],
  [2298,"キッチン・調理器具","Microplane Premium Classic Zester Grater","Microplane","Amazon US","https://us.microplane.com/","https://us.microplane.com/","support@microplane.com","★★★★★","世界のシェフが愛用するプレミアムスチール製ゼスターグレーター。チーズ・柑橘・スパイスの細かすりおろしを美しく実現。"],
  [2299,"キッチン・調理器具","Zwilling J.A. Henckels 8-inch Chef's Knife","Zwilling J.A. Henckels","Amazon US","https://www.zwilling.com/us/","https://www.zwilling.com/","support@zwilling.com","★★★★★","ドイツSolingen製・シグマフォージ技術・氷焼き入れ鋼の最高峰シェフズナイフ。刃物文化が深い日本市場でも高い評価を得る製品。"],
  [2300,"キッチン・調理器具","Breville Barista Express Impress Espresso","Breville","Amazon US","https://www.breville.com/breville-barista-express-impress","https://www.breville.com/","support@breville.com","★★★★★","グラインダー内蔵・自動タンピング・圧力ゲージ搭載のオールインワンエスプレッソマシン。カフェ品質のエスプレッソを自宅で完全再現。"],
  [2301,"キッチン・調理器具","Anova Precision Oven Steam Combi","Anova Culinary","Anova公式","https://anovaculinary.com/anova-precision-oven","https://anovaculinary.com/","support@anovaculinary.com","★★★★★","スチーム・コンベクション・低温調理を一台で実現するスマートコンビオーブン。Wi-Fi連動でプロシェフのレシピを自動再現。"],
  [2302,"キッチン・調理器具","Bialetti Moka Express Stovetop Espresso Maker","Bialetti","Amazon US","https://www.bialetti.com/moka-express","https://www.bialetti.com/","support@bialetti.com","★★★★★","イタリア家庭の象徴的コーヒーメーカー。90年以上の歴史を持つモカポット。コーヒー愛好家向けの定番・入門製品として日本市場向け。"],
  [2303,"キッチン・調理器具","Zojirushi Stainless Vacuum Insulated Lunch Jar","Zojirushi","Amazon US","https://www.zojirushi.com/app/product/sweke","https://www.zojirushi.com/","https://www.zojirushi.com/app/contact","★★★★★","象印のステンレス真空二重構造ランチジャー。弁当・スープを長時間保温する日本製品を海外へ紹介する逆輸入事例として掲載価値大。"],

  // ===== スマートホーム・インテリア・照明 (2304-2323) =====
  [2304,"スマートホーム・インテリア・照明","Ecobee SmartThermostat Premium","Ecobee","Ecobee公式","https://www.ecobee.com/smart-thermostat-premium","https://www.ecobee.com/","support@ecobee.com","★★★★★","空気質センサー・煙/CO検知・スマートスピーカー内蔵のオールインワンスマートサーモスタット。省エネと安全を両立するプレミアム製品。"],
  [2305,"スマートホーム・インテリア・照明","Sonos Era 300 Spatial Audio Speaker","Sonos","Sonos公式","https://www.sonos.com/era-300","https://www.sonos.com/","support@sonos.com","★★★★★","Dolby Atmos・6スピーカー配列・空間オーディオ対応のプレミアムスマートスピーカー。音楽を立体的に再現する次世代ホームオーディオ。"],
  [2306,"スマートホーム・インテリア・照明","Bose Smart Soundbar 900","Bose","Bose公式","https://www.bose.com/soundbar-900","https://www.bose.com/","support@bose.com","★★★★★","Dolby Atmos・TrueSpace技術・ADAPTiQ自動音場補正搭載のプレミアムスマートサウンドバー。ホームシアターの音響体験を圧倒的に向上。"],
  [2307,"スマートホーム・インテリア・照明","Nest Cam Outdoor with Floodlight","Google Nest","Google公式","https://store.google.com/nest-cam-floodlight","https://store.google.com/","https://support.google.com/googlenest","★★★★★","2400ルーメン・4K・AI人物/動物/車検知のスマートフラッドライトカメラ。Googleの高度なAI分析で精度の高い侵入検知を実現。"],
  [2308,"スマートホーム・インテリア・照明","Yale Assure Lock 2 Smart Lock","Yale","Yale公式","https://www.yalehome.com/assure-lock-2","https://www.yalehome.com/","support@yale.com","★★★★★","Matter対応・指紋認証・Apple HomeKey・Alexa連携のスマートデッドボルト。世界最古の錠前ブランドの最新スマートロック。"],
  [2309,"スマートホーム・インテリア・照明","Levoit Core 600S Smart Air Purifier","Levoit","Amazon US","https://levoit.com/core600s","https://levoit.com/","support@levoit.com","★★★★★","635㎡対応・True HEPA・スマホ操作・自動モードの高性能空気清浄機。花粉症・アレルギー・PM2.5対策として日本市場で高需要。"],
  [2310,"スマートホーム・インテリア・照明","Dyson Pure Hot+Cool HP07","Dyson","Dyson公式","https://www.dyson.com/hp07","https://www.dyson.com/","support@dyson.com","★★★★★","空気清浄・暖房・冷風の3機能搭載のオールシーズン型スマート空気清浄機。DysonブランドはAmazon Japanで極めて高い人気。"],
  [2311,"スマートホーム・インテリア・照明","Furrion Aurora 4K Outdoor TV","Furrion","Furrion公式","https://www.furrion.com/aurora","https://www.furrion.com/","support@furrion.com","★★★★★","4K・HDR・IP55防水・耐紫外線コーティングのアウトドア設置専用スマートTV。ガーデニング・テラスライフ需要が高まる日本向け。"],
  [2312,"スマートホーム・インテリア・照明","Lutron Aurora Smart Dimmer Switch","Lutron","Amazon US","https://www.casetawireless.com/aurora","https://www.casetawireless.com/","support@lutron.com","★★★★★","既存のスイッチに被せるだけで設置できるスマート調光器。工事不要・賃貸対応のスマート照明コントローラとして日本市場に最適。"],
  [2313,"スマートホーム・インテリア・照明","Arlo Essential Indoor Camera 2K","Arlo","Arlo公式","https://www.arlo.com/essential-indoor-camera","https://www.arlo.com/","support@arlo.com","★★★★★","2K解像度・AI動体検知・プライバシーシャッター搭載の室内スマートカメラ。ペット・赤ちゃん・家の安全監視に最適な多用途カメラ。"],
  [2314,"スマートホーム・インテリア・照明","Philips Hue Play HDMI Sync Box","Philips Hue","Philips Hue公式","https://www.philips-hue.com/sync-box","https://www.philips-hue.com/","support@philips-hue.com","★★★★★","映像・ゲーム・音楽に同期してPhilips Hue照明を自動変色させるHDMIシンクボックス。没入型エンタメ体験を実現する革新的アクセサリー。"],
  [2315,"スマートホーム・インテリア・照明","Eve Door & Window Sensor Matter","Eve Systems","Eve公式","https://www.evehome.com/door-window","https://www.evehome.com/","support@evehome.com","★★★★☆","Matter対応・電池駆動・設置簡単のスマートドア/窓センサー。開閉を検知してオートメーションをトリガー。スマートホームの基本パーツ。"],
  [2316,"スマートホーム・インテリア・照明","Hue Festoon Outdoor String Lights","Philips Hue","Philips Hue公式","https://www.philips-hue.com/festoon","https://www.philips-hue.com/","support@philips-hue.com","★★★★★","Bluetooth・7mのスマート屋外LEDガーランドライト。テラス・バルコニー・庭のおしゃれな装飾に最適。アウトドアリビングをスマート化。"],
  [2317,"スマートホーム・インテリア・照明","Govee Neon Flex Light Bar","Govee","Govee公式","https://www.govee.com/","https://www.govee.com/","marketing@govee.com","★★★★★","曲げ自在・DIYカット可能・RGBIC個別点灯のネオンフレックスライトバー。ゲーミングルームやTikTok用背景として世界的人気。"],
  [2318,"スマートホーム・インテリア・照明","WeMo Stage Smart Dimmer Switch","WeMo by Belkin","Amazon US","https://www.belkin.com/wemo-stage-smart-dimmer","https://www.belkin.com/","support@belkin.com","★★★★☆","Matter対応・工事簡単のスマート調光スイッチ。Alexa・Google Home・Apple HomeKit全対応で選択肢が多い汎用スマートスイッチ。"],
  [2319,"スマートホーム・インテリア・照明","Ring Indoor Cam 2nd Gen","Ring","Amazon US","https://ring.com/indoor-cam","https://ring.com/","support@ring.com","★★★★☆","1080p・プライバシーシャッター・カラーナイトビジョンの低価格スマート室内カメラ。Amazon Alexa連携で音声操作が可能。"],
  [2320,"スマートホーム・インテリア・照明","Shark IQ AV1010WD Robot Vacuum Cleaner","Shark","Amazon US","https://www.sharkclean.com/","https://www.sharkclean.com/","support@shark.com","★★★★★","自動ゴミ収集・マッピング・Wi-Fi接続のロボット掃除機。iRobot Roombaに対するコスパ最高の競合製品として人気急上昇中。"],
  [2321,"スマートホーム・インテリア・照明","Amazon Echo Show 15 Smart Display","Amazon","Amazon公式","https://www.amazon.com/echo-show-15","https://www.amazon.com/","https://www.amazon.com/contact-us","★★★★★","15.6インチ大画面・家族スケジュール・ウィジェット・Alexa内蔵のスマートディスプレイ。キッチン・リビングの情報ハブとして活用。"],
  [2322,"スマートホーム・インテリア・照明","Cync Smart LED Bulb GE Lighting","GE Lighting","Amazon US","https://cync.com/","https://cync.com/","support@cync.com","★★★★☆","ハブ不要・Bluetooth+Wi-Fi・1600万色のスマート電球。GEの新ブランドCyncでスマートホーム入門向けにお手頃価格を実現。"],
  [2323,"スマートホーム・インテリア・照明","Bosch Smart Home Flex Universal Switch II","Bosch Smart Home","Bosch公式","https://www.bosch-smarthome.com/","https://www.bosch-smarthome.com/","support@bosch-smarthome.com","★★★★☆","壁・家具・どこにでも取り付けられるワイヤレスフレキシブルスマートスイッチ。配線不要でスマートホームのコントロールポイントを増設。"],

  // ===== ウェアラブル・ヘルス・フィットネス (2324-2343) =====
  [2324,"ウェアラブル・ヘルス・フィットネス","Suunto Wing Open-Ear Bone Conduction Sport","Suunto","Suunto公式","https://www.suunto.com/wing","https://www.suunto.com/","support@suunto.com","★★★★★","骨伝導・心拍センサー内蔵・IP67防水のスポーツ向け開放型イヤホン。ランニング・サイクリングで周囲音を聞きながら音楽を楽しめる。"],
  [2325,"ウェアラブル・ヘルス・フィットネス","Jabra Evolve2 75 Wireless Headset","Jabra","Jabra公式","https://www.jabra.com/evolve2-75","https://www.jabra.com/","support@jabra.com","★★★★★","36時間バッテリー・AIノイズキャンセリング・Microsoft Teams認定のプレミアムビジネスヘッドセット。在宅勤務需要増加の日本向け。"],
  [2326,"ウェアラブル・ヘルス・フィットネス","Beats Studio Pro Wireless Headphones","Beats by Dr. Dre","Amazon US","https://www.beatsbydre.com/studio-pro","https://www.beatsbydre.com/","support@beatsbydre.com","★★★★★","40時間バッテリー・ANC・USB-C/Lightning両対応のプレミアムワイヤレスヘッドフォン。Apple+Beatsの親和性でiPhoneユーザー向けに最適。"],
  [2327,"ウェアラブル・ヘルス・フィットネス","Sony WH-1000XM6 Wireless Headphones","Sony","Sony公式","https://www.sony.com/wh-1000xm6","https://www.sony.com/","support@sony.com","★★★★★","業界最高峰ANC・30時間・マルチポイント接続の次世代ノイキャンヘッドフォン。ソニーは日本ブランドだが海外版として逆輸入紹介価値あり。"],
  [2328,"ウェアラブル・ヘルス・フィットネス","Withings ScanWatch 2","Withings","Withings公式","https://www.withings.com/scanwatch-2","https://www.withings.com/","hello@withings.com","★★★★★","ECG・血中酸素・睡眠無呼吸スクリーニング機能搭載のハイブリッドスマートウォッチ。医療グレードセンサーをアナログ時計に詰め込んだ傑作。"],
  [2329,"ウェアラブル・ヘルス・フィットネス","Tile Ultra Tracker","Tile by Life360","Amazon US","https://www.tile.com/ultra","https://www.tile.com/","support@tile.com","★★★★★","UWB精密位置特定・Alexa音声検索・Apple Find My対応のプレミアムBluetooth/UWBトラッカー。財布・鍵・バッグの紛失対策。"],
  [2330,"ウェアラブル・ヘルス・フィットネス","Garmin Lily 2 Smartwatch for Women","Garmin","Garmin公式","https://www.garmin.com/lily-2","https://www.garmin.com/","support@garmin.com","★★★★★","小型・ファッショナブル・月経周期管理・ストレス検知搭載の女性向けスマートウォッチ。シンプルなデザインと豊富な健康機能を両立。"],
  [2331,"ウェアラブル・ヘルス・フィットネス","Bose QuietComfort Ultra Earbuds","Bose","Bose公式","https://www.bose.com/qc-ultra-earbuds","https://www.bose.com/","support@bose.com","★★★★★","Immersive Audio空間音響・CustomTune音場補正・World mode搭載のBose最上位ワイヤレスイヤホン。ANC性能は業界トップクラス。"],
  [2332,"ウェアラブル・ヘルス・フィットネス","Samsung Galaxy Ring","Samsung","Samsung公式","https://www.samsung.com/galaxy-ring","https://www.samsung.com/","https://www.samsung.com/us/support/contact/","★★★★★","Samsung初のスマートリング。睡眠・心拍・活動量・エネルギースコアをトラッキング。Galaxyエコシステムとの深い連携が強み。"],
  [2333,"ウェアラブル・ヘルス・フィットネス","Apple Watch Ultra 2","Apple","Apple公式","https://www.apple.com/apple-watch-ultra-2","https://www.apple.com/","https://www.apple.com/contact/","★★★★★","チタンケース・49mm・高精度GPS・Action Button搭載のApple Watch最上位モデル。登山・海洋スポーツ・トライアスロン向け究極のスマートウォッチ。"],
  [2334,"ウェアラブル・ヘルス・フィットネス","Jabra Enhance Plus OTC Hearing Aid","Jabra","Jabra公式","https://www.jabra.com/enhance","https://www.jabra.com/","support@jabra.com","★★★★★","処方箋不要のOTC補聴器。普通のイヤホンに見えるデザインでスティグマなく聴力を補助。高齢化社会の日本で潜在的に巨大な需要。"],
  [2335,"ウェアラブル・ヘルス・フィットネス","Nuheara IQbuds2 MAX Hearing Enhancement","Nuheara","Nuheara公式","https://nuheara.com/","https://nuheara.com/","support@nuheara.com","★★★★★","AIで聴力プロファイルを作成し個別最適化するヒアリングエンハンスメントイヤーバッド。難聴の早期対策として日本の高齢者向けに有望。"],
  [2336,"ウェアラブル・ヘルス・フィットネス","Nuro Duo Transcranial DC Stimulation","Nuro","Nuro公式","https://nuro.ai/","https://nuro.ai/","info@nuro.ai","★★★★★","経頭蓋直流刺激（tDCS）で集中力・学習能力・気分を高めるニューロフィードバックデバイス。ブレインハック需要の高い現代人向け。"],
  [2337,"ウェアラブル・ヘルス・フィットネス","Pulsetto Vagus Nerve Stimulator","Pulsetto","Pulsetto公式","https://pulsetto.tech/","https://pulsetto.tech/","support@pulsetto.tech","★★★★★","首に装着して迷走神経を電気刺激するウェアラブルデバイス。ストレス軽減・睡眠改善・集中力向上をデータで実証。現代日本のストレス対策向け。"],
  [2338,"ウェアラブル・ヘルス・フィットネス","Core Meditation Trainer Biofeedback","Core","Core公式","https://hellocore.com/","https://hellocore.com/","hello@hellocore.com","★★★★★","心拍バイオフィードバックと誘導瞑想を組み合わせたスマートな瞑想トレーニングデバイス。マインドフルネス実践者向け。"],
  [2339,"ウェアラブル・ヘルス・フィットネス","Dreem 3 Clinical Sleep EEG Headband","Dreem","Dreem公式","https://dreem.com/","https://dreem.com/","info@dreem.com","★★★★★","EEG脳波で睡眠段階を精密測定・音響神経刺激で深睡眠を増加させるクリニカルグレード睡眠改善ヘッドバンド。不眠に悩む日本人向け。"],
  [2340,"ウェアラブル・ヘルス・フィットネス","Zephyr HxM Smart Heart Rate Monitor","Zephyr","Zephyr公式","https://zephyranywhere.com/","https://zephyranywhere.com/","info@zephyranywhere.com","★★★★☆","スマートフォン連動の Bluetooth 胸部心拍計。ランニングアプリとリアルタイム連携で心拍トレーニングゾーン管理が可能。"],
  [2341,"ウェアラブル・ヘルス・フィットネス","ViMove2 Wearable Movement Analysis","DorsaVi","DorsaVi公式","https://www.dorsavi.com/","https://www.dorsavi.com/","info@dorsavi.com","★★★★★","背中・腰・膝の動きをリアルタイム分析してケガリスクを評価するウェアラブル動作分析システム。スポーツ・リハビリ・高齢者ケア向け。"],
  [2342,"ウェアラブル・ヘルス・フィットネス","Omron HeartGuide Wearable Blood Pressure Watch","Omron","Omron公式","https://www.omron-healthcare.com/heartguide","https://www.omron-healthcare.com/","support@omron.com","★★★★★","世界初の腕時計型血圧計。日本オムロン発の革新製品として海外市場に逆輸入紹介。高血圧が増加する日本と世界の健康管理向け。"],
  [2343,"ウェアラブル・ヘルス・フィットネス","Masimo W1 Health Wearable","Masimo","Masimo公式","https://www.masimo.com/w1","https://www.masimo.com/","support@masimo.com","★★★★★","血中酸素・脈拍・呼吸数・皮膚温度・活動量を医療グレードで連続測定するスマートウォッチ型ヘルスウェアラブル。臨床精度を一般向けに。"],

  // ===== アウトドア・スポーツ・旅行 (2344-2363) =====
  [2344,"アウトドア・スポーツ・旅行","Wahoo KICKR Core Smart Bike Trainer","Wahoo Fitness","Wahoo公式","https://www.wahoofitness.com/kickr-core","https://www.wahoofitness.com/","support@wahoofitness.com","★★★★★","Zwift・RGT対応・1400W・±2%精度のスマートインドアバイクトレーナー。サイクリストのインドアトレーニング革命製品。日本でも人気急上昇。"],
  [2345,"アウトドア・スポーツ・旅行","Therabody RecoveryAir JetBoots","Therabody","Therabody公式","https://www.therabody.com/recoveryair","https://www.therabody.com/","support@therabody.com","★★★★★","AIで回復プログラムを個別最適化するブーツ型空気圧マッサージデバイス。NFLチームも採用するプロアスリート向けリカバリーツール。"],
  [2346,"アウトドア・スポーツ・旅行","Hyperice Hypervolt Go 2 Massage Gun","Hyperice","Hyperice公式","https://hyperice.com/hypervolt-go-2","https://hyperice.com/","support@hyperice.com","★★★★★","5段階強度・40dB静音・コンパクトな携帯型パーカッションマッサージガン。NBA公式リカバリーパートナーブランドの軽量モデル。"],
  [2347,"アウトドア・スポーツ・旅行","Theragun Mini 2.0 Massage Device","Therabody","Therabody公式","https://www.therabody.com/theragun-mini-2","https://www.therabody.com/","support@therabody.com","★★★★★","手のひらサイズ・3段階強度・150分バッテリーのコンパクトマッサージガン。旅行・通勤バッグに入れて持ち歩けるTherapody入門製品。"],
  [2348,"アウトドア・スポーツ・旅行","NordicTrack Commercial 1750 Treadmill","NordicTrack","NordicTrack公式","https://www.nordictrack.com/treadmills/1750","https://www.nordictrack.com/","support@nordictrack.com","★★★★★","iFit AI対応・傾斜±15%・32km/h・14インチスクリーン搭載のスマートランニングマシン。自宅トレーニング需要の高まりで日本向けに有望。"],
  [2349,"アウトドア・スポーツ・旅行","Peloton Bike+ Connected Fitness","Peloton","Peloton公式","https://www.onepeloton.com/bike-plus","https://www.onepeloton.com/","support@onepeloton.com","★★★★★","23.8インチ回転スクリーン・自動抵抗調整のコネクテッドフィットネスバイク。ライブクラス・オンデマンドでスタジオ体験を自宅で再現。"],
  [2350,"アウトドア・スポーツ・旅行","Mirror Interactive Home Gym by Lululemon","Lululemon","Lululemon公式","https://www.lululemon.com/mirror","https://www.lululemon.com/","support@lululemon.com","★★★★★","スマートミラーにフィットネスインストラクターが映り、AIがリアルタイムフォーム指導するホームジムミラー。狭い日本の住宅でも設置可能。"],
  [2351,"アウトドア・スポーツ・旅行","Tonal Smart Home Gym","Tonal","Tonal公式","https://www.tonal.com/","https://www.tonal.com/","support@tonal.com","★★★★★","AIが最適な重量を自動設定する壁掛け型スマートホームジム。全身筋トレがこれ一台で完結。ホームジム需要が高まる日本向けに最適。"],
  [2352,"アウトドア・スポーツ・旅行","Hydrow Wave Rowing Machine","Hydrow","Hydrow公式","https://hydrow.com/wave","https://hydrow.com/","support@hydrow.com","★★★★★","スリムデザイン・ライブ&オンデマンド水上ローイング体験のコネクテッドローイングマシン。全身運動として人気のボート競技を自宅で再現。"],
  [2353,"アウトドア・スポーツ・旅行","Bowflex Velocore Bike Connected Fitness","Bowflex","Bowflex公式","https://www.bowflex.com/velocore","https://www.bowflex.com/","support@bowflex.com","★★★★★","左右に傾くLean Mode・16インチスクリーン・JRNY AI搭載の没入型コネクテッドバイク。コアトレーニング効果も高い革新的フィットネス製品。"],
  [2354,"アウトドア・スポーツ・旅行","Ergatta Water Rowing Machine","Ergatta","Ergatta公式","https://ergatta.com/","https://ergatta.com/","support@ergatta.com","★★★★★","ゲーム型コンペティティブトレーニング搭載のウォータータンク式スマートローイングマシン。競争要素でモチベーションを維持する革新設計。"],
  [2355,"アウトドア・スポーツ・旅行","Echelon EX-5s Connected Spin Bike","Echelon","Echelon公式","https://www.echelon.com/ex-5s","https://www.echelon.com/","support@echelon.com","★★★★☆","Pelotonより低価格・タブレット連動・LIVEクラス対応のコネクテッドスピンバイク。Pelotonのコスパ代替品として世界的に急成長。"],
  [2356,"アウトドア・スポーツ・旅行","Therabody Vybe Pro Massage Gun","Therabody","Therabody公式","https://www.therabody.com/vybe-pro","https://www.therabody.com/","support@therabody.com","★★★★☆","8段階強度・2400回転/分・スマートモード搭載のコスパ重視パーカッションマッサージガン。Therabodyの入門価格帯モデル。"],
  [2357,"アウトドア・スポーツ・旅行","Trekz Titanium Bone Conduction Headphones","Shokz","Amazon US","https://shokz.com/","https://shokz.com/","support@shokz.com","★★★★★","骨伝導・IP55防水・6時間再生・オープンイヤーのスポーツ向けヘッドフォン。ランニング・サイクリング中の安全な音楽体験を実現。"],
  [2358,"アウトドア・スポーツ・旅行","Wahoo ELEMNT ROAM GPS Bike Computer","Wahoo Fitness","Wahoo公式","https://www.wahoofitness.com/elemnt-roam","https://www.wahoofitness.com/","support@wahoofitness.com","★★★★★","カラースクリーン・マルチバンドGPS・Strava連動のサイクルコンピューター。Garmin EdgeとシェアNo.1を争うプレミアムブランド。"],
  [2359,"アウトドア・スポーツ・旅行","Coros Pace 3 GPS Running Watch","Coros","Coros公式","https://www.coros.com/pace-3","https://www.coros.com/","support@coros.com","★★★★★","デュアルバンドGPS・17日バッテリー・VO2max・レースETA予測搭載の軽量ランニングウォッチ。Garminより軽量でコスパ最高。"],
  [2360,"アウトドア・スポーツ・旅行","Amazfit T-Rex 3 Adventure Watch","Amazfit","Amazfit公式","https://www.amazfit.com/t-rex-3","https://www.amazfit.com/","support@amazfit.com","★★★★☆","MIL-STD-810G・23日バッテリー・デュアルバンドGPS搭載の耐久性重視アウトドアスポーツウォッチ。登山・アドベンチャー向け。"],
  [2361,"アウトドア・スポーツ・旅行","Garmin Approach S70 Golf GPS Watch","Garmin","Garmin公式","https://www.garmin.com/approach-s70","https://www.garmin.com/","support@garmin.com","★★★★★","4.2万以上のゴルフコースマップ・スウィング計測・スコアカード機能搭載のゴルフGPSスマートウォッチ。ゴルフ大国日本向けに特化した製品。"],
  [2362,"アウトドア・スポーツ・旅行","TopGolf Callaway Golf Radar Launch Monitor","Callaway","Callaway公式","https://www.callawaygolf.com/","https://www.callawaygolf.com/","support@callawaygolf.com","★★★★★","ドップラーレーダーでボール速度・スピン・弾道をリアルタイム計測するゴルフ打球測定器。プロ・アマゴルファー向け練習革命ツール。"],
  [2363,"アウトドア・スポーツ・旅行","Bushnell Pro X3 Golf Laser Rangefinder","Bushnell","Amazon US","https://www.bushnell.com/pro-x3","https://www.bushnell.com/","support@bushnell.com","★★★★★","スロープ計算・PinSeeker・BITE磁気クリップ付きのプレミアムゴルフ用レーザー距離計。ゴルフ愛好者の日本市場向けに有望な製品。"],

  // ===== ペット用品 (2364-2383) =====
  [2364,"ペット用品","Minitailz Smart Pet Health Tracker 2025","Invoxia","Invoxia公式","https://www.invoxia.com/minitailz","https://www.invoxia.com/","hello@invoxia.com","★★★★★","心拍・呼吸・活動量・GPS追跡をリアルタイムモニタリングするペットの究極健康トラッカー。獣医師との連携機能も搭載。"],
  [2365,"ペット用品","Petnet SmartFeeder 2.0","Petnet","Petnet公式","https://petnet.io/","https://petnet.io/","support@petnet.io","★★★★☆","AIで体重・活動量・犬種に応じた最適給餌量を計算するスマート自動給餌器。肥満予防と健康的な食事管理を自動化。"],
  [2366,"ペット用品","Vava Pet Camera with Night Vision","VAVA","Amazon US","https://www.vava.com/","https://www.vava.com/","support@vava.com","★★★★☆","1080p・フルカラーナイトビジョン・双方向音声のペット見守りカメラ。コスパが高く機能充実でペット飼育者の留守番監視向け。"],
  [2367,"ペット用品","CleanPaws Paw Plunger Premium","CleanPaws","Amazon US","https://cleanpaws.us/","https://cleanpaws.us/","support@cleanpaws.us","★★★★☆","散歩帰りに足を突っ込むだけで汚れを落とす大型犬対応パウプランジャー。清潔好きな日本の飼育者向け定番ペット用品。"],
  [2368,"ペット用品","Radio Systems PetSafe Drinkwell Platinum","Radio Systems (PetSafe)","Amazon US","https://www.petsafe.net/drinkwell-platinum","https://www.petsafe.net/","support@petsafe.net","★★★★★","1.5L大容量・流水ポンプ・炭フィルター搭載のペット用流水給水器。猫の腎臓病予防に効果的な水分摂取促進ができる製品。"],
  [2369,"ペット用品","Scoopfree Crystal Cat Litter Tray","PetSafe","Amazon US","https://www.petsafe.net/scoopfree-crystal","https://www.petsafe.net/","support@petsafe.net","★★★★☆","クリスタルリターで臭いを強力吸収・使い捨てトレー交換式の猫トイレシステム。猫トイレ掃除の手間を最小化する日本の猫飼育者向け。"],
  [2370,"ペット用品","Bissell BarkBath Portable Dog Bath System","Bissell","Amazon US","https://www.bissell.com/barkbath","https://www.bissell.com/","support@bissell.com","★★★★★","犬の毛に直接水を当てて洗えるポータブルシャワーシステム。浴室に連れて行かずに自宅でどこでも犬を洗えるユニークな製品。"],
  [2371,"ペット用品","Sleepypod Air In-Cabin Dog Cat Carrier","Sleepypod","Amazon US","https://sleepypod.com/air","https://sleepypod.com/","support@sleepypod.com","★★★★★","航空会社認定の機内持ち込み可能なペットキャリー。安全テスト済みで愛犬愛猫を機内に連れて行ける唯一無二の設計。旅行好きペット飼育者向け。"],
  [2372,"ペット用品","Ruffwear Front Range Dog Harness","Ruffwear","Amazon US","https://ruffwear.com/front-range","https://ruffwear.com/","support@ruffwear.com","★★★★★","デュアルクリップ・パッド入り・ID用ポケット搭載のプレミアムドッグハーネス。引っ張り防止・快適性・安全性を兼ね備えた定番製品。"],
  [2373,"ペット用品","PetSafe Gentle Leader Head Collar","PetSafe","Amazon US","https://www.petsafe.net/gentle-leader","https://www.petsafe.net/","support@petsafe.net","★★★★☆","引っ張り防止設計のヘッドカラー。鼻先にリードを接続して引っ張り癖を自然に矯正。しつけに悩む日本の犬飼育者向けに効果的。"],
  [2374,"ペット用品","Roam Odyssey Cat Backpack Carrier","Roam","Amazon US","https://roampets.com/","https://roampets.com/","support@roampets.com","★★★★★","バブル窓・通気性・容量拡張できる猫用バックパックキャリー。SNSで人気の猫バックパックの中でも機能性と安全性が高い製品。"],
  [2375,"ペット用品","Outward Hound Puzzle Brick Feeder","Outward Hound","Amazon US","https://www.outwardhound.com/","https://www.outwardhound.com/","support@outwardhound.com","★★★★☆","食事探索で脳を刺激するスローフィーダー型ブリックパズル。早食い防止・肥満対策・知育に役立つペット向け食事トイ。"],
  [2376,"ペット用品","Neater Feeder Deluxe Dog Bowl Set","Neater Pet Brands","Amazon US","https://neaterpetbrands.com/","https://neaterpetbrands.com/","support@neaterpetbrands.com","★★★★☆","こぼれた水・食べ物を下部トレーで受け止めるスピルプルーフフードスタンド。床を汚さない清潔な食事環境を提供。清潔好きな日本向け。"],
  [2377,"ペット用品","PetSafe Premium Busy Buddy Chew Toy","PetSafe","Amazon US","https://www.petsafe.net/busy-buddy","https://www.petsafe.net/","support@petsafe.net","★★★★☆","おやつを詰めて長時間噛み続けられる耐久性の高い噛みおもちゃ。留守番中の犬の退屈・ストレス解消に効果的なロングセラー製品。"],
  [2378,"ペット用品","iCat Fresh Baked Cat Treats Subscription","iCat","iCat公式","https://icat.com/","https://icat.com/","info@icat.com","★★★★☆","新鮮素材で焼いた猫用おやつを定期配送するサブスクリプションサービス。プレミアムペットフード需要の高まる日本の猫飼育者向け。"],
  [2379,"ペット用品","Zymox Veterinary Strength Pet Ear Solution","Zymox","Amazon US","https://www.zymox.com/","https://www.zymox.com/","support@zymox.com","★★★★★","処方箋なしで外耳炎・耳の感染症を治療できる獣医グレードの酵素系耳ケアソリューション。動物病院費用を削減したい日本の飼育者向け。"],
  [2380,"ペット用品","Thundershirt Anxiety Relief Dog Wrap","Thunderworks","Amazon US","https://www.thundershirt.com/","https://www.thundershirt.com/","support@thundershirt.com","★★★★★","雷・花火・不安・分離ストレスに対応する加圧ラップ。獣医師・トレーナー推奨の不安軽減ウェア。日本の花火大会・台風時期に特に需要大。"],
  [2381,"ペット用品","OmegaVet Vet-Strength Omega-3 for Dogs","OmegaVet","Amazon US","https://omegavet.com/","https://omegavet.com/","support@omegavet.com","★★★★☆","獣医推奨の超高濃度オメガ3魚油サプリメント。皮膚・被毛・関節・心臓の健康を総合サポート。愛犬の健康管理に積極的な日本向け。"],
  [2382,"ペット用品","Petkit Fresh Element Gemini Smart Feeder","PETKIT","Petkit公式","https://petkit.com/fresh-element-gemini","https://petkit.com/","support@petkit.com","★★★★★","2匹対応・Wi-Fi・カメラ内蔵・食事記録のデュアルボウルスマート自動給餌器。多頭飼いのペット管理を自動化する理想的製品。"],
  [2383,"ペット用品","BestVet Portable Dog Water Bottle","BestVet","Amazon US","https://bestvetpet.com/","https://bestvetpet.com/","support@bestvetpet.com","★★★★☆","ワンプッシュでボウルが出る携帯型犬用ウォーターボトル。散歩・ハイキングで愛犬に手軽に水を与えられる定番アウトドアペット用品。"],

  // ===== テクノロジー・ガジェット (2384-2403) =====
  [2384,"テクノロジー・ガジェット","Meta Quest 3S VR Headset","Meta","Meta公式","https://www.meta.com/quest-3s","https://www.meta.com/","support@meta.com","★★★★★","カラーパススルー・MR・最先端Snapdragon XR2搭載の最新VR/MRヘッドセット。スタンドアロンでPCなしに没入体験が可能。エンタメ・教育向け。"],
  [2385,"テクノロジー・ガジェット","Apple Vision Pro Spatial Computer","Apple","Apple公式","https://www.apple.com/apple-vision-pro","https://www.apple.com/","https://www.apple.com/contact/","★★★★★","空間コンピューティングを実現する画期的なMRヘッドセット。visionOSで仮想スクリーン・3Dアプリを空間に配置。次世代コンピューティング。"],
  [2386,"テクノロジー・ガジェット","Insta360 Go 3S Action Camera","Insta360","Insta360公式","https://www.insta360.com/go-3s","https://www.insta360.com/","support@insta360.com","★★★★★","4K・35g超小型・磁気クリップ装着のアクションカメラ。どこにでも装着でき日常生活・スポーツ・旅行を気軽に記録できる製品。"],
  [2387,"テクноロジー・ガジェット","DJI Osmo Pocket 3 Camera","DJI","DJI公式","https://www.dji.com/osmo-pocket-3","https://www.dji.com/","support@dji.com","★★★★★","1インチCMOS・4K120fps・3軸ジンバル搭載のコンパクトポケットカメラ。Vlog・旅行・日常記録向けに滑らかな映像を簡単撮影。"],
  [2388,"テクノロジー・ガジェット","Rode NT-USB Mini Studio Microphone","Rode","Rode公式","https://rode.com/nt-usb-mini","https://rode.com/","support@rode.com","★★★★★","スタジオ品質のコンデンサーマイクがUSB接続一本で使えるポッドキャスト・配信向けマイク。ホームスタジオの音質を劇的に向上。"],
  [2389,"テクノロジー・ガジェット","Govee AI Gaming Sync Box","Govee","Govee公式","https://www.govee.com/","https://www.govee.com/","marketing@govee.com","★★★★★","AIカメラでゲーム画面を解析してAmbilight効果を実現するスマートLEDシンクボックス。ゲーミング没入感を圧倒的に向上させる製品。"],
  [2390,"テクノロジー・ガジェット","Logitech G Pro X 2 Lightspeed Gaming Mouse","Logitech","Logitech公式","https://www.logitechg.com/g-pro-x-2","https://www.logitech.com/","support@logitech.com","★★★★★","HERO 25Kセンサー・Lightspeed無線・95g軽量のプロゲーマー向けワイヤレスゲーミングマウス。Eスポーツ選手御用達のフラッグシップモデル。"],
  [2391,"テクノロジー・ガジェット","Razer Blade 16 Gaming Laptop 2024","Razer","Razer公式","https://www.razer.com/blade-16","https://www.razer.com/","support@razer.com","★★★★★","OLED・RTX 4090・240Hz・16インチのプレミアムゲーミングラップトップ。薄型ながら最高スペックを詰め込んだゲーミングPC最高峰。"],
  [2392,"テクノロジー・ガジェット","Asus ROG Strix OLED Gaming Monitor 27\"","Asus ROG","Asus公式","https://rog.asus.com/monitors/","https://www.asus.com/","support@asus.com","★★★★★","27インチ・OLED・240Hz・0.03ms応答速度のプレミアムゲーミングモニター。コントラスト・色再現・応答速度の全てで最高スペック。"],
  [2393,"テクノロジー・ガジェット","Samsung Odyssey G9 Curved Gaming Monitor","Samsung","Samsung公式","https://www.samsung.com/odyssey-g9","https://www.samsung.com/","https://www.samsung.com/us/support/contact/","★★★★★","49インチ・DQHD・240Hz・デュアル4K超ワイド湾曲ゲーミングモニター。究極のゲーム・マルチタスク体験を一台で実現する超大画面。"],
  [2394,"テクноロジー・ガジェット","Creative Sound Blaster X5 USB DAC Amp","Creative","Creative公式","https://us.creative.com/sound-blaster-x5","https://us.creative.com/","support@creative.com","★★★★★","Xamp独立バイアンプ・Hi-Res・Dolby Atmos対応のプレミアムUSB DAC/ヘッドフォンアンプ。ゲーム・音楽・映像のオーディオ品質を劇的向上。"],
  [2395,"テクノロジー・ガジェット","Anker Soundcore Motion X600 Bluetooth Speaker","Anker","Amazon US","https://www.soundcore.com/motion-x600","https://www.soundcore.com/","support@soundcore.com","★★★★★","50W・空間オーディオ・IP67防水のポータブルBluetoothスピーカー。Ankerの最上位スピーカーで音質とコスパを両立。アウトドア向け。"],
  [2396,"テクノロジー・ガジェット","JBL Xtreme 4 Waterproof Speaker","JBL","JBL公式","https://www.jbl.com/xtreme4","https://www.jbl.com/","support@jbl.com","★★★★★","IP67防水・24時間・パワーバンク機能搭載のJBL大型ポータブルスピーカー。キャンプ・ビーチ・パーティー向けに人気の定番製品。"],
  [2397,"テクノロジー・ガジェット","Sonos Move 2 Portable Smart Speaker","Sonos","Sonos公式","https://www.sonos.com/move-2","https://www.sonos.com/","support@sonos.com","★★★★★","24時間・IP56防水・自動TruePlay音場補正・Wi-Fi+Bluetooth のポータブルスマートスピーカー。屋内外シームレスに使える最高品質。"],
  [2398,"テクノロジー・ガジェット","Jackery Explorer 300 Plus Power Station","Jackery","Jackery公式","https://www.jackery.com/explorer-300-plus","https://www.jackery.com/","support@jackery.com","★★★★★","288Wh・コンパクト・ソーラー充電対応の小型ポータブル電源。キャンプ・防災向け入門製品として手頃な価格帯でJackeryブランドを体験。"],
  [2399,"テクノロジー・ガジェット","Moft Z Laptop Stand 5-in-1","Moft","Amazon US","https://www.moft.us/moft-z","https://www.moft.us/","support@moft.us","★★★★★","5段階高さ調整・折り畳みコンパクトのアルミニウムラップトップスタンド。在宅勤務・カフェワーク向けに姿勢改善と作業効率を向上。"],
  [2400,"テクノロジー・ガジェット","Twelve South MagSafe BookArc for Mac","Twelve South","Amazon US","https://www.twelvesouth.com/bookarc","https://www.twelvesouth.com/","support@twelvesouth.com","★★★★★","MacBookを縦置きしてデスクスペースを節約するステンレス製スタンド。Mac・Apple製品愛好家向けのプレミアムアクセサリーブランド。"],
  [2401,"テクノロジー・ガジェット","CalDigit TS4 Thunderbolt 4 Dock","CalDigit","CalDigit公式","https://www.caldigit.com/ts4","https://www.caldigit.com/","support@caldigit.com","★★★★★","18ポート・98W充電・Thunderbolt 4接続のプレミアムドッキングステーション。MacユーザーのデスクセットアップでNo.1評価の定番ドック。"],
  [2402,"テクノロジー・ガジェット","Samsung T7 Shield Portable SSD","Samsung","Amazon US","https://www.samsung.com/t7-shield","https://www.samsung.com/","https://www.samsung.com/us/support/contact/","★★★★★","IP65耐塵防水・1050MB/s・ラバーボディの耐衝撃ポータブルSSD。クリエイター・旅行者向けに高速データ転送と高耐久性を両立。"],
  [2403,"テクノロジー・ガジェット","WD My Passport 4TB Portable Hard Drive","Western Digital","Amazon US","https://www.westerndigital.com/my-passport","https://www.westerndigital.com/","support@westerndigital.com","★★★★☆","4TB・USB-C・ハードウェア暗号化のポータブルHDD。写真・動画・データバックアップ向けの定番大容量外付けストレージ。"],

  // ===== 美容・スキンケア (2404-2423) =====
  [2404,"美容・スキンケア","Dennis Gross DRx SpectraLite FaceWare Pro","Dr. Dennis Gross","Dennis Gross公式","https://www.dgskincare.com/spectraLite","https://www.dgskincare.com/","support@dgskincare.com","★★★★★","FDA認可・100個LED・赤&近赤外&青色光の組み合わせフェイスマスク。1日3分のLED治療でアンチエイジング・ニキビ対策を実現。"],
  [2405,"美容・スキンケア","CurrentBody Skin LED Neck & Dec Perfector","CurrentBody","CurrentBody公式","https://www.currentbody.com/neck-dec","https://www.currentbody.com/","hello@currentbody.com","★★★★★","首・デコルテ専用LED光療法デバイス。セルフケアが難しい首・胸元のハリ・シワ・色素沈着を改善。日本女性が気にする首元ケア向け。"],
  [2406,"美容・スキンケア","Glov On The Go Makeup Remover Glove","Glov","Glov公式","https://glovglobal.com/","https://glovglobal.com/","support@glovglobal.com","★★★★☆","水だけでメイクを完全オフできるマイクロファイバーグローブ。クレンジング剤不要・環境に優しいエコ美容ツール。クリーンビューティー向け。"],
  [2407,"美容・スキンケア","Agent Nateur Holi(Youth) Pearl Serum","Agent Nateur","Agent Nateur公式","https://agentnateur.com/","https://agentnateur.com/","support@agentnateur.com","★★★★★","真珠パウダー・レチノール・ヒアルロン酸配合のクリーンラグジュアリー美容液。セレブ愛用のプレミアムクリーンビューティーブランド。"],
  [2408,"美容・スキンケア","Epara Brightening Face Oil","Epara Skincare","Epara公式","https://eparaskincare.com/","https://eparaskincare.com/","support@eparaskincare.com","★★★★★","アフリカン・ボタニカルズ配合のプレミアムフェイスオイル。有色人種の肌のために開発されたダイバーシティスキンケアブランド。"],
  [2409,"美容・スキンケア","Goop G.Tox Detox Scalp Serum","Goop","Goop公式","https://goop.com/g-tox-detox-scalp","https://goop.com/","support@goop.com","★★★★☆","グウィネス・パルトロウ監修のラグジュアリーウェルネスブランドGoopの頭皮デトックスセラム。インフルエンサー経済圏で日本でも話題性大。"],
  [2410,"美容・スキンケア","Tatcha Violet-C Brightening Serum","Tatcha","Tatcha公式","https://www.tatcha.com/violet-c","https://www.tatcha.com/","support@tatcha.com","★★★★★","20%ビタミンC+10%AHA配合の日本の伝統美容からインスパイアされたブライトニング美容液。日本文化の美容哲学を世界に発信するブランド。"],
  [2411,"美容・スキンケア","Biossance Squalane + Vitamin C Rose Oil","Biossance","Biossance公式","https://biossance.com/","https://biossance.com/","support@biossance.com","★★★★★","サトウキビ由来スクワラン+ビタミンC配合のサステナブルクリーン美容オイル。環境・肌に優しいサスティナブルビューティーブランド。"],
  [2412,"美容・スキンケア","Supergoop Unseen Sunscreen SPF 40","Supergoop","Supergoop公式","https://supergoop.com/unseen","https://supergoop.com/","support@supergoop.com","★★★★★","塗っても白くならない・さらっとしたテクスチャーの日焼け止め下地。SPFへの意識を変えたゲームチェンジャーサンスクリーンブランド。"],
  [2413,"美容・スキンケア","Krave Beauty Great Barrier Relief Serum","Krave Beauty","Krave公式","https://kravebeauty.com/","https://kravebeauty.com/","support@kravebeauty.com","★★★★★","ローズヒップ種子油+パントテン酸配合のバリア修復美容液。ミニマルスキンケアムーブメントをリードするKorean美容ブランド。"],
  [2414,"美容・スキンケア","The Inkey List Retinol Serum","The Inkey List","Amazon US","https://www.theinkeylist.com/retinol","https://www.theinkeylist.com/","support@theinkeylist.com","★★★★★","0.5%レチノール・スクワラン・タウリン配合の高コスパレチノール美容液。成分の透明性とコスパで世界中にファンを持つブランド。"],
  [2415,"美容・スキンケア","Fenty Beauty Skin Aqua Serum Foundation","Fenty Beauty","Fenty Beauty公式","https://fentybeauty.com/aqua-serum","https://fentybeauty.com/","support@fentybeauty.com","★★★★★","リアーナ監修・50色展開のインクルーシブ美容ブランドの軽いセラムファンデ。多様性・インクルーシビティを重視する日本の若い世代向け。"],
  [2416,"美容・スキンケア","Rhode Skin Peptide Glazing Fluid","Rhode","Rhode公式","https://rhodeskin.com/glazing-fluid","https://rhodeskin.com/","support@rhodeskin.com","★★★★★","ヘイリー・ビーバー設立のバイラルグレーズドスキン美容液。TikTokで世界的にバイラルになったペプチド保湿美容液。Z世代向け。"],
  [2417,"美容・スキンケア","e.l.f. Halo Glow Liquid Filter","e.l.f. Cosmetics","Amazon US","https://www.elfcosmetics.com/halo-glow","https://www.elfcosmetics.com/","support@elfcosmetics.com","★★★★★","3倍価格のシャーロット・ティルバリー比較でバイラルになったコスパ最強のグロウ下地。TikTokで数百万回再生の話題製品。"],
  [2418,"美容・スキンケア","Gua Sha Facial Tool Rose Quartz","Lanshin","Amazon US","https://lanshin.com/","https://lanshin.com/","support@lanshin.com","★★★★☆","ローズクォーツ製グアシャ美顔かっさ。リンパドレナージュ・フェイスラインリフトアップ・血行促進効果。中国伝統美容の現代版として世界的人気。"],
  [2419,"美容・スキンケア","Herbivore Botanicals Prism 20% AHA+5% BHA","Herbivore Botanicals","Herbivore公式","https://www.herbivorebotanicals.com/prism","https://www.herbivorebotanicals.com/","support@herbivorebotanicals.com","★★★★★","クリーン・ビーガン・クルエルティーフリーのAHA/BHAエクスフォリアントセラム。スキンケアの透明性を重視するクリーンビューティーブランド。"],
  [2420,"美容・スキンケア","Charlotte Tilbury Flawless Filter Primer","Charlotte Tilbury","Charlotte Tilbury公式","https://www.charlottetilbury.com/flawless-filter","https://www.charlottetilbury.com/","support@charlottetilbury.com","★★★★★","セレブメイクアーティスト監修の話題の下地。天使の皮膚を作ると評判のマルチユースプライマー&ハイライター。TikTokでバイラル人気。"],
  [2421,"美容・スキンケア","Milk Makeup Cooling Water Jelly Tint","Milk Makeup","Milk Makeup公式","https://www.milkmakeup.com/cooling-water","https://www.milkmakeup.com/","support@milkmakeup.com","★★★★☆","リップ&チーク兼用・スティックタイプのジェリーティント。Sephoraで人気のニューヨーク発クリーンメイクブランド。若い世代向け。"],
  [2422,"美容・スキンケア","ILIA Super Serum Skin Tint SPF 40","ILIA Beauty","ILIA公式","https://iliabeauty.com/super-serum","https://iliabeauty.com/","support@iliabeauty.com","★★★★★","SPF+美容液+スキンティントを一本で実現するクリーンビューティーの傑作。日焼け止めとスキンケアを同時に行いたい日本の女性向け。"],
  [2423,"美容・スキンケア","Beauty Pie Japanfusion Pure Transforming Cleanser","Beauty Pie","Beauty Pie公式","https://www.beautypie.com/japanfusion","https://www.beautypie.com/","support@beautypie.com","★★★★★","日本の炭素・活性炭・酵素配合の会員制高級スキンケアブランドのクレンジング。日本の美容成分を欧州から逆紹介する独自のポジション。"],

  // ===== 子供・教育 (2424-2443) =====
  [2424,"子供・教育","Khan Academy Kids Free Learning App","Khan Academy","Khan Academy公式","https://www.khanacademy.org/kids","https://www.khanacademy.org/","https://www.khanacademy.org/contact","★★★★★","AIでパーソナライズ学習する無料教育プラットフォーム。幼児〜K-5向け算数・読書・SEL・創造的表現を楽しく学べる定番サービス。"],
  [2425,"子供・教育","Highlights Hidden Pictures Puzzle Book","Highlights","Amazon US","https://www.highlights.com/","https://www.highlights.com/","support@highlights.com","★★★★☆","世界で最も読まれている子供向け隠し絵・パズル雑誌。集中力・観察力・問題解決能力を育む。1946年創刊の米国老舗教育ブランド。"],
  [2426,"子供・教育","Melissa & Doug Deluxe Pounding Bench","Melissa & Doug","Amazon US","https://www.melissaanddoug.com/","https://www.melissaanddoug.com/","support@melissaanddoug.com","★★★★★","木製・色認識・手眼協調を育てる幼児向けハンマリングベンチ玩具。毒素フリーの安心素材で世界中の親に信頼される教育玩具ブランド。"],
  [2427,"子供・教育","Roblox Gift Card for Kids Gaming","Roblox","Amazon US","https://www.roblox.com/","https://www.roblox.com/","support@roblox.com","★★★★☆","世界2億人以上が遊ぶユーザー制作ゲームプラットフォームのギフトカード。ゲームしながら創造性・プログラミング・社会性を育む。"],
  [2428,"子供・教育","Crayola Light Up Tracing Pad","Crayola","Amazon US","https://www.crayola.com/light-up-tracing","https://www.crayola.com/","support@crayola.com","★★★★☆","LEDライトパッドで絵のトレーシング・デザイン練習ができる子供向けアート教育ツール。描くことの楽しさと美術の基礎を育成。"],
  [2429,"子供・教育","LeapFrog Scribble and Write Learning Pad","LeapFrog","Amazon US","https://www.leapfrog.com/scribble-write","https://www.leapfrog.com/","support@leapfrog.com","★★★★★","文字・数字・絵の描き方を音声ガイドで教える幼児向けライティング学習パッド。書き順・発音を楽しく学べる就学前教育ツール。"],
  [2430,"子供・教育","Osmo Little Genius Starter Kit for iPad","Osmo","Osmo公式","https://www.playosmo.com/little-genius","https://www.playosmo.com/","support@playosmo.com","★★★★★","3-5歳向け・iPadカメラでフィジカルアクションを検知するインタラクティブ幼児教育システム。STEAM入門として最適な体験型学習。"],
  [2431,"子供・教育","ThinkFun Gravity Maze Marble Run Logic Game","ThinkFun","Amazon US","https://www.thinkfun.com/gravitymaze","https://www.thinkfun.com/","support@thinkfun.com","★★★★★","60レベルの論理パズルで問題解決力・空間認識力を鍛えるマーブルランゲーム。STEM教育向けのベストセラー思考ゲーム。"],
  [2432,"子供・教育","National Geographic Kids Mega Science Kit","National Geographic","Amazon US","https://www.natgeokids.com/","https://www.natgeokids.com/","support@natgeo.com","★★★★★","地質・化学・物理実験セットが詰まった子供向けサイエンスキット。ナショナルジオグラフィックのブランド力で科学への興味を育む。"],
  [2433,"子供・教育","Hand2Mind VersaTiles Literacy System","hand2mind","Amazon US","https://www.hand2mind.com/versalites","https://www.hand2mind.com/","support@hand2mind.com","★★★★★","自己採点機能付きのK-5向けリテラシー・数学練習ブック。子供が一人で答え合わせができる革新的な学習システム。"],
  [2434,"子供・教育","BrainQuest Grade 2 Workbook","Workman Publishing","Amazon US","https://www.workman.com/brainquest","https://www.workman.com/","support@workman.com","★★★★☆","米国学習指導要領準拠のグレード別ワークブック。夏休みの学習ドリルとして世界中の日本人家庭でも使える英語教材。"],
  [2435,"子供・教育","Kiwi Co Kiwi Crate STEM Box Subscription","KiwiCo","KiwiCo公式","https://www.kiwico.com/kiwi-crate","https://www.kiwico.com/","hello@kiwico.com","★★★★★","5-8歳向けの毎月届くSTEM工作サブスクリプション。手を動かしながらサイエンス・テクノロジー・エンジニアリングを学ぶ。"],
  [2436,"子供・教育","Playmation Marvel Avengers Starter Pack","Hasbro","Amazon US","https://www.hasbro.com/","https://www.hasbro.com/","support@hasbro.com","★★★★☆","ウェアラブルデバイスとアクションフィギュアが連動するインタラクティブマーベルゲーム。デジタルとフィジカルの融合型玩具。"],
  [2437,"子供・教育","LEGO Technic 4x4 X-treme Off-Roader","LEGO","Amazon US","https://www.lego.com/technic","https://www.lego.com/","support@lego.com","★★★★★","エンジン・サスペンション・ステアリング機構の仕組みを学べるLEGO Technic最新キット。工学・機械への興味を育む高度なブロック製品。"],
  [2438,"子供・教育","Fat Brain Toys Tobbles Neo Stacking Toy","Fat Brain Toys","Amazon US","https://www.fatbraintoys.com/tobbles-neo","https://www.fatbraintoys.com/","support@fatbraintoys.com","★★★★★","重力・バランス・色を学ぶ乳幼児向けスタッキングトイ。シンプルながら奥深い遊びで長く楽しめる知育玩具の傑作。"],
  [2439,"子供・教育","Magnetic Tiles Magna-Tiles Clear Colors","Magna-Tiles","Amazon US","https://www.magnatiles.com/","https://www.magnatiles.com/","support@magnatiles.com","★★★★★","世界中の幼稚園・小学校で採用される磁石タイル構造玩具。2Dから3Dへの立体思考・数学的概念・創造性を自由な遊びで育む。"],
  [2440,"子供・教育","Goldie Blox Builder Kit for Girls","GoldieBlox","Amazon US","https://www.goldieblox.com/","https://www.goldieblox.com/","support@goldieblox.com","★★★★★","女の子向けにデザインされたエンジニアリング・建設遊びキット。STEM教育のジェンダーギャップ解消を目的として開発された革新製品。"],
  [2441,"子供・教育","Perler Beads Mega Activity Bucket","Perler","Amazon US","https://www.perlerbeads.com/","https://www.perlerbeads.com/","support@perlerbeads.com","★★★★☆","アイロンで溶着するフューズビーズ22,000個セット。色・形・パターンを学びながら独自デザインを制作。手先の器用さと創造性を育む。"],
  [2442,"子供・教育","Snap Circuits Jr. SC-100 Electronics Kit","Elenco","Amazon US","https://www.elenco.com/snap-circuits-jr","https://www.elenco.com/","support@elenco.com","★★★★★","ハンダ不要・スナップ式コネクターで100種の回路を組む入門電子工作キット。電気・電子の基礎を楽しく学べる世界的ベストセラー。"],
  [2443,"子供・教育","Play-Doh Kitchen Creations Ultimate Barbecue","Play-Doh","Amazon US","https://www.play-doh.com/","https://www.play-doh.com/","support@hasbro.com","★★★★☆","80以上のアクセサリー・型・BBQセット付きのPlay-Doh大型キッチンクリエーションセット。粘土遊びで創造性・想像力・手先の器用さを育む。"],

  // ===== ファッション・アクセサリー (2444-2463) =====
  [2444,"ファッション・アクセサリー","Coach Tabby Shoulder Bag","Coach","Coach公式","https://www.coach.com/tabby-shoulder-bag","https://www.coach.com/","support@coach.com","★★★★★","タビー金具・本革・多機能ポケットのCoachのシグネチャーショルダーバッグ。中価格帯ラグジュアリーバッグとして日本でも人気拡大中。"],
  [2445,"ファッション・アクセサリー","Kate Spade New York Knott Satchel","Kate Spade","Kate Spade公式","https://www.katespade.com/knott-satchel","https://www.katespade.com/","support@katespade.com","★★★★★","ニューヨーク発のカラフル・ファン・女性向けデザインのプレミアムサッチェルバッグ。ブランドの遊び心と実用性を両立した人気モデル。"],
  [2446,"ファッション・アクセサリー","Fossil Gen 6 Smartwatch Hybrid HR","Fossil","Amazon US","https://www.fossil.com/gen-6-hybrid-hr","https://www.fossil.com/","support@fossil.com","★★★★☆","本革ストラップ・e-Ink表示・心拍センサー搭載のハイブリッドスマートウォッチ。Alexa内蔵でアナログデザインのままスマートウォッチ機能を追加。"],
  [2447,"ファッション・アクセサリー","Rag & Bone Classic Newbury Boot","Rag & Bone","Rag & Bone公式","https://www.rag-bone.com/","https://www.rag-bone.com/","support@rag-bone.com","★★★★★","英国製・本革・ハンドステッチのプレミアムアンクルブーツ。ニューヨーク発のラグジュアリーブランドとして日本でも認知度が急上昇中。"],
  [2448,"ファッション・アクセサリー","Le Labo Santal 33 Eau de Parfum","Le Labo","Le Labo公式","https://www.lelabofragrances.com/santal-33","https://www.lelabofragrances.com/","support@lelabofragrances.com","★★★★★","サンダルウッド・バイオレット・レザーノートの世界で最も話題の香水。ニューヨーク発のアルチザン香水ブランドとして日本でも人気急上昇。"],
  [2449,"ファッション・アクセサリー","Byredo Bal d'Afrique Eau de Parfum","Byredo","Byredo公式","https://www.byredo.com/bal-dafrique","https://www.byredo.com/","support@byredo.com","★★★★★","アフリカのマリゴールド・アフリカンバイオレット・セダーウッドの高級香水。スウェーデン発のラグジュアリー香水ブランドで世界的人気。"],
  [2450,"ファッション・アクセサリー","Shinola Runwell Watch 41mm","Shinola","Shinola公式","https://www.shinola.com/runwell","https://www.shinola.com/","support@shinola.com","★★★★★","デトロイト製・Argonite1069ムーブメント搭載のアメリカンラグジュアリー時計。Made in USAというストーリーが日本市場でも訴求力大。"],
  [2451,"ファッション・アクセサリー","Frye Harness 8R Boot","Frye","Amazon US","https://www.thefryecompany.com/harness-8r","https://www.thefryecompany.com/","support@frye.com","★★★★★","1863年創業・1世紀以上の歴史を持つ米国老舗レザーブーツブランドの定番モデル。本革の経年変化を楽しむ製品として日本向けに訴求。"],
  [2452,"ファッション・アクセサリー","Madewell The Perfect Jean Medium Wash","Madewell","Madewell公式","https://www.madewell.com/perfect-jean","https://www.madewell.com/","support@madewell.com","★★★★★","100%オーガニックコットン・フラットリング・日本製デニムを使用したシグネチャージーンズ。日本のデニム文化との深い親和性が高い製品。"],
  [2453,"ファッション・アクセサリー","Alex and Ani Path of Symbols Bracelet","Alex and Ani","Alex and Ani公式","https://www.alexandani.com/","https://www.alexandani.com/","support@alexandani.com","★★★★☆","リサイクルメタル製・充電式のEco Bangles開運ブレスレット。意味を持つシンボルがデザインされたスタックブレスレット。スピリチュアル市場向け。"],
  [2454,"ファッション・アクセサリー","Pandora Moments Silver Charm Bracelet","Pandora","Pandora公式","https://www.pandora.net/moments-bracelet","https://www.pandora.net/","support@pandora.net","★★★★★","世界最大のジュエリーブランドのチャームブレスレット。記念日・誕生日のギフト向けに個性的なチャームを追加していくカスタムジュエリー。"],
  [2455,"ファッション・アクセサリー","Gola Classics Harrier Trainer Sneakers","Gola","Amazon US","https://www.golashoes.co.uk/","https://www.golashoes.co.uk/","info@golashoes.co.uk","★★★★☆","1905年英国創業のスポーツブランドのビンテージインスパイアスニーカー。レトロファッションブームで再評価中のブランド。日本向け輸入余地大。"],
  [2456,"ファッション・アクセサリー","Herschel Supply Co. Retreat Backpack","Herschel Supply","Amazon US","https://herschel.com/retreat","https://herschel.com/","support@herschel.com","★★★★★","シグネチャーストライプ裏地・ラップトップポケット搭載の定番ファッションバックパック。カナダ発のレトロスタイルブランドで日本でも人気。"],
  [2457,"ファッション・アクセサリー","Le Specs Mama Rag Cat-Eye Sunglasses","Le Specs","Le Specs公式","https://lespecs.com/","https://lespecs.com/","support@lespecs.com","★★★★☆","豪州発・セレブ愛用・高品質レンズのキャットアイサングラス。コスパ最高のラグジュアリースタイルのサングラスブランド。"],
  [2458,"ファッション・アクセサリー","Oliver Peoples Cary Grant Acetate Sunglasses","Oliver Peoples","Oliver Peoples公式","https://www.oliverpeoples.com/","https://www.oliverpeoples.com/","support@oliverpeoples.com","★★★★★","ハリウッドスター御用達のロサンゼルス発高級アイウェアブランド。日本の日野実業との提携でアジア展開が進む。日本市場との縁も深い。"],
  [2459,"ファッション・アクセサリー","Cuyana Classic Slim Leather Wallet","Cuyana","Cuyana公式","https://www.cuyana.com/slim-wallet","https://www.cuyana.com/","support@cuyana.com","★★★★★","イタリアレザー・超スリム・4カードスロットのミニマリスト本革財布。キャッシュレス化に合わせたコンパクト財布として日本向け需要大。"],
  [2460,"ファッション・アクセサリー","Mejuri Bold Chain Necklace Gold","Mejuri","Mejuri公式","https://mejuri.com/bold-chain","https://mejuri.com/","support@mejuri.com","★★★★★","14金・直営店・オンライン完結のダイレクトtoコンシューマー高級ジュエリー。中間業者を省いてラグジュアリー品質をリーズナブルな価格で提供。"],
  [2461,"ファッション・アクセサリー","Orseund Iris Ring with Gem","Orseund Iris","Orseund公式","https://orseundiris.com/","https://orseundiris.com/","info@orseundiris.com","★★★★☆","ニューヨーク発・ユニークな石・アート感覚のハンドメイドジュエリーリング。インディーズ・アーティスト系のブランドとして日本のファン向け。"],
  [2462,"ファッション・アクセサリー","Quay Australia High Key Sunglasses","Quay Australia","Quay公式","https://www.quay.com/high-key","https://www.quay.com/","support@quay.com","★★★★☆","インフルエンサー・セレブコラボで急成長中のオーストラリア発ファッションサングラス。30ドル台のプチプラ価格でトレンドスタイルを実現。"],
  [2463,"ファッション・アクセサリー","Stoney Clover Lane Patch Pouch Bag","Stoney Clover Lane","Stoney Clover公式","https://stoneycloverlane.com/","https://stoneycloverlane.com/","support@stoneycloverlane.com","★★★★☆","カスタムパッチ・パーソナライズ・ポップカラーのトラベルポーチバッグ。SNSで大人気の米国発カスタマイズバッグブランド。"],

  // ===== クリーニング・収納・整理 (2464-2483) =====
  [2464,"クリーニング・収納・整理","Roborock Q Revo MaxV Robot Vacuum","Roborock","Roborock公式","https://www.roborock.com/q-revo-maxv","https://www.roborock.com/","support@roborock.com","★★★★★","FlexiArm障害物回避・RockSteadyカメラ・自動モップ洗浄ステーション搭載のロボット掃除機上位モデル。ペットや散らかった部屋でも安心。"],
  [2465,"クリーニング・収納・整理","ECOVACS DEEBOT T30S Combo Robot","ECOVACS","ECOVACS公式","https://www.ecovacs.com/t30s","https://www.ecovacs.com/","support@ecovacs.com","★★★★★","6000Pa吸引力・TrueDetect 3D・ホットウォーターモップ洗浄のオールインワンロボット掃除機。高吸引力と水拭き機能を両立したフラッグシップ。"],
  [2466,"クリーニング・収納・整理","Miele Complete C3 Cat & Dog Vacuum","Miele","Miele公式","https://www.miele.com/complete-c3-cat-dog","https://www.miele.com/","https://www.miele.com/en/com/contact.htm","★★★★★","ペット用ターボブラシ・HEPAフィルター・2200W吸引力のペット飼育者向けキャニスター掃除機。ドイツ品質の20年耐久設計。"],
  [2467,"クリーニング・収納・整理","Hoover ONEPWR Blade+ Cordless Vacuum","Hoover","Amazon US","https://www.hoover.com/onepwr-blade-plus","https://www.hoover.com/","support@hoover.com","★★★★★","交換式バッテリー・セルフ立て掛け・リバーシブルブラシローラーの次世代コードレス掃除機。OPNEPWRバッテリーで複数Hoover製品と共有可能。"],
  [2468,"クリーニング・収納・整理","Bissell Pet Hair Eraser Turbo Plus Vacuum","Bissell","Amazon US","https://www.bissell.com/pet-hair-eraser","https://www.bissell.com/","support@bissell.com","★★★★★","専用タービンエルゴーノミックハンドル・ペット毛特化フィルターのハンディ掃除機。ソファ・ベッド・カーペットのペット毛除去に特化した製品。"],
  [2469,"クリーニング・収納・整理","Rubbermaid Commercial FG354100 Mop Bucket","Rubbermaid","Amazon US","https://www.rubbermaid.com/","https://www.rubbermaid.com/","support@rubbermaid.com","★★★★☆","商業グレードの耐久性を持つワンガロンモップバケツセット。プロ清掃業者御用達ブランドが家庭向けに提供する高品質製品。"],
  [2470,"クリーニング・収納・整理","Bona Hardwood Floor Cleaner Spray","Bona","Amazon US","https://www.bona.com/hardwood-floor-cleaner","https://www.bona.com/","support@bona.com","★★★★★","Greenguard認定・無残留・速乾のフローリング専用クリーナー。スウェーデン発のフローリングケアブランドとして世界で高評価。"],
  [2471,"クリーニング・収納・整理","Bar Keepers Friend Powder Cleanser","Bar Keepers Friend","Amazon US","https://barkeepersfriend.com/","https://barkeepersfriend.com/","support@barkeepersfriend.com","★★★★★","シュウ酸配合でサビ・水垢・焦げ付きを強力除去するパウダークレンザー。プロの清掃業者・料理人が愛用する100年以上のロングセラー。"],
  [2472,"クリーニング・収納・整理","Libman Freedom Spray Mop","Libman","Amazon US","https://www.libman.com/freedom-spray-mop","https://www.libman.com/","support@libman.com","★★★★☆","洗剤内蔵スプレー・マイクロファイバーパッド交換式のスプレーモップ。水と洗剤を自動噴霧しながらフローリングを効率的に清掃。"],
  [2473,"クリーニング・収納・整理","Swiffer PowerMop Multi-Surface Mopping Kit","Swiffer","Amazon US","https://www.swiffer.com/powermop","https://www.swiffer.com/","https://www.pgeveryday.com/contact-us","★★★★★","電動スプレー・マイクロファイバーパッド・洗浄液のスマートモップシステム。洗浄液の自動噴霧でフローリング・タイル清掃を大幅に効率化。"],
  [2474,"クリーニング・収納・整理","Hoover MAXLife Elite Swivel Upright Vacuum","Hoover","Amazon US","https://www.hoover.com/maxlife-elite","https://www.hoover.com/","support@hoover.com","★★★★★","360度スイベルヘッド・HEPA・5年フィルター保証のアップライト掃除機。大型住宅・カーペットの多い欧米向けに最適な吸引力とフレキシビリティ。"],
  [2475,"クリーニング・収納・整理","SpaceAid Bamboo Drawer Divider Set","SpaceAid","Amazon US","https://spaceaid.com/","https://spaceaid.com/","support@spaceaid.com","★★★★★","竹製・幅調整可能な引き出し仕切りセット。キッチン・オフィス・寝室の引き出し整理に最適。断捨離・収納整理需要の高い日本向け定番製品。"],
  [2476,"クリーニング・収納・整理","mDesign Plastic Pantry Storage Bin Set","mDesign","Amazon US","https://www.mdesignhome.com/","https://www.mdesignhome.com/","support@mdesignhome.com","★★★★☆","透明・スタッキング可能・ラベルスロット付きのパントリー収納ケースセット。冷蔵庫・棚の食品整理に最適な定番収納ブランドの製品。"],
  [2477,"クリーニング・収納・整理","Windex Electronics Wipes for Screen Cleaning","Windex","Amazon US","https://www.windex.com/electronics-wipes","https://www.windex.com/","https://www.sc johnson.com/contact","★★★★☆","スマホ・タブレット・PC画面専用のアルコールフリー除菌クリーニングワイプ。電子機器を傷つけずに清潔に保つ専用クリーナー。"],
  [2478,"クリーニング・収納・整理","IRIS USA WeatherPro Storage Box 82Qt","IRIS USA","Amazon US","https://www.irisusainc.com/","https://www.irisusainc.com/","support@irisusa.com","★★★★★","アイリスオーヤマの北米向け防水収納ボックス。日本のアイリスオーヤマが米国で大成功した逆輸入事例として掲載価値が高い製品。"],
  [2479,"クリーニング・収納・整理","ClosetMaid Wire Shelving System","ClosetMaid","Amazon US","https://www.closetmaid.com/","https://www.closetmaid.com/","support@closetmaid.com","★★★★☆","DIY設置・カスタマイズ可能なワイヤー棚収納システム。クローゼット・パントリー・ガレージの整理に最適な米国定番収納ブランド。"],
  [2480,"クリーニング・収納・整理","Honey-Can-Do Bamboo Shelf Closet Organizer","Honey-Can-Do","Amazon US","https://www.honeycando.com/","https://www.honeycando.com/","support@honeycando.com","★★★★☆","環境配慮型竹製の棚付きクローゼットオーガナイザー。靴・バッグ・衣類の整理整頓に最適。エコ素材で日本の環境意識に合致。"],
  [2481,"クリーニング・収納・整理","Umite Chef Magnetic Spice Jars Set","Umite Chef","Amazon US","https://umitechef.com/","https://umitechef.com/","support@umitechef.com","★★★★★","磁石でどこにでも貼り付けられる透明スパイスジャーセット。冷蔵庫サイド・金属面に貼るだけでキッチンスペースを有効活用。収納整理向け。"],
  [2482,"クリーニング・収納・整理","OXO Good Grips Large Salad Spinner","OXO","Amazon US","https://www.oxo.com/large-salad-spinner","https://www.oxo.com/","support@oxo.com","★★★★★","ソフトポンプ・ブレーキボタン・食洗機対応の大型サラダスピナー。野菜の水切りを完全自動化し料理の下準備を効率化する定番キッチン用品。"],
  [2483,"クリーニング・収納・整理","Simplehuman Tension Shower Caddy","Simplehuman","Amazon US","https://www.simplehuman.com/shower-caddy","https://www.simplehuman.com/","support@simplehuman.com","★★★★★","工事不要・突っ張り棒式・さび止めステンレス製のシャワーキャディ。シャンプー・ボディソープを整理する機能的なバスルーム収納製品。"],

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
