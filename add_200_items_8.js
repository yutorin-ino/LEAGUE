/**
 * add_200_items_8.js — No.2084〜2283 (200件追加)
 * node add_200_items_8.js
 */

const XLSX = require('xlsx');
const path = require('path');

const INPUT_FILE  = path.join(__dirname, '海外便利グッズリスト_日本未上陸2105件_評価付.xlsx');
const OUTPUT_FILE = path.join(__dirname, '海外便利グッズリスト_日本未上陸2305件_評価付.xlsx');
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

  // ===== キッチン・調理器具 (2084-2103) =====
  [2084,"キッチン・調理器具","Alchema Home Cider & Wine Maker","Alchema","Alchema公式","https://www.alchema.com/","https://www.alchema.com/","hello@alchema.com","★★★★★","AIレシピ・自動発酵管理でリンゴ酒・ワイン・クラフトビールを家庭で醸造するスマート醸造機。CES 2026出展。自家醸造ブームの日本向け。"],
  [2085,"キッチン・調理器具","Combustion Inc. Predictive Thermometer","Combustion Inc.","Combustion公式","https://combustion.inc/","https://combustion.inc/","support@combustion.inc","★★★★★","8つのセンサーとAI予測エンジンで肉の内部温度・残り時間を予測するワイヤレス料理用温度計。プロ並みの焼き加減を家庭で実現。"],
  [2086,"キッチン・調理器具","Zojirushi Rizo Induction Rice Cooker NW-QAE","Zojirushi","Amazon US","https://www.zojirushi.com/app/product/nwqae","https://www.zojirushi.com/","https://www.zojirushi.com/app/contact","★★★★★","象印の北米向けIH圧力炊飯器最新モデル。米国・カナダ市場で圧倒的人気。日本ブランドとして海外展開の逆輸入事例として紹介価値大。"],
  [2087,"キッチン・調理器具","Olly Smart Modular Fridge","Olly","Kickstarter","https://heyolly.com/","https://heyolly.com/","hello@heyolly.com","★★★★★","モジュール式で容量を自在に拡張できるスマート冷蔵庫。省スペースで日本の賃貸住宅に設置しやすく、需要の拡大が見込まれる。"],
  [2088,"キッチン・調理器具","Breville Smart Oven Air Fryer Pro","Breville","Amazon US","https://www.breville.com/us/en/products/smart-oven/bov900bss.html","https://www.breville.com/","support@breville.com","★★★★★","13種の調理モード・コンベクション・エアフライ機能搭載のスマートオーブントースター。プロ品質の調理を家庭で実現する定番ブランド。"],
  [2089,"キッチン・調理器具","Mymee Food Sensitivity App & Test Kit","Mymee","Mymee公式","https://www.mymee.com/","https://www.mymee.com/","hello@mymee.com","★★★★☆","AIが食事・症状の日記を解析して食物感受性のトリガーを特定するパーソナライズ健康管理サービス。フードアレルギー意識の高まる日本向け。"],
  [2090,"キッチン・調理器具","Kenwood Chef Titanium Kitchen Machine","Kenwood","Amazon US","https://www.kenwood.com/","https://www.kenwood.com/","support@kenwood.com","★★★★★","6.7Lボウル・800W・20種のアタッチメントに対応するプレミアムキッチンマシン。英国老舗ブランドの業務用品質。製パン愛好家向け。"],
  [2091,"キッチン・調理器具","Vitamix Explorian E310 Blender","Vitamix","Amazon US","https://www.vitamix.com/","https://www.vitamix.com/","support@vitamix.com","★★★★★","スムージー・スープ・ナッツバターまで対応する業務用クラスの高性能ブレンダー。Vitamixはプロ料理人にも愛用される最高峰ブランド。"],
  [2092,"キッチン・調理器具","Ankarsrum Original Kitchen Machine","Ankarsrum","Ankarsrum公式","https://ankarsrum.com/","https://ankarsrum.com/","info@ankarsrum.com","★★★★★","スウェーデン製・1940年代から続く最高峰キッチンマシン。ボウル回転式の独自設計で生地をダメージなくこねる。パン職人・パティシエ向け。"],
  [2093,"キッチン・調理器具","Aarke Carbonator Pro Sparkling Water Maker","Aarke","Aarke公式","https://aarke.com/","https://aarke.com/","support@aarke.com","★★★★★","スウェーデンデザインのステンレス製プレミアム炭酸水メーカー。SodaStreamと互換。インテリアとしても映えるスカンジナビアデザイン。"],
  [2094,"キッチン・調理器具","Supercook AI Recipe App with Pantry Scan","Supercook","Supercook公式","https://www.supercook.com/","https://www.supercook.com/","hello@supercook.com","★★★★☆","冷蔵庫・パントリーにある食材をスキャンして作れるレシピをAIが提案するアプリ。食材の無駄を減らす食品ロス削減ツール。日本向け。"],
  [2095,"キッチン・調理器具","Balmuda The Toaster Pro","Balmuda","Balmuda公式","https://www.balmuda.com/","https://www.balmuda.com/","https://www.balmuda.com/pages/contact","★★★★★","スチーム加熱で外はカリッと中はふっくら焼き上げる日本発のプレミアムトースター。海外でも人気急上昇中の日本ブランド逆輸入事例。"],
  [2096,"キッチン・調理器具","Yummly Smart Thermometer","Yummly","Yummly公式","https://www.yummly.com/smart-thermometer","https://www.yummly.com/","support@yummly.com","★★★★☆","アプリ連動でレシピと連携するスマート調理用温度計。Yummlyのレシピアプリと組み合わせて完璧な調理ガイダンスを提供。"],
  [2097,"キッチン・調理器具","Modernist Pantry Sodium Alginate Molecular Kit","Modernist Pantry","Modernist Pantry公式","https://modernistpantry.com/","https://modernistpantry.com/","hello@modernistpantry.com","★★★★☆","分子調理学の食材キット。球形化・ゲル化・泡立てなどの革新的調理技法を家庭で体験。料理好き・フードクリエイター向け。"],
  [2098,"キッチン・調理器具","Instant Pot Duo Crisp 11-in-1 Air Fryer","Instant Brands","Amazon US","https://www.instantpot.com/","https://www.instantpot.com/","support@instantpot.com","★★★★★","圧力調理・エアフライ・蒸し・焼き・炒めなど11機能を一台で実現するマルチクッカー。世界的人気ブランドが日本未展開で需要大。"],
  [2099,"キッチン・調理器具","Traeger Ironwood 650 Wood Pellet Grill","Traeger","Traeger公式","https://www.traegergrills.com/","https://www.traegergrills.com/","support@traeger.com","★★★★★","Wi-Fi接続・アプリ操作・自動温度管理のスマートウッドペレットグリル。本格BBQ文化が根付く前に日本市場を先行開拓できる製品。"],
  [2100,"キッチン・調理器具","Oxo Brew 9-Cup Coffee Maker","OXO","Amazon US","https://www.oxo.com/oxo-brew-9-cup-coffee-maker.html","https://www.oxo.com/","support@oxo.com","★★★★★","SCAA認定・精密温度制御・レインシャワーヘッドのスペシャルティコーヒー向け全自動コーヒーメーカー。コーヒー文化の深化する日本向け。"],
  [2101,"キッチン・調理器具","Fellow Stagg EKG Electric Kettle","Fellow","Fellow公式","https://fellowproducts.com/products/stagg-ekg-electric-pour-over-kettle","https://fellowproducts.com/","support@fellowproducts.com","★★★★★","1度単位の温度設定・ストップウォッチ内蔵・グースネックのバリスタ向けスマート電気ケトル。スペシャルティコーヒー愛好家向け定番。"],
  [2102,"キッチン・調理器具","Imperia Pasta Machine Electric Motorized","Imperia","Amazon US","https://www.imperiausa.com/","https://www.imperiausa.com/","info@imperiausa.com","★★★★☆","イタリア製・電動モーター付きの本格パスタマシン。ラーメン・うどん・パスタの自家製麺づくりに。料理趣味の日本市場向け製品。"],
  [2103,"キッチン・調理器具","Cali Bamboo EarthHero Bamboo Cutting Board Set","Cali Bamboo","Amazon US","https://www.calibamboo.com/","https://www.calibamboo.com/","support@calibamboo.com","★★★☆☆","抗菌・耐久性の高い竹製まな板セット。竹素材の持続可能性と日本の竹文化との親和性が高い。エコ製品として日本市場向け。"],

  // ===== スマートホーム・インテリア・照明 (2104-2123) =====
  [2104,"スマートホーム・インテリア・照明","Bosch Smart Home Controller II","Bosch Smart Home","Bosch公式","https://www.bosch-smarthome.com/","https://www.bosch-smarthome.com/","support@bosch-smarthome.com","★★★★★","Bosch製スマートホームハブ。照明・暖房・セキュリティを一元管理。Matter対応でエコシステムを問わない接続性。欧州品質で信頼性が高い。"],
  [2105,"スマートホーム・インテリア・照明","Ring Floodlight Cam Wired Pro","Ring","Amazon US","https://ring.com/products/floodlight-cam-wired-pro","https://ring.com/","support@ring.com","★★★★★","3D動体検知・1080p HDR・AI警告音搭載の有線フラッドライトカメラ。戸建て・マンション共用部のセキュリティ強化に最適。"],
  [2106,"スマートホーム・インテリア・照明","Nest Thermostat 2024","Google Nest","Google公式","https://store.google.com/us/product/nest_thermostat","https://store.google.com/","https://support.google.com/googlenest","★★★★★","機械学習で生活パターンを自動学習・省エネ最適化するGoogleのスマートサーモスタット。電気代節約ニーズが高い日本市場に合致。"],
  [2107,"スマートホーム・インテリア・照明","Schlage Encode Plus Smart Lock","Schlage","Amazon US","https://www.schlage.com/en/home/smart-locks/encode-plus.html","https://www.schlage.com/","support@schlage.com","★★★★★","Apple HomeKey・Matter対応のスマートデッドボルト。スマホ・Apple Watchでドア解錠。鍵を持ち歩かないスマートロック。"],
  [2108,"スマートホーム・インテリア・照明","Govee Neon Rope Light 2","Govee","Govee公式","https://www.govee.com/","https://www.govee.com/","marketing@govee.com","★★★★★","セグメント別色設定・音楽同期・アプリ操作のスマートネオンロープライト。TikTok・YouTube動画背景として人気沸騰中。Z世代向け。"],
  [2109,"スマートホーム・インテリア・照明","August Wi-Fi Smart Lock Pro","August Home","August公式","https://august.com/products/august-wifi-smart-lock","https://august.com/","support@august.com","★★★★★","既存のデッドボルトに後付け可能なWi-Fiスマートロック。工事不要で賃貸住宅でも設置可能。日本の賃貸市場向けに特に有望。"],
  [2110,"スマートホーム・インテリア・照明","Lutron Caseta Pro Smart Lighting Starter Kit","Lutron","Lutron公式","https://www.casetawireless.com/","https://www.casetawireless.com/","support@lutron.com","★★★★★","北米プロ仕様のスマート調光スイッチ。工事業者向けプロ対応の信頼性が高い照明制御システム。Matter対応で将来性高い。"],
  [2111,"スマートホーム・インテリア・照明","Hatch Restore 2 Smart Sleep Light","Hatch","Hatch公式","https://www.hatch.co/restore","https://www.hatch.co/","support@hatch.co","★★★★★","AIで起床・就寝ルーティンを個別最適化するスマート目覚まし・睡眠ライト。光・音・瞑想・読書モード搭載。睡眠悩みが多い日本向け。"],
  [2112,"スマートホーム・インテリア・照明","Elgato Key Light Air","Elgato","Amazon US","https://www.elgato.com/us/en/p/key-light-air","https://www.elgato.com/","support@elgato.com","★★★★★","Wi-Fi制御・2800ルーメン・色温度調整のコンテンツクリエイター向けスマートキーライト。在宅ビデオ会議・YouTube配信向け需要急増中。"],
  [2113,"スマートホーム・インテリア・照明","SimpliSafe Home Security System","SimpliSafe","SimpliSafe公式","https://simplisafe.com/","https://simplisafe.com/","support@simplisafe.com","★★★★★","DIY設置・月額プロ監視オプション・AIカメラ搭載のオールインワンホームセキュリティ。工事不要で導入しやすく賃貸でも使える。"],
  [2114,"スマートホーム・インテリア・照明","Signify WiZ Color Smart Bulb","Signify (WiZ)","Amazon US","https://www.wizconnected.com/","https://www.wizconnected.com/","support@wizconnected.com","★★★★☆","ハブ不要・Wi-Fi直接接続・1600万色のコスパ最高スマート電球。フィリップスの姉妹ブランドで信頼性が高く、日本向け入門製品。"],
  [2115,"スマートホーム・インテリア・照明","Meross Smart Garage Door Opener","Meross","Amazon US","https://www.meross.com/","https://www.meross.com/","support@meross.com","★★★★☆","既存のガレージドアに後付けできるスマートガレージオープナー。Apple HomeKit・Alexa・Google対応。戸建て住宅のガレージ管理向け。"],
  [2116,"スマートホーム・インテリア・照明","Lexi Light AI Desk Lamp","Lexi Light","Kickstarter","https://lexilight.com/","https://lexilight.com/","info@lexilight.com","★★★★★","AIが目の疲れを検知してリアルタイムに色温度・明るさを自動調整するスマートデスクランプ。在宅勤務・受験生向け。眼精疲労対策に最適。"],
  [2117,"スマートホーム・インテリア・照明","Airthings Wave Plus Indoor Air Quality Monitor","Airthings","Airthings公式","https://www.airthings.com/","https://www.airthings.com/","support@airthings.com","★★★★★","ラドン・CO2・VOC・湿度・気圧をリアルタイムモニタリングするスマート空気質センサー。シックハウス問題・アレルギー対策に必要な製品。"],
  [2118,"スマートホーム・インテリア・照明","Vedum Smart Bath Mirror with Speaker","Vedum","Vedum公式","https://vedum.com/","https://vedum.com/","info@vedum.com","★★★★☆","スピーカー・LED照明・防曇・Bluetoothのスマートバスルームミラー。スカンジナビアデザインで高品質。浴室をスマート化する製品。"],
  [2119,"スマートホーム・インテリア・照明","Philips SmartSleep Wake-Up Light HF3500","Philips","Amazon US","https://www.philips.com/c-p/HF3500_60/smartsleep-wake-up-light","https://www.philips.com/","support@philips.com","★★★★★","日の出を模した光で自然な目覚めを促すスマート光目覚まし。睡眠の質を重視する日本のユーザー向けウェルネス製品。"],
  [2120,"スマートホーム・インテリア・照明","Furrion Vision S Wireless RV Camera","Furrion","Amazon US","https://www.furrion.com/","https://www.furrion.com/","support@furrion.com","★★★★☆","360度視野・ナイトビジョン・IP67防水のワイヤレスバックアップカメラ。キャンピングカー・車中泊ブームの日本向けカーカメラ。"],
  [2121,"スマートホーム・インテリア・照明","IKEA DIRIGERA Smart Home Hub","IKEA","IKEA公式","https://www.ikea.com/","https://www.ikea.com/","https://www.ikea.com/us/en/customer-service/","★★★★☆","IKEAの全スマートホーム製品をアプリで一元管理するMatter対応ハブ。IKEAブランドの親しみやすさで日本のスマートホーム入門者向け。"],
  [2122,"スマートホーム・インテリア・照明","Tuya Smart Life Platform Devices","Tuya","Tuya公式","https://www.tuya.com/","https://www.tuya.com/","info@tuya.com","★★★★☆","世界最大のIoTプラットフォーム上のスマートホームデバイス群。コスパが高く種類が豊富でスマートホーム構築の選択肢として有用。"],
  [2123,"スマートホーム・インテリア・照明","SwitchBot Hub 2 Matter Smart Bridge","SwitchBot","Amazon US","https://www.switch-bot.com/","https://www.switch-bot.com/","support@switch-bot.com","★★★★★","赤外線リモコン家電をスマートホーム化するMatter対応ブリッジ。既存のテレビ・エアコン・照明をアプリ操作に変換。日本市場に最適。"],

  // ===== ウェアラブル・ヘルス・フィットネス (2124-2143) =====
  [2124,"ウェアラブル・ヘルス・フィットネス","Movano Evie Ring Smart Ring for Women","Movano Health","Movano公式","https://www.movanohealth.com/","https://www.movanohealth.com/","hello@movanohealth.com","★★★★★","FDA承認・女性向け設計のスマートリング。月経周期・血中酸素・心拍・睡眠を精密トラッキング。女性の健康管理需要が高い日本市場向け。"],
  [2125,"ウェアラブル・ヘルス・フィットネス","Amazfit Helio Ring","Amazfit","Amazfit公式","https://www.amazfit.com/helio-ring","https://www.amazfit.com/","support@amazfit.com","★★★★★","スマートウォッチとセットで使う初のスポーツフォーカス型スマートリング。Oura Ringより安価でトレーニングデータと統合管理が可能。"],
  [2126,"ウェアラブル・ヘルス・フィットネス","Dexcom G7 Continuous Glucose Monitor","Dexcom","Dexcom公式","https://www.dexcom.com/g7","https://www.dexcom.com/","support@dexcom.com","★★★★★","世界最小クラスのウェアラブル持続血糖モニター。スマホ・Apple Watch連動。糖尿病管理・健康志向者向けに日本でも需要が高まっている。"],
  [2127,"ウェアラブル・ヘルス・フィットネス","Lumen Metabolism Tracker","Lumen","Lumen公式","https://www.lumen.me/","https://www.lumen.me/","support@lumen.me","★★★★★","息を吹きかけるだけで代謝状態を測定するウェアラブル代謝トラッカー。脂肪燃焼vs炭水化物燃焼をリアルタイム判定。ダイエット・健康管理向け。"],
  [2128,"ウェアラブル・ヘルス・フィットネス","Whoop 4.0 Fitness Wearable","WHOOP","WHOOP公式","https://www.whoop.com/","https://www.whoop.com/","support@whoop.com","★★★★★","回復・ストレイン・睡眠を24時間トラッキングするコーチング型フィットネスウェアラブル。NBA・NFL選手も愛用。アスリート向け定番ブランド。"],
  [2129,"ウェアラブル・ヘルス・フィットネス","Biostrap Kick AI Shoe Clip","Biostrap","Biostrap公式","https://biostrap.com/kick/","https://biostrap.com/","support@biostrap.com","★★★★☆","靴に取り付けてランニングフォーム・歩行パターン・バランスをAI分析するウェアラブルセンサー。怪我予防・リハビリ向け。"],
  [2130,"ウェアラブル・ヘルス・フィットネス","Zephyr BioHarness Wearable Physiology Monitor","Zephyr Technology","Zephyr公式","https://zephyranywhere.com/","https://zephyranywhere.com/","info@zephyranywhere.com","★★★★★","心拍・呼吸・体温・加速度を医療グレードで測定するスポーツ・軍事向けバイオメトリクスセンサー。アスリート・消防士向け。"],
  [2131,"ウェアラブル・ヘルス・フィットネス","Hexoskin Smart Shirt","Carre Technologies","Hexoskin公式","https://www.hexoskin.com/","https://www.hexoskin.com/","support@hexoskin.com","★★★★★","シャツ型ウェアラブルで心拍・呼吸・活動量を連続測定。医療・スポーツ科学・航空宇宙分野で活用されるクリニカルグレード製品。"],
  [2132,"ウェアラブル・ヘルス・フィットネス","Levels Health Metabolic Fitness CGM","Levels Health","Levels公式","https://www.levelshealth.com/","https://www.levelshealth.com/","support@levelshealth.com","★★★★★","持続血糖モニター＋AIコーチングで食事・運動・睡眠が血糖値に与える影響をリアルタイム分析。代謝的健康の最適化を目指すサービス。"],
  [2133,"ウェアラブル・ヘルス・フィットネス","Komodo Technologies AIO Sleeve","Komodo Technologies","Komodo公式","https://www.komodotechnologies.com/","https://www.komodotechnologies.com/","info@komodotechnologies.com","★★★★☆","腕に装着するコンパクトなフィットネストラッカー。スマートウォッチ不要で手首を自由に使いながら心拍・活動量をトラッキング。"],
  [2134,"ウェアラブル・ヘルス・フィットネス","Tempdrop Fertility & Sleep Tracker","Tempdrop","Tempdrop公式","https://www.tempdrop.com/","https://www.tempdrop.com/","support@tempdrop.com","★★★★★","腋に装着して睡眠中に基礎体温を連続測定する女性向け妊活・避妊サポートウェアラブル。AI学習で精度が向上する独自アルゴリズム搭載。"],
  [2135,"ウェアラブル・ヘルス・フィットネス","Muse 2 Meditation Headband","InteraXon","Muse公式","https://choosemuse.com/muse-2/","https://choosemuse.com/","hello@choosemuse.com","★★★★★","EEG・心拍・呼吸・体の動きで瞑想状態をリアルタイムフィードバックするスマートヘッドバンド。マインドフルネス実践者向け定番製品。"],
  [2136,"ウェアラブル・ヘルス・フィットネス","Garmin HRM-Pro Plus Heart Rate Monitor","Garmin","Garmin公式","https://www.garmin.com/en-US/p/770739","https://www.garmin.com/","support@garmin.com","★★★★★","走行ダイナミクス・心拍・ランニングパワーを測定する胸部装着型HRMセンサー。Garminデバイスと完全連携でプロ品質のランデータ取得。"],
  [2137,"ウェアラブル・ヘルス・フィットネス","Neurosity Crown EEG Headset","Neurosity","Neurosity公式","https://neurosity.co/","https://neurosity.co/","hello@neurosity.co","★★★★★","開発者向けEEG脳波ヘッドセット。集中度・瞑想・精神状態をリアルタイム計測しアプリ開発・神経フィードバックに活用。",""],
  [2138,"ウェアラブル・ヘルス・フィットネス","Embr Wave 2 Thermal Wristband","Embr Labs","Embr公式","https://embrlabs.com/","https://embrlabs.com/","support@embrlabs.com","★★★★★","局所温度刺激で体感温度をコントロールするウェアラブルリストバンド。更年期ホットフラッシュ・冷え性対策として日本女性向けに有望。"],
  [2139,"ウェアラブル・ヘルス・フィットネス","Fitbit Charge 6 Fitness Tracker","Google Fitbit","Amazon US","https://www.fitbit.com/charge6","https://www.fitbit.com/","support@fitbit.com","★★★★★","ECG・血中酸素・YouTube Music・Google Maps連動の最新フィットビット。Google統合で機能が大幅向上。健康管理入門向けウェアラブル。"],
  [2140,"ウェアラブル・ヘルス・フィットネス","Halo Rise Sleep Tracker by Amazon","Amazon","Amazon公式","https://www.amazon.com/halo-rise/","https://www.amazon.com/","https://www.amazon.com/contact-us","★★★★☆","接触不要でベッドサイドから睡眠を追跡・自然光で起床をサポートするスマート睡眠トラッカー。装着不要で負担なく睡眠データ取得。"],
  [2141,"ウェアラブル・ヘルス・フィットネス","Apollo Neuro Wearable Stress Relief","Apollo Neuroscience","Apollo公式","https://apolloneuro.com/","https://apolloneuro.com/","support@apolloneuro.com","★★★★★","振動によって自律神経を調整しストレス軽減・集中力向上・睡眠改善を実現するウェアラブル。ストレス社会の日本に特に有望な製品。"],
  [2142,"ウェアラブル・ヘルス・フィットネス","Polar H10 Heart Rate Sensor","Polar","Amazon US","https://www.polar.com/us-en/sensors/h10-heart-rate-sensor","https://www.polar.com/","support@polar.com","★★★★★","業界最高精度クラスのBluetooth・ANT+胸部HRMセンサー。スポーツ科学・リハビリ・コーチングでの活用に最適なプロ向け製品。"],
  [2143,"ウェアラブル・ヘルス・フィットネス","Sensate Vagus Nerve Stimulator","BioSelf Technology","Sensate公式","https://www.getsensate.com/","https://www.getsensate.com/","support@getsensate.com","★★★★★","迷走神経に振動刺激を与えてストレス・不安を緩和するウェアラブルデバイス。10分の使用でリラクゼーション効果。現代日本のストレス対策向け。"],

  // ===== アウトドア・スポーツ・旅行 (2144-2163) =====
  [2144,"アウトドア・スポーツ・旅行","Garmin Fenix 8 Solar Multisport Watch","Garmin","Garmin公式","https://www.garmin.com/fenix8","https://www.garmin.com/","support@garmin.com","★★★★★","太陽光充電・AMOLED・マルチバンドGPS搭載のプレミアムアウトドア向けスマートウォッチ。登山・トレイルラン・トライアスロン対応。"],
  [2145,"アウトドア・スポーツ・旅行","BioLite BaseCamp Campfire Stove","BioLite","BioLite公式","https://www.bioliteenergy.com/products/basecamp-fire-pit","https://www.bioliteenergy.com/","support@bioliteenergy.com","★★★★★","発電機能付き・スマホ充電・電動ファン内蔵の薪焚きキャンプファイヤーストーブ。煙が少なくアウトドア調理・焚き火の両立が可能。"],
  [2146,"アウトドア・スポーツ・旅行","Helinox Chair Zero Ultralight Camp Chair","Helinox","Amazon US","https://helinox.com/","https://helinox.com/","info@helinox.com","★★★★★","韓国発・490gの超軽量コンパクトキャンプチェア。登山・ハイキング・ソロキャンプ向けに世界で人気急上昇中のブランド。"],
  [2147,"アウトドア・スポーツ・旅行","Away Bigger Carry-On Aluminum Suitcase","Away","Away公式","https://www.awaytravel.com/","https://www.awaytravel.com/","hello@awaytravel.com","★★★★★","USBポート内蔵・軽量・360度回転輪のプレミアムキャリーオンスーツケース。旅行ブランドのニュースタンダードとして世界的人気。"],
  [2148,"アウトドア・スポーツ・旅行","Nemo Tensor Insulated Sleeping Pad","Nemo Equipment","Nemo公式","https://www.nemoequipment.com/","https://www.nemoequipment.com/","support@nemoequipment.com","★★★★★","R値4.2・超軽量・コンパクト収納の高断熱インフレータブルスリーピングパッド。3シーズン対応で登山・バックパッキング向け。"],
  [2149,"アウトドア・スポーツ・旅行","Hyperlite Mountain Gear Dirigo 2 Tent","Hyperlite Mountain Gear","Hyperlite公式","https://www.hyperlitemountaingear.com/","https://www.hyperlitemountaingear.com/","info@hyperlitemountaingear.com","★★★★★","Dyneema素材・760g超軽量・4シーズン対応の2人用山岳テント。超軽量バックパッキング向けのUSAプレミアムブランド。"],
  [2150,"アウトドア・スポーツ・旅行","Garmin GPSMAP 67i Satellite Communicator","Garmin","Garmin公式","https://www.garmin.com/gpsmap67i","https://www.garmin.com/","support@garmin.com","★★★★★","マルチバンドGPS・衛星SOS・2ウェイメッセージ対応のハンディGPS。山岳遭難・海難での救助要請に対応。登山大国日本での必需品。"],
  [2151,"アウトドア・スポーツ・旅行","Sea to Summit Ultralight Dry Sack","Sea to Summit","Amazon US","https://www.seatosummit.com/","https://www.seatosummit.com/","support@seatosummit.com","★★★★☆","超軽量・完全防水のロールトップドライサック。カヤック・登山・釣りで装備を水濡れから保護。アウトドア収納の定番ブランド。"],
  [2152,"アウトドア・スポーツ・旅行","GoPro Hero 13 Black Action Camera","GoPro","GoPro公式","https://gopro.com/hero13-black","https://gopro.com/","support@gopro.com","★★★★★","5.3K60fps・HyperSmooth 7.0手ブレ補正・AI自動編集のアクションカメラ最新モデル。サーフィン・スキー・ダイビング撮影向け定番。"],
  [2153,"アウトドア・スポーツ・旅行","Garmin Descent Mk3i Dive Computer Watch","Garmin","Garmin公式","https://www.garmin.com/descent-mk3i","https://www.garmin.com/","support@garmin.com","★★★★★","AMOLED・ダイブコンピューター・スマートウォッチ機能を兼備するプレミアムダイブウォッチ。スキューバダイビング愛好者向け。"],
  [2154,"アウトドア・スポーツ・旅行","BioLite SolarHome 620 Off-Grid Kit","BioLite","BioLite公式","https://www.bioliteenergy.com/products/solarhome-620","https://www.bioliteenergy.com/","support@bioliteenergy.com","★★★★★","ソーラーパネル・バッテリー・LED照明・スマホ充電機能を含むオフグリッド生活キット。防災・キャンプ・発展途上国支援向け。"],
  [2155,"アウトドア・スポーツ・旅行","Lifeproof FRE iPhone Waterproof Case","Lifeproof","Amazon US","https://www.lifeproof.com/","https://www.lifeproof.com/","support@lifeproof.com","★★★★☆","IP68防水・耐衝撃・スクリーンプロテクター一体型のアクティビティ向けiPhoneケース。サーフィン・登山・アウトドア向け定番ブランド。"],
  [2156,"アウトドア・スポーツ・旅行","MSR PocketRocket 2 Mini Stove","MSR","Amazon US","https://www.msrgear.com/","https://www.msrgear.com/","support@msrgear.com","★★★★★","73gの超軽量・高火力ガスバーナーストーブ。世界中の登山家・バックパッカーが愛用する定番ブランド。ソロキャンプブームの日本に最適。"],
  [2157,"アウトドア・スポーツ・旅行","Sawyer Squeeze Water Filter System","Sawyer Products","Amazon US","https://www.sawyer.com/","https://www.sawyer.com/","support@sawyer.com","★★★★★","100万ガロン以上の浄水能力を持つ軽量水フィルター。川・池の水を安全に飲料化。登山・サバイバル・途上国支援にも活用可能。"],
  [2158,"アウトドア・スポーツ・旅行","Shimano Ultegra Di2 Electric Groupset","Shimano","Shimano公式","https://www.shimano.com/","https://www.shimano.com/","support@shimano.com","★★★★★","電動変速・自動トリミング・Wi-Fi設定対応の電動自転車コンポーネント。日本Shimanoの最上位ロードバイクコンポ。逆輸入事例として価値大。"],
  [2159,"アウトドア・スポーツ・旅行","Topeak SmartGauge D2X Bicycle Tire Gauge","Topeak","Amazon US","https://www.topeak.com/","https://www.topeak.com/","support@topeak.com","★★★★☆","デジタル表示・高精度のスマート自転車タイヤゲージ。グラベルバイク・ロードバイク向けに正確な空気圧管理が可能。自転車愛好家向け。"],
  [2160,"アウトドア・スポーツ・旅行","Lezyne GPS Super Pro Bike Computer","Lezyne","Lezyne公式","https://www.lezyne.com/","https://www.lezyne.com/","support@lezyne.com","★★★★★","Strava・Garmin Connect連動・マルチバンドGPS搭載の高精度サイクルコンピューター。ロードバイク・グラベル愛好家向けプレミアム製品。"],
  [2161,"アウトドア・スポーツ・旅行","LifeStraw Go Water Filter Bottle","LifeStraw","Amazon US","https://lifestraw.com/products/lifestraw-go","https://lifestraw.com/","support@lifestraw.com","★★★★★","ボトルに内蔵されたフィルターで飲む場所の水を浄化する携帯浄水ボトル。登山・旅行・防災向け必需品として日本市場でもニーズ大。"],
  [2162,"アウトドア・スポーツ・旅行","Nitecore NB10000 Carbon Fiber Power Bank","Nitecore","Amazon US","https://www.nitecore.com/","https://www.nitecore.com/","support@nitecore.com","★★★★☆","カーボンファイバーボディ・183gの超軽量モバイルバッテリー。登山・バックパッキング向けに重量を最優先した製品。"],
  [2163,"アウトドア・スポーツ・旅行","Garmin Edge 1050 Cycling Computer","Garmin","Garmin公式","https://www.garmin.com/edge1050","https://www.garmin.com/","support@garmin.com","★★★★★","4.3インチAMOLED・ClimbPro・AIコーチング機能搭載のプレミアムサイクルコンピューター。プロレーサーからサイクリスト愛好家まで対応。"],

  // ===== ペット用品 (2164-2183) =====
  [2164,"ペット用品","Petkit Purobot Ultra Self-Cleaning Litter Box","PETKIT","Petkit公式","https://petkit.com/","https://petkit.com/","support@petkit.com","★★★★★","AI健康モニタリング・自動クリーニング・消臭機能搭載のスマート猫トイレ。Litter-Robot 4の競合として低価格でプレミアム機能を提供。"],
  [2165,"ペット用品","ScoopFree Smart Self-Cleaning Litter Box","PetSafe","Amazon US","https://www.petsafe.net/ScoopFree","https://www.petsafe.net/","support@petsafe.net","★★★★☆","クリスタルリター使用・自動清掃・排泄回数モニタリングのスマート猫トイレ。メンテナンスが簡単で忙しい日本の飼育者向け。"],
  [2166,"ペット用品","Embark Dog DNA Test","Embark Veterinary","Embark公式","https://embarkvet.com/","https://embarkvet.com/","support@embarkvet.com","★★★★★","350以上の犬種・190以上の遺伝性疾患リスクを判定する犬向けDNA検査キット。愛犬の健康リスクを事前に把握できる人気サービス。"],
  [2167,"ペット用品","Fi Series 3 GPS Dog Collar","Fi","Fi公式","https://tryfi.com/","https://tryfi.com/","support@tryfi.com","★★★★★","LTE GPS追跡・活動量モニタリング・バッテリー3ヶ月持続の次世代スマート犬用首輪。脱走・迷子対策として日本のペット飼育者向け。"],
  [2168,"ペット用品","Rollwi Smart Pet Water Fountain","Rollwi","Amazon US","https://rollwi.com/","https://rollwi.com/","support@rollwi.com","★★★★☆","静音ポンプ・フィルター・水量センサー搭載のスマートペット給水器。猫の飲水不足による腎臓病予防に効果的な製品。"],
  [2169,"ペット用品","Petcube Play 2 Wi-Fi Pet Camera","Petcube","Petcube公式","https://petcube.com/play2","https://petcube.com/","hello@petcube.com","★★★★★","レーザーポインター内蔵・1080p・双方向音声のスマートペットカメラ。猫と遠隔でレーザー遊びが可能。留守中のペットとのつながりを実現。"],
  [2170,"ペット用品","PawSafe Dog Car Seat Harness","PawSafe","Amazon US","https://www.pawsafepets.com/","https://www.pawsafepets.com/","support@pawsafepets.com","★★★★☆","衝突安全テスト済みの犬用カーシートハーネス。ペット同乗時の安全基準が強化される中、安全意識の高い日本の飼育者向け製品。"],
  [2171,"ペット用品","Whistle Health Smart Collar Insight","Whistle","Whistle公式","https://www.whistle.com/health","https://www.whistle.com/","support@whistle.com","★★★★★","24時間健康モニタリング・AI異常行動検知・GPSを一体化したオールインワンスマートペット首輪。ペットの健康を総合管理できる製品。"],
  [2172,"ペット用品","Petlibro Capsule Pet Feeder Connect","Petlibro","Amazon US","https://petlibro.com/","https://petlibro.com/","support@petlibro.com","★★★★★","360度カメラ・1080p・ビデオ通話しながら給餌できるスマート自動給餌器。ペットとのつながりを保ちながら食事管理が可能。"],
  [2173,"ペット用品","Anker Eufy Pet Camera E220","Eufy","Amazon US","https://www.eufylife.com/","https://www.eufylife.com/","support@eufy.com","★★★★★","2Kカメラ・AI吠え検知・犬猫顔認識・双方向音声のペットカメラ。月額不要でローカル保存対応。コスパ最高のペットカメラ。"],
  [2174,"ペット用品","DOGTV Streaming Service","DOGTV","DOGTV公式","https://www.dogtv.com/","https://www.dogtv.com/","info@dogtv.com","★★★★☆","犬の分離不安・ストレス軽減向けに制作された犬専用映像ストリーミングサービス。留守番中の犬のメンタルヘルスケア向けサービス。"],
  [2175,"ペット用品","Wickedbone Smart Interactive Dog Toy","Wickedball","Amazon US","https://cheerble.com/wickedbone","https://cheerble.com/","support@cheerble.com","★★★★☆","アプリ操作・自動モードで動き回る犬用スマートインタラクティブトイ。一人遊び・分離不安対策として留守番犬向けに最適。"],
  [2176,"ペット用品","Puppr Dog Training App with Sara Carson","Puppr","Puppr公式","https://www.pupprapp.com/","https://www.pupprapp.com/","support@pupprapp.com","★★★★☆","プロドッグトレーナー監修の段階的犬のしつけトレーニングアプリ。初心者でも家庭でプロ品質のトレーニングが可能。"],
  [2177,"ペット用品","Smart Pet Scale Body Condition System","Petivity","Petivity公式","https://petivity.com/","https://petivity.com/","info@petivity.com","★★★★★","Purina製のスマートペット体重計・活動量モニタリングシステム。体重変化・行動パターン変化から健康問題を早期発見。"],
  [2178,"ペット用品","CleanPaws Pet Paw Washer Portable","CleanPaws","Amazon US","https://cleanpaws.us/","https://cleanpaws.us/","support@cleanpaws.us","★★★★☆","電動ブラシ内蔵の携帯型ペット足洗い器。散歩後の足洗いを30秒で完了。清潔好きな日本の犬の飼育者向けに需要が高い製品。"],
  [2179,"ペット用品","Outward Hound Nina Ottosson Dog Puzzle","Outward Hound","Amazon US","https://www.outwardhound.com/","https://www.outwardhound.com/","support@outwardhound.com","★★★★☆","犬の知育・脳トレを促すパズルトイ。留守番中の退屈解消・認知症予防として老犬から子犬まで対応。ペット知育玩具の定番ブランド。"],
  [2180,"ペット用品","K&H Pet Products Self-Warming Lounge Sleeper","K&H Pet Products","Amazon US","https://www.khmfg.com/","https://www.khmfg.com/","support@khmfg.com","★★★★☆","電源不要・体温で発熱する自己発熱型ペットベッド。冬の寒さ対策・老犬の関節ケアに最適。省エネ設計で日本の省エネ志向に合致。"],
  [2181,"ペット用品","PetFusion 3-Sided Cat Scratcher Lounge","PetFusion","Amazon US","https://petfusion.com/","https://petfusion.com/","support@petfusion.com","★★★★★","3面使用可能・リバーシブルのプレミアム猫用爪とぎ兼ベッド。Amazonでカテゴリ1位の定番猫用品。爪とぎ対策として日本向け。"],
  [2182,"ペット用品","Comsmart Pet Nail Grinder Trimmer","Comsmart","Amazon US","https://comsmart.us/","https://comsmart.us/","support@comsmart.us","★★★☆☆","USB充電・超静音・LED搭載のペット用爪グラインダー。切る恐怖なく安全に爪ケアが可能。ペットグルーミングを自宅で行いたい飼育者向け。"],
  [2183,"ペット用品","SureFeed Microchip Pet Feeder Connect","SureFlap","SureFlap公式","https://www.surepetcare.com/sureFeed","https://www.surepetcare.com/","support@surepetcare.com","★★★★★","マイクロチップ認識で特定のペットのみ給餌・Wi-Fi接続・食事量記録のスマート自動給餌器。多頭飼い・療法食管理に最適。"],

  // ===== テクノロジー・ガジェット (2184-2203) =====
  [2184,"テクノロジー・ガジェット","Humane AI Pin Wearable AI Device","Humane","Humane公式","https://hu.ma.ne/","https://hu.ma.ne/","support@hu.ma.ne","★★★★★","スクリーンレス・レーザー投影・AI音声操作のウェアラブルAIデバイス。スマートフォンの次世代として注目。革新的なスクリーンフリー体験。"],
  [2185,"テクノロジー・ガジェット","Rabbit R1 AI Pocket Device","Rabbit Inc.","Rabbit公式","https://www.rabbit.tech/","https://www.rabbit.tech/","support@rabbit.tech","★★★★☆","Large Action Model搭載でアプリを代わりに操作するポケット型AI端末。CES 2024でバイラルになった次世代AIデバイス。"],
  [2186,"テクノロジー・ガジェット","Frame AI Smart Glasses","Brilliant Labs","Brilliant Labs公式","https://brilliant.xyz/frame","https://brilliant.xyz/","hello@brilliant.xyz","★★★★★","ChatGPT・Perplexity連動・オープンソースのAI機能搭載スマートグラス。軽量でファッショナブル。AI活用の日常化に向けた製品。"],
  [2187,"テクノロジー・ガジェット","Lenovo ThinkPad X1 Fold 16 Foldable PC","Lenovo","Lenovo公式","https://www.lenovo.com/thinkpad-x1-fold","https://www.lenovo.com/","https://www.lenovo.com/us/en/contact/","★★★★★","16.3インチ折り畳みOLED・タブレット・ラップトップ両用のフォルダブルPC。次世代PCフォームファクターとして世界中で注目される革新製品。"],
  [2188,"テクノロジー・ガジェット","Asus ROG Ally X Gaming Handheld","Asus","Asus公式","https://rog.asus.com/gaming-handhelds/rog-ally/","https://www.asus.com/","support@asus.com","★★★★★","AMD Ryzen Z1 Extreme・24GB RAM・1TBのプレミアムWindowsゲームハンドヘルド。Steam Deck対抗の最強ゲームデバイス。"],
  [2189,"テクノロジー・ガジェット","Polaroid Now+ Gen 2 Instant Camera","Polaroid","Polaroid公式","https://www.polaroid.com/now-plus","https://www.polaroid.com/","support@polaroid.com","★★★★★","Bluetoothアプリ連動・多重露光・Lパック対応のスマートインスタントカメラ。フィルムカメラ・レトロ写真ブームで需要急上昇中。"],
  [2190,"テクノロジー・ガジェット","Fujifilm Instax Mini 12 Instant Camera","Fujifilm","Fujifilm公式","https://instax.com/mini12/","https://www.fujifilm.com/","support@fujifilm.com","★★★★★","コンパクト・パステルカラー展開・自撮り対応の大人気チェキシリーズ新モデル。世界的に売れているチェキブランドを海外に紹介する逆輸入事例。"],
  [2191,"テクノロジー・ガジェット","Moment iPhone Anamorphic Lens","Moment","Moment公式","https://www.shopmoment.com/","https://www.shopmoment.com/","support@shopmoment.com","★★★★★","スマートフォンをシネマカメラ化するアナモルフィックレンズ。映画的な横長フォーマットと光のフレアを実現。動画クリエイター向け。"],
  [2192,"テクノロジー・ガジェット","Hollyland Lark M2 Wireless Microphone","Hollyland","Hollyland公式","https://www.hollyland.com/","https://www.hollyland.com/","support@hollyland.com","★★★★★","クリップ型・2人同時収録・ノイズキャンセリング搭載の超コンパクト無線マイク。YouTube・TikTok・取材向けマイクとして業界標準化しつつある製品。"],
  [2193,"テクノロジー・ガジェット","Bambu Lab P1P Multi-Material 3D Printer","Bambu Lab","Bambu Lab公式","https://bambulab.com/p1p","https://bambulab.com/","support@bambulab.com","★★★★★","4色同時印刷・AMS自動材料チェンジャー・高速マルチカラー対応3Dプリンター。世界最速クラスのFDM 3Dプリンターとして話題。"],
  [2194,"テクノロジー・ガジェット","Elmo Doc Camera TT-12 Document Camera","Elmo","Elmo公式","https://www.elmousa.com/","https://www.elmousa.com/","info@elmousa.com","★★★★★","日本発の老舗書画カメラブランドElmoの最新モデル。教育・ビジネスプレゼンにおける書類・現物の共有に最適。海外市場での日本ブランド紹介事例。"],
  [2195,"テクノロジー・ガジェット","Teenage Engineering EP-133 K.O. II Sampler","Teenage Engineering","Teenage Engineering公式","https://teenage.engineering/products/ep-133","https://teenage.engineering/","support@teenage.engineering","★★★★★","スウェーデン発の革新的サンプラー＆ビートマシン。直感的な操作性とプロ品質サウンドで音楽制作の民主化を実現。音楽制作好きな日本向け。"],
  [2196,"テクノロジー・ガジェット","Anker Nano Power Bank Built-in Cable","Anker","Amazon US","https://www.anker.com/products/a1259","https://www.anker.com/","support@anker.com","★★★★★","Lightning/USB-Cケーブル内蔵・22.5W・名刺サイズの超小型モバイルバッテリー。持ち運びが楽な軽量設計で日本の通勤者向けに最適。"],
  [2197,"テクノロジー・ガジェット","Elgato Wave XLR Studio Podcasting Kit","Elgato","Amazon US","https://www.elgato.com/wave-xlr","https://www.elgato.com/","support@elgato.com","★★★★★","XLRマイク接続・リアルタイムミキシング・クラウドリフトプリアンプ搭載のポッドキャスト向けオーディオインターフェース。配信者向け。"],
  [2198,"テクノロジー・ガジェット","Insta360 Link 2C 4K Webcam","Insta360","Insta360公式","https://www.insta360.com/product/insta360-link2c","https://www.insta360.com/","support@insta360.com","★★★★★","AI自動追跡・ジンバル搭載・4K解像度の次世代Webカメラ。在宅ワーク・ライブ配信でプロ品質の映像を自動で維持する製品。"],
  [2199,"テクノロジー・ガジェット","Logitech Sight Meeting Camera","Logitech","Logitech公式","https://www.logitech.com/sight","https://www.logitech.com/","support@logitech.com","★★★★★","テーブル中心部に置く会議室向け360度AI会議カメラ。発言者を自動追尾・全員の顔を同時フレームに収める次世代会議ソリューション。"],
  [2200,"テクノロジー・ガジェット","Creality K1 Max 3D Printer","Creality","Creality公式","https://www.creality.com/products/creality-k1-max","https://www.creality.com/","support@creality.com","★★★★★","300×300mmビルドプレート・600mm/s高速印刷・AIカメラ監視の大型高速3Dプリンター。コスパ最高クラスの製品として世界中でヒット。"],
  [2201,"テクノロジー・ガジェット","DJI Mic 2 Wireless Microphone System","DJI","DJI公式","https://www.dji.com/mic-2","https://www.dji.com/","support@dji.com","★★★★★","250mワイヤレス・ノイズキャンセリング・32bit Float録音のコンテンツクリエイター向けワイヤレスマイクシステム。DJIブランドで信頼性高い。"],
  [2202,"テクノロジー・ガジェット","Plaud Note AI Voice Recorder Card","Plaud","Plaud公式","https://www.plaud.ai/note","https://www.plaud.ai/","hello@plaud.ai","★★★★★","iPhone・Androidと磁気吸着でくっつくカード型AI議事録レコーダー。会議・インタビューを録音・文字起こし・要約を自動化。"],
  [2203,"テクノロジー・ガジェット","Jackery Solar Generator 1000 Plus","Jackery","Jackery公式","https://www.jackery.com/products/solar-generator-1000-plus","https://www.jackery.com/","support@jackery.com","★★★★★","2000Wh拡張可能・双方向充電・ソーラー対応のポータブル電源。キャンプ・防災向けに世界最大シェアのJackeryブランド最新中型機。"],

  // ===== 美容・スキンケア (2204-2223) =====
  [2204,"美容・スキンケア","Solawave 4-in-1 Skincare Wand","Solawave","Solawave公式","https://solawave.co/","https://solawave.co/","support@solawave.co","★★★★★","赤色LED・微弱電流・温熱・振動の4機能を搭載したアンチエイジングスキンケアワンド。2分でプロフェッショナルケアを自宅で実現。"],
  [2205,"美容・スキンケア","Therabody Theraface PRO Facial Massage","Therabody","Therabody公式","https://www.therabody.com/theraface-pro","https://www.therabody.com/","support@therabody.com","★★★★★","パーカッション振動・LED光・マイクロカレントのフェイシャルマッサージデバイス。Theragunのフェイス版として筋肉ケアブランドが美容に進出。"],
  [2206,"美容・スキンケア","Hydrafacial Keravive Scalp Treatment","HydraFacial","HydraFacial公式","https://www.hydrafacial.com/","https://www.hydrafacial.com/","info@hydrafacial.com","★★★★★","頭皮の詰まりを取り除き・保湿・成長因子を補給する頭皮専門美容機器。薄毛・頭皮ケアに積極的な日本の高齢男性・女性向け。"],
  [2207,"美容・スキンケア","Glo2Facial Portable Oxygeneo Device","Geneo","Geneo公式","https://geneo.com/","https://geneo.com/","info@geneo.com","★★★★★","毛穴洗浄・酸素補給・栄養成分導入の3ステップを一台で行うオキシジェネオ美容機器。ポータブルサイズで自宅エステを実現。"],
  [2208,"美容・スキンケア","iS Clinical Active Serum","iS Clinical","iS Clinical公式","https://isclinical.com/","https://isclinical.com/","support@isclinical.com","★★★★★","グリコール酸・白金・乳酸配合の最高評価スキンケアセラム。皮膚科医・エステティシャンから高く評価されるクリニカルグレード製品。"],
  [2209,"美容・スキンケア","Fillmed Art Filler Pure Volume Skin Booster","Fillmed","Fillmed公式","https://www.fillmed.com/","https://www.fillmed.com/","info@fillmed.com","★★★★★","ヒアルロン酸・NCTF成長因子配合のプロフェッショナルスキンブースター。クリニック・エステ向けの高品質フィラー製品。アンチエイジング需要向け。"],
  [2210,"美容・スキンケア","Augustinus Bader The Cream","Augustinus Bader","Augustinus Bader公式","https://www.augustinusbader.com/","https://www.augustinusbader.com/","support@augustinusbader.com","★★★★★","幹細胞研究から生まれたTFC8テクノロジー配合の最高峰フェイスクリーム。世界のセレブ愛用高級スキンケアとして日本の富裕層向け。"],
  [2211,"美容・スキンケア","Elemis Pro-Collagen Marine Cream","Elemis","Elemis公式","https://www.elemis.com/","https://www.elemis.com/","support@elemis.com","★★★★★","英国No.1スキンケアブランドのパドルウィード・グリロニア配合アンチエイジングクリーム。2分で肌が引き締まるとレビューで高評価。"],
  [2212,"美容・スキンケア","COSRX Advanced Snail 96 Mucin Essence","COSRX","Amazon US","https://www.cosrx.com/","https://www.cosrx.com/","support@cosrx.com","★★★★★","カタツムリ粘液96%配合の韓国発人気美容液。Amazon韓国コスメカテゴリで圧倒的人気。Kビューティとして日本でも需要急拡大中。"],
  [2213,"美容・スキンケア","Sunday Riley Good Genes Lactic Acid Treatment","Sunday Riley","Sunday Riley公式","https://www.sundayriley.com/","https://www.sundayriley.com/","support@sundayriley.com","★★★★★","乳酸酸・ライコリシス配合でくすみ・ニキビ跡・シワを改善するAHAトリートメント。セフォラで最も人気の高い美容液の一つ。"],
  [2214,"美容・スキンケア","Skin Better Science AlphaRet Overnight Cream","SkinBetter Science","SkinBetter公式","https://skinbetter.com/","https://skinbetter.com/","support@skinbetter.com","★★★★★","レチノイドとAHAを組み合わせた刺激の少ない高効果アンチエイジングナイトクリーム。皮膚科医推薦の最高評価製品。"],
  [2215,"美容・スキンケア","Drunk Elephant Protini Polypeptide Cream","Drunk Elephant","Drunk Elephant公式","https://www.drunkelephant.com/","https://www.drunkelephant.com/","support@drunkelephant.com","★★★★★","9種のシグナルペプチド・成長因子配合の究極の保湿クリーム。世界的人気クリーンビューティーブランドとして日本でも需要大。"],
  [2216,"美容・スキンケア","Paula's Choice BHA Skin Perfecting 2% Liquid","Paula's Choice","Paula's Choice公式","https://www.paulaschoice.com/","https://www.paulaschoice.com/","support@paulaschoice.com","★★★★★","毛穴詰まり・ニキビ・角質改善の定番BHAエクスフォリアント。スキンケア愛好家の間で世界的に最も評価が高い製品の一つ。"],
  [2217,"美容・スキンケア","Tatcha The Water Cream","Tatcha","Tatcha公式","https://www.tatcha.com/","https://www.tatcha.com/","support@tatcha.com","★★★★★","日本の和の美意識から生まれた米国発高級スキンケアブランド。日本文化を逆輸入した形で海外での日本美容成分の人気を紹介。"],
  [2218,"美容・スキンケア","FOREO LUNA 4 Face Massager","Foreo","Foreo公式","https://www.foreo.com/luna4","https://www.foreo.com/","support@foreo.com","★★★★★","T-Sonicパルス振動・シリコン毛・16段階強度調整の最新スマートクレンジングデバイス。アプリ連動で個別スキンケアルーティンを提案。"],
  [2219,"美容・スキンケア","111Skin Y Theorem Repair Serum NASA","111Skin","111Skin公式","https://www.111skin.com/","https://www.111skin.com/","support@111skin.com","★★★★★","NASA宇宙飛行士向け傷治癒研究から生まれたNAC Y²配合の高機能修復美容液。宇宙科学技術×美容という独自ストーリーが日本市場に響く。"],
  [2220,"美容・スキンケア","Dermalogica Daily Microfoliant Powder","Dermalogica","Dermalogica公式","https://www.dermalogica.com/","https://www.dermalogica.com/","support@dermalogica.com","★★★★★","ライスブラン・パパイン酵素配合の毎日使える酵素パウダー洗顔料。世界のプロエステティシャンに最も支持されるブランドの定番製品。"],
  [2221,"美容・スキンケア","Murad Retinol Youth Renewal Serum","Murad","Murad公式","https://www.murad.com/","https://www.murad.com/","support@murad.com","★★★★★","3種のレチノールフォーム配合でシワ・たるみ・ハリ不足に対応するアンチエイジングセラム。皮膚科医設立ブランドの高信頼性製品。"],
  [2222,"美容・スキンケア","Shiseido Bio-Performance Skin Filler Serum","Shiseido","Shiseido公式","https://www.shiseido.com/","https://www.shiseido.com/","https://www.shiseido.com/us/en/customer-service.html","★★★★★","資生堂の最先端バイオ研究による肌充填美容液。海外市場での資生堂ブランドの高評価を逆輸入する形でリストに掲載。"],
  [2223,"美容・スキンケア","Summer Fridays Jet Lag Mask","Summer Fridays","Summer Fridays公式","https://summerfridays.com/","https://summerfridays.com/","support@summerfridays.com","★★★★★","旅行・疲れた肌向けの濃密保湿マスク。インフルエンサー発のクリーンビューティーブランドとして世界的に人気急上昇中。"],

  // ===== 子供・教育 (2224-2243) =====
  [2224,"子供・教育","Kano PC Tablet Build-it-Yourself Kit","Kano Computing","Kano公式","https://kano.me/","https://kano.me/","support@kano.me","★★★★★","子供自身で組み立てるタッチスクリーンPC・プログラミング学習セット。コンピューターの仕組みを理解しながら組み立てる体験型STEM製品。"],
  [2225,"子供・教育","LittleBits Rule Your Room Kit","Sphero (LittleBits)","Amazon US","https://littlebits.com/","https://littlebits.com/","support@littlebits.com","★★★★★","電子回路モジュールを組み合わせてスマートルームデバイスを作るエレクトロニクス学習キット。工作・電子工学入門として日本の親子向け。"],
  [2226,"子供・教育","Anki Cozmo Robot for Education","Digital Dream Labs","Amazon US","https://www.digitaldreamlabs.com/cozmo","https://www.digitaldreamlabs.com/","support@digitaldreamlabs.com","★★★★★","感情表現・ゲーム・Scratch/Pythonプログラミングに対応するAI教育ロボット。米国の小学校・塾でSTEM教育向けに広く採用。"],
  [2227,"子供・教育","Snap Circuits Elite Electronics Kit","Elenco","Amazon US","https://www.elenco.com/snap-circuits","https://www.elenco.com/","support@elenco.com","★★★★★","750種以上のプロジェクトに対応する電子回路学習キット。ハンダ不要でスナップ式に回路を組み立て電気・電子の基礎を学ぶ。"],
  [2228,"子供・教育","Cubetto Wooden Coding Robot","Primo Toys","Cubetto公式","https://www.primotoys.com/","https://www.primotoys.com/","hello@primotoys.com","★★★★★","スクリーンを一切使わない木製プログラミングロボット。3歳から始められるモンテッソーリ式コーディング玩具として世界で人気。"],
  [2229,"子供・教育","Dash Wonder Workshop Coding Toy","Wonder Workshop","Amazon US","https://www.makewonder.com/dash","https://www.makewonder.com/","support@makewonder.com","★★★★★","音・光・障害物回避機能を持つプログラミング学習ロボット。Scratch Jrからブロックリー・JavaScriptと段階的にスキルアップ可能。"],
  [2230,"子供・教育","Qubit Education Quantum Computing Kit","Qubit by Qubit","Qubit公式","https://www.qubitbyqubit.org/","https://www.qubitbyqubit.org/","info@qubitbyqubit.org","★★★★★","中高生向け量子コンピューティング教育プログラム。IBMと提携の次世代STEM教育として米国の有名大学進学校で採用。将来の量子人材育成向け。"],
  [2231,"子供・教育","Prodigy Math Educational Game","Prodigy Education","Prodigy公式","https://www.prodigygame.com/","https://www.prodigygame.com/","support@prodigygame.com","★★★★★","RPGゲームで算数を楽しく学ぶオンライン数学教育プラットフォーム。全米の学校で採用されており、日本の学習塾でも応用可能。"],
  [2232,"子供・教育","LeapFrog LeapStart Learning System","LeapFrog","Amazon US","https://www.leapfrog.com/leapstart","https://www.leapfrog.com/","support@leapfrog.com","★★★★★","専用ペンで本を読むと音声・動画で学べるインタラクティブ幼児学習システム。幼児教育の定番米国ブランドで日本でも認知度向上中。"],
  [2233,"子供・教育","Orbo Play AI Tutor Companion","Orbo","Kickstarter","https://orboedu.com/","https://orboedu.com/","hello@orboedu.com","★★★★★","子供の学習スタイル・弱点をAIが分析してパーソナライズ学習計画を作成するAI学習コンパニオン。EdTech最新世代の製品。"],
  [2234,"子供・教育","Play Osmo Words & Numbers Bundle","Osmo","Osmo公式","https://www.playosmo.com/","https://www.playosmo.com/","support@playosmo.com","★★★★★","物理カードとiPadアプリが連動する幼児向けインタラクティブ学習ゲーム。文字・数字・算数を遊びながら覚えられる体験型教育製品。"],
  [2235,"子供・教育","Bitsbox Coding Subscription for Kids","Bitsbox","Bitsbox公式","https://bitsbox.com/","https://bitsbox.com/","hello@bitsbox.com","★★★★☆","毎月コーディングプロジェクトが届くサブスクリプション。Scratch・HTML・JavaScriptを段階的に学べる子供向けプログラミング教材。"],
  [2236,"子供・教育","Science Museum Activity Kits","Science Museum Group","Amazon UK","https://www.sciencemuseumshop.co.uk/","https://www.sciencemuseumshop.co.uk/","info@sciencemuseumshop.co.uk","★★★★☆","英国科学博物館公認の科学実験キット。電気・化学・物理の実験を自宅で体験。理科教育向けに日本の親子向けに最適なSTEMキット。"],
  [2237,"子供・教育","Roblox Education Coding Platform","Roblox Corporation","Roblox公式","https://education.roblox.com/","https://education.roblox.com/","education@roblox.com","★★★★★","Robloxゲームエンジンを使ってLuaプログラミングを学ぶオンライン教育プラットフォーム。世界中の子供が遊ぶRobloxで楽しくコーディング習得。"],
  [2238,"子供・教育","Tinkercad 3D Modeling for Kids","Autodesk","Tinkercad公式","https://www.tinkercad.com/","https://www.tinkercad.com/","https://www.autodesk.com/company/contact-us","★★★★★","ブラウザ上で無料で使える子供向け3Dモデリングツール。3Dプリント用データ作成・回路設計・コードブロックも学べる統合STEM環境。"],
  [2239,"子供・教育","Learning Resources Botley 2.0 Coding Robot","Learning Resources","Amazon US","https://www.learningresources.com/","https://www.learningresources.com/","info@learningresources.com","★★★★★","スクリーンフリーでコーディング概念を学ぶ幼児向けロボット。シーケンス・ループ・関数などプログラミングの基礎概念を遊びながら習得。"],
  [2240,"子供・教育","Lingokids App Bilingual Learning for Kids","Lingokids","Lingokids公式","https://lingokids.com/","https://lingokids.com/","support@lingokids.com","★★★★★","AIと英語ネイティブ講師によるゲーム型子供英語学習アプリ。2〜8歳向け。英語教育熱が高い日本の親にとって魅力的なサービス。"],
  [2241,"子供・教育","BrainPOP Jr. Educational Videos","BrainPOP","BrainPOP公式","https://jr.brainpop.com/","https://jr.brainpop.com/","support@brainpop.com","★★★★☆","アニメキャラクターが学校の授業を楽しく解説するK-3向け動画教育プラットフォーム。米国学校で広く採用されている教育コンテンツサービス。"],
  [2242,"子供・教育","DoodleMaths Personalised Maths App","DoodleLearning","DoodleMaths公式","https://www.doodlelearning.com/","https://www.doodlelearning.com/","support@doodlelearning.com","★★★★★","AIが苦手分野を特定して毎日10分のパーソナライズ算数学習を提供するアプリ。英国の小学校で採用。塾競争の激しい日本の親向け。"],
  [2243,"子供・教育","Wonder Workshop Dash & Dot Wonder Pack","Wonder Workshop","Amazon US","https://www.makewonder.com/wonder-pack","https://www.makewonder.com/","support@makewonder.com","★★★★★","DashとDotの2台セット・50以上のアクセサリー・200時間以上のカリキュラム付きの総合コーディング学習セット。学校向け・家庭向け両対応。"],

  // ===== ファッション・アクセサリー (2244-2263) =====
  [2244,"ファッション・アクセサリー","Caraa Sport Convertible Gym Bag","Caraa","Caraa公式","https://www.caraa.com/","https://www.caraa.com/","hello@caraa.com","★★★★★","ジム→オフィス→街中と場面に合わせて形状が変わるコンバーチブルバッグ。スポーツ×ファッション両立の現代的ワーキングウーマン向け。"],
  [2245,"ファッション・アクセサリー","Unbound Merino Wool Travel Shirt","Unbound Merino","Unbound公式","https://unboundmerino.com/","https://unboundmerino.com/","support@unboundmerino.com","★★★★★","防臭・温度調節・旅行着不要のメリノウール製旅行用シャツ。1枚で1週間着られるトラベルウェア。旅行・長期滞在向けに世界的人気。"],
  [2246,"ファッション・アクセサリー","Wooly Pocket Sustainable Wall Planter","Wooly Pocket","Amazon US","https://woolypocket.com/","https://woolypocket.com/","info@woolypocket.com","★★★★☆","リサイクルプラスチックボトル素材の壁掛け型プランターバッグ。インドア植物ブームの日本向けインテリアグリーンアイテム。"],
  [2247,"ファッション・アクセサリー","Raden A22 Smart Carry-On Suitcase","Raden","Raden公式","https://www.raden.com/","https://www.raden.com/","support@raden.com","★★★★★","USB充電・重量センサー・位置情報共有機能付きのスマートスーツケース。テック系ビジネス旅行者向けのハイテクキャリーオン。"],
  [2248,"ファッション・アクセサリー","TUMI Alpha 3 Ballistic Nylon Backpack","TUMI","TUMI公式","https://www.tumi.com/","https://www.tumi.com/","support@tumi.com","★★★★★","バリスティックナイロン・ステルスプロテクション仕様のプレミアムビジネスバックパック。日本でも高い人気を誇るラグジュアリービジネスバッグ。"],
  [2249,"ファッション・アクセサリー","Lululemon Everywhere Belt Bag 1L","Lululemon","Lululemon公式","https://shop.lululemon.com/","https://shop.lululemon.com/","support@lululemon.com","★★★★★","世界的にバイラルになったルルレモンのベルトバッグ。ヨガ・スポーツ・カジュアルシーンでも使いやすく日本の若い女性に人気拡大中。"],
  [2250,"ファッション・アクセサリー","Matein Travel Laptop Backpack","Matein","Amazon US","https://mateingear.com/","https://mateingear.com/","support@mateingear.com","★★★★☆","TSAロック・USB充電ポート・防水素材の多機能旅行用バックパック。コスパ最強クラスで旅行・通学・ビジネス向けに日本での需要大。"],
  [2251,"ファッション・アクセサリー","Allbirds Tree Dasher 2 Running Shoe","Allbirds","Allbirds公式","https://www.allbirds.com/","https://www.allbirds.com/","support@allbirds.com","★★★★★","ユーカリ繊維・メリノウール使用のカーボンフットプリント表示付きサスティナブルランニングシューズ。環境配慮ブランドとして世界的人気。"],
  [2252,"ファッション・アクセサリー","Baggallini Quick Pocket Crossbody Bag","Baggallini","Amazon US","https://www.baggallini.com/","https://www.baggallini.com/","support@baggallini.com","★★★★☆","軽量・防水・スリ防止ジッパー・多機能収納のトラベルクロスボディバッグ。女性旅行者に大人気の機能的バッグとして日本向け。"],
  [2253,"ファッション・アクセサリー","Corkcicle Commuter Cup Insulated","Corkcicle","Amazon US","https://corkcicle.com/","https://corkcicle.com/","support@corkcicle.com","★★★★★","3層真空断熱・片手操作・傾き防止底面のプレミアムコミューターカップ。車通勤・自転車通勤向けに最適な高機能タンブラー。"],
  [2254,"ファッション・アクセサリー","Patagonia Black Hole 25L Pack","Patagonia","Patagonia公式","https://www.patagonia.com/","https://www.patagonia.com/","support@patagonia.com","★★★★★","100%リサイクル素材・防水・耐久性の高いパタゴニアの定番デイパック。環境活動でも有名なブランドで環境意識の高い日本の若者向け。"],
  [2255,"ファッション・アクセサリー","Aime Leon Dore Wool-Mohair Camp Hat","Aime Leon Dore","ALD公式","https://www.aimeleondore.com/","https://www.aimeleondore.com/","support@aimeleondore.com","★★★★☆","ニューヨーク発のラグジュアリーストリートブランド。プレミアム素材×クラシックアメリカンスタイルで世界的人気を誇るブランド。"],
  [2256,"ファッション・アクセサリー","Ridge Wallet Carbon Fiber RFID Blocker","Ridge","Ridge公式","https://www.ridgewallet.com/","https://www.ridgewallet.com/","support@ridgewallet.com","★★★★★","カーボンファイバー・40枚まで収納・RFID遮断のスリム財布。キャッシュレス化が進む日本の若い世代向けのミニマリスト財布。"],
  [2257,"ファッション・アクセサリー","Olive + Oak Genuine Leather Belt","Olive + Oak","Amazon US","https://oliveandoak.com/","https://oliveandoak.com/","support@oliveandoak.com","★★★★☆","本革・手縫い・ステッチ仕上げのプレミアムレザーベルト。品質重視の日本人向けにビジネス・カジュアル両用の定番アクセサリー。"],
  [2258,"ファッション・アクセサリー","Fossil Gen 6 Hybrid Smartwatch","Fossil","Amazon US","https://www.fossil.com/gen-6-hybrid","https://www.fossil.com/","support@fossil.com","★★★★☆","アナログ時計×スマート機能のハイブリッドスマートウォッチ。健康・通知機能を持ちながら従来の時計デザインを維持。スーツに合うスマートウォッチ。"],
  [2259,"ファッション・アクセサリー","Zipp Featherweight Water Bottle","Zipp","Amazon US","https://www.zipp.com/","https://www.zipp.com/","support@zipp.com","★★★★☆","Zipp製カーボンファイバー採用の世界最軽量クラス自転車用ウォーターボトル。プロサイクリスト御用達の高性能アクセサリー。"],
  [2260,"ファッション・アクセサリー","Dagne Dover Scout Backpack","Dagne Dover","Dagne Dover公式","https://www.dagnedover.com/","https://www.dagnedover.com/","support@dagnedover.com","★★★★★","ネオプレン素材・ノートPC収納・ウォーターボトルポケット搭載の都市型デイパック。機能美を追求した女性向けプレミアムバックパック。"],
  [2261,"ファッション・アクセサリー","Tangle Teezer Compact Styler Detangling Brush","Tangle Teezer","Amazon US","https://www.tangleteezer.com/","https://www.tangleteezer.com/","support@tangleteezer.com","★★★★☆","独自2段式ブリスル設計のデタングリングブラシ。絡まった髪を痛みなくほどく英国発の人気ヘアケアアイテム。日本でも認知度上昇中。"],
  [2262,"ファッション・アクセサリー","Brilliant Earth Sustainable Lab Diamond Ring","Brilliant Earth","Brilliant Earth公式","https://www.brilliantearth.com/","https://www.brilliantearth.com/","support@brilliantearth.com","★★★★★","ラボグロウンダイヤモンド・リサイクルゴールド使用のエシカルジュエリー。環境・倫理に配慮した婚約指輪として世界的に需要が急拡大。"],
  [2263,"ファッション・アクセサリー","Johnny Bigg Plus Size Men's Fashion","Johnny Bigg","Johnny Bigg公式","https://us.johnnybigg.com/","https://us.johnnybigg.com/","support@johnnybigg.com","★★★★☆","2XL〜9XLまで対応するプラスサイズ男性向けファッションブランド。体型多様性への意識が高まる日本でも注目されるインクルーシブファッション。"],

  // ===== クリーニング・収納・整理 (2264-2283) =====
  [2264,"クリーニング・収納・整理","iRobot Braava Jet m6 Robot Mop","iRobot","iRobot公式","https://www.irobot.com/braava-jet-m6","https://www.irobot.com/","support@irobot.com","★★★★★","AI間取りマッピング・精密噴射・Roombaと連携する最上位ロボットモップ。Roomba掃除後にBraavaが水拭きする完全自動フロアケア。"],
  [2265,"クリーニング・収納・整理","Tineco FLOOR ONE S7 Steam Wet Dry Vacuum","Tineco","Tineco公式","https://www.tineco.com/","https://www.tineco.com/","support@tineco.com","★★★★★","吸引・スチーム・水拭きを一台で実現するスマートフロアクリーナー。iLoop AIセンサーが汚れを自動検知して洗浄力を調整する次世代製品。"],
  [2266,"クリーニング・収納・整理","Bissell CrossWave Pet Pro Wet Dry Vacuum","Bissell","Amazon US","https://www.bissell.com/crosswave-pet-pro","https://www.bissell.com/","support@bissell.com","★★★★★","吸引+水洗いを同時に行うコードレスフロアクリーナー。ペットの毛・汚れを一度に除去。ペット飼育者の多い日本向けに特に有用。"],
  [2267,"クリーニング・収納・整理","Shark Stratos Cordless Vacuum","Shark","Shark公式","https://www.sharkclean.com/","https://www.sharkclean.com/","support@shark.com","★★★★★","AI量の検出・CleanSense IQ技術・FlexEdge設計のコードレス掃除機。Dysonと並ぶ米国シェアを誇るブランドの最上位コードレス機。"],
  [2268,"クリーニング・収納・整理","Nanit Pro Smart Baby Monitor","Nanit","Nanit公式","https://www.nanit.com/","https://www.nanit.com/","support@nanit.com","★★★★★","AIで赤ちゃんの睡眠パターン・呼吸モニタリング・授乳記録を行うスマートベビーモニター。少子化対策に関心の高い日本の子育て世代向け。"],
  [2269,"クリーニング・収納・整理","Whirlpool WFW9620HW Front Load Washer","Whirlpool","Whirlpool公式","https://www.whirlpool.com/","https://www.whirlpool.com/","https://www.whirlpool.com/contact-us.html","★★★★★","スマホ操作・AI汚れ検知・省エネのドラム式スマート洗濯機。米国最大シェアの洗濯機ブランドが日本未進出で高い参入余地。"],
  [2270,"クリーニング・収納・整理","Electrolux PerfectCare 800 Steam Iron","Electrolux","Electrolux公式","https://www.electrolux.com/","https://www.electrolux.com/","support@electrolux.com","★★★★☆","スウェーデン電機ブランドの高性能スチームアイロン。最高温度・連続スチーム・自己洗浄機能搭載。衣類のシワ取り需要が高い日本向け。"],
  [2271,"クリーニング・収納・整理","iRobot Roomba j9+ Self-Emptying Robot Vacuum","iRobot","iRobot公式","https://www.irobot.com/roomba-j9-plus","https://www.irobot.com/","support@irobot.com","★★★★★","Dirt Detective AI・Smart Scrub・自動ゴミ収集ステーション搭載のRoomba最上位機種。AIが汚れの場所を学習して徹底的に清掃。"],
  [2272,"クリーニング・収納・整理","OXO SteeL Soap Dispensing Palm Brush","OXO","Amazon US","https://www.oxo.com/","https://www.oxo.com/","support@oxo.com","★★★★☆","洗剤を内蔵したハンドル部分から適量を供給しながら洗えるソープディスペンシングパームブラシ。食器洗いを効率化する実用的キッチンアイテム。"],
  [2273,"クリーニング・収納・整理","Anker Eufy Clean X9 Pro Robot Vacuum","Eufy","Eufy公式","https://www.eufylife.com/","https://www.eufylife.com/","support@eufy.com","★★★★★","双回転モップ・自動洗浄ステーション・AI障害物認識搭載のロボット掃除機。Eufyブランドの旗艦モデルで高いコスパとプレミアム機能を両立。"],
  [2274,"クリーニング・収納・整理","Steamfast SF-370WH Steam Cleaner","Steamfast","Amazon US","https://www.steamfast.com/","https://www.steamfast.com/","support@steamfast.com","★★★★☆","99.9%の殺菌効果・多用途アタッチメント付きの家庭用スチームクリーナー。化学洗剤不要で床・タイル・キッチンを清潔に維持。"],
  [2275,"クリーニング・収納・整理","Brabantia Sort & Go Waste Bin System","Brabantia","Amazon US","https://www.brabantia.com/","https://www.brabantia.com/","support@brabantia.com","★★★★★","分別収集・スタッキング可能なモジュール式ゴミ箱システム。ゴミの分別が厳しい日本の分別ルールに完全対応できる機能的ゴミ箱。"],
  [2276,"クリーニング・収納・整理","Elago Floating Sink Soap Dispenser","Elago","Amazon US","https://www.elago.com/","https://www.elago.com/","support@elago.com","★★★★☆","水面に浮かんだように見えるユニークデザインのシンク置きソープディスペンサー。ミニマルなデザインが日本のバスルーム・キッチン向け。"],
  [2277,"クリーニング・収納・整理","Clorox ToiletWand Disposable Toilet Cleaning System","Clorox","Amazon US","https://www.clorox.com/","https://www.clorox.com/","https://www.clorox.com/contact/","★★★★☆","使い捨てパッド・ブラシ不要のトイレ掃除システム。衛生意識が高く手軽なトイレ掃除を求める日本の単身・共働き世帯向け。"],
  [2278,"クリーニング・収納・整理","LIXIL Smart Toilet Shower SATIS S","LIXIL","LIXIL公式","https://www.lixil.com/","https://www.lixil.com/","https://www.lixil.com/en/contact/","★★★★★","LIXILの最上位スマートトイレ。海外市場で日本のウォシュレット文化を輸出する逆輸入製品として価値大。日本ブランドの世界展開事例。"],
  [2279,"クリーニング・収納・整理","Method Antibac Kitchen Cleaner","Method","Amazon US","https://methodproducts.com/","https://methodproducts.com/","hello@methodproducts.com","★★★★☆","植物由来の抗菌成分・生分解性・スタイリッシュボトルのエコ対応キッチン洗浄剤。デザイン性の高い洗剤として日本でも需要が広がりつつある。"],
  [2280,"クリーニング・収納・整理","Sterilite WeatherTight Storage Box Set","Sterilite","Amazon US","https://www.sterilite.com/","https://www.sterilite.com/","support@sterilite.com","★★★★☆","パッキンシール・ラッチ付きの防湿防塵収納ボックスセット。押入れ・ガレージ収納に最適。断捨離・収納整理が好きな日本市場向け定番。"],
  [2281,"クリーニング・収納・整理","Simplehuman Slim Rectangular Step Trash Can","Simplehuman","Amazon US","https://www.simplehuman.com/","https://www.simplehuman.com/","support@simplehuman.com","★★★★★","ステンレス・ソフトクローズ蓋・スリムデザインの高級ゴミ箱。狭い日本の住宅にも収まるスリム設計で機能美を兼ね備えたプレミアム製品。"],
  [2282,"クリーニング・収納・整理","Hoover ONEPWR Evolve Pet Cordless Vacuum","Hoover","Amazon US","https://www.hoover.com/","https://www.hoover.com/","support@hoover.com","★★★★☆","交換式バッテリー・ペット毛専用タービンツール搭載のコードレス掃除機。ペット飼育者向けに特化した機能で日本のペットオーナー向け。"],
  [2283,"クリーニング・収納・整理","Pura Smart Home Fragrance Diffuser","Pura","Pura公式","https://pura.com/","https://pura.com/","support@pura.com","★★★★★","アプリでスケジュール管理・香り強度調整・交互フレグランス切替のスマートアロマディフューザー。香りのサブスクサービス連携で日本向け芳香市場に訴求。"],

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
