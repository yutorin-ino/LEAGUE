/**
 * add_grommet_new_100_401_500.js
 * The Grommet専用ダッシュボード用・追加100件 (No.401〜500)
 * grommet_400.js (No.1-400) と合算してNo.1〜500に拡張する
 * 各カテゴリ10件ずつ均等配分（既存400件は各40件→拡張後は各50件）
 */

// [番号, カテゴリ, 製品名, メーカー, ECサイト, 製品URL, メーカーHP, メール, ★, コメント]
const NEW_PRODUCTS_100 = [
  // ── キッチン・調理器具 (10件・401-410) ──
  [401,'キッチン・調理器具','Zulay Original Milk Frother Deluxe','Zulay Kitchen','The Grommet','https://thegrommet.com/product/kitchen/zulay-original-milk-frother-deluxe','https://zulaykitchen.com','hello@zulaykitchen.com','★★★★','コードレス・高速攪拌でカプチーノ泡立てを実現するミルクフォーマー'],
  [402,'キッチン・調理器具','Talisman Designs Adjustable Rolling Pin','Talisman Designs','The Grommet','https://thegrommet.com/product/kitchen/talisman-designs-adjustable-rolling-pin','https://talismandesigns.com','support@talismandesigns.com','★★★','リング交換式・生地の厚みを均一調整できる麺棒'],
  [403,'キッチン・調理器具','Yedi Instant Yogurt Maker','Yedi Houseware','The Grommet','https://thegrommet.com/product/kitchen/yedi-instant-yogurt-maker','https://yedihouseware.com','support@yedihouseware.com','★★★★','8つの発酵モード搭載・自宅で本格ヨーグルトを作れる調理器'],
  [404,'キッチン・調理器具','Bar Keepers Friend Cookware Cleaner','Bar Keepers Friend','The Grommet','https://thegrommet.com/product/kitchen/bar-keepers-friend-cookware-cleaner','https://barkeepersfriend.com','info@barkeepersfriend.com','★★★★','シュウ酸配合・頑固な焦げ付き/水垢を落とす調理器具クリーナー'],
  [405,'キッチン・調理器具','Salt & Pour Adjustable Salt Cellar','Salt & Pour','The Grommet','https://thegrommet.com/product/kitchen/salt-and-pour-adjustable-cellar','https://saltandpour.com','hello@saltandpour.com','★★★','スライド式蓋・湿気を防ぐ調整可能ソルトセラー'],
  [406,'キッチン・調理器具','Cuisinart Griddler 5-in-1','Cuisinart','The Grommet','https://thegrommet.com/product/kitchen/cuisinart-griddler-5-in-1','https://cuisinart.com','support@cuisinart.com','★★★★','グリル/パニーニ/グリドル切替可能な5-in-1調理家電'],
  [407,'キッチン・調理器具','Souper Cubes Silicone Freezer Tray','Souper Cubes','The Grommet','https://thegrommet.com/product/kitchen/souper-cubes-silicone-freezer-tray','https://soupercubes.com','hello@soupercubes.com','★★★★★','1カップ分ずつ冷凍保存できるスープ/だし用シリコントレー'],
  [408,'キッチン・調理器具','Levo Oil Infusion Machine','Levo','The Grommet','https://thegrommet.com/product/kitchen/levo-oil-infusion-machine','https://levooil.com','support@levooil.com','★★★★','ボタン一つでハーブ/香辛料オイルを抽出する自動インフューザー'],
  [409,'キッチン・調理器具','Nostalgia Retro Mini Donut Maker','Nostalgia','The Grommet','https://thegrommet.com/product/kitchen/nostalgia-retro-mini-donut-maker','https://nostalgiaproducts.com','support@nostalgiaproducts.com','★★★','レトロデザイン・4分で焼き上がるミニドーナツメーカー'],
  [410,'キッチン・調理器具','Fox Run Adjustable Rolling Pin Rings','Fox Run','The Grommet','https://thegrommet.com/product/kitchen/fox-run-adjustable-rolling-pin-rings','https://foxrunbrands.com','info@foxrunbrands.com','★★★','4段階厚み調整リング付き伸縮ローリングピン'],

  // ── スマートホーム・インテリア・照明 (10件・411-420) ──
  [411,'スマートホーム・インテリア・照明','Mr. Beams Wireless Motion Sensor Light','Mr. Beams','The Grommet','https://thegrommet.com/product/home/mr-beams-wireless-motion-sensor-light','https://mrbeams.com','support@mrbeams.com','★★★★','電池式・工事不要でどこでも設置できるモーションセンサーライト'],
  [412,'スマートホーム・インテリア・照明','Chilipad Sleep System Cooling Mattress Pad','ChiliSleep','The Grommet','https://thegrommet.com/product/home/chilipad-sleep-system','https://chilisleep.com','support@chilisleep.com','★★★★★','水循環式・体感温度を個別制御できる冷却マットレスパッド'],
  [413,'スマートホーム・インテリア・照明','Umbra Anigram Wall Decor Hooks','Umbra','The Grommet','https://thegrommet.com/product/home/umbra-anigram-wall-decor-hooks','https://umbra.com','hello@umbra.com','★★★','動物シルエット型・壁を飾りながら使えるウォールフック'],
  [414,'スマートホーム・インテリア・照明','Bissell Little Green Portable Carpet Cleaner','Bissell','The Grommet','https://thegrommet.com/product/home/bissell-little-green-carpet-cleaner','https://bissell.com','support@bissell.com','★★★★★','コンパクト設計・ソファ/階段/車内対応ポータブル洗浄機'],
  [415,'スマートホーム・インテリア・照明','Cricut Joy Compact Cutting Machine','Cricut','The Grommet','https://thegrommet.com/product/home/cricut-joy-compact-cutting-machine','https://cricut.com','support@cricut.com','★★★★','スマホ連携・ラベルやカードを手軽に作れる小型カッティングマシン'],
  [416,'スマートホーム・インテリア・照明','Levoit Classic 300S Smart Humidifier','Levoit','The Grommet','https://thegrommet.com/product/home/levoit-classic-300s-humidifier','https://levoit.com','support@levoit.com','★★★★','アプリ/音声操作対応・湿度自動調整スマート加湿器'],
  [417,'スマートホーム・インテリア・照明','Command Strips Picture Hanging Kit','Command','The Grommet','https://thegrommet.com/product/home/command-strips-picture-hanging-kit','https://command.com','support@command.com','★★★','壁を傷つけず貼って剥がせる額縁用取り付けストリップ'],
  [418,'スマートホーム・インテリア・照明','Twinkly Smart App-Controlled LED String Lights','Twinkly','The Grommet','https://thegrommet.com/product/home/twinkly-smart-led-string-lights','https://twinkly.com','support@twinkly.com','★★★★','個別LED制御・アプリでパターン作成できるスマートイルミネーション'],
  [419,'スマートホーム・インテリア・照明','Simplehuman Voice-Activated Soap Pump','simplehuman','The Grommet','https://thegrommet.com/product/home/simplehuman-voice-activated-soap-pump','https://simplehuman.com','hello@simplehuman.com','★★★★','音声認識・タッチフリーで泡ソープを出すディスペンサー'],
  [420,'スマートホーム・インテリア・照明','Homedics UV-Clean Portable Sanitizer Bag','HoMedics','The Grommet','https://thegrommet.com/product/home/homedics-uv-clean-sanitizer-bag','https://homedics.com','support@homedics.com','★★★','UV-C照射・スマホ/鍵などの小物を除菌するポータブルバッグ'],

  // ── ウェアラブル・ヘルス・フィットネス (10件・421-430) ──
  [421,'ウェアラブル・ヘルス・フィットネス','Theraface Connect Percussive Facial Device','Therabody','The Grommet','https://thegrommet.com/product/health-wellness/theraface-connect','https://therabody.com','support@therabody.com','★★★★','振動ヘッド交換式・顔の緊張をほぐすフェイシャルデバイス'],
  [422,'ウェアラブル・ヘルス・フィットネス','Vive Cold Therapy Wrap','Vive Health','The Grommet','https://thegrommet.com/product/health-wellness/vive-cold-therapy-wrap','https://vivehealth.com','support@vivehealth.com','★★★','ジェルパック内蔵・関節/筋肉痛の冷却療法用ラップ'],
  [423,'ウェアラブル・ヘルス・フィットネス','Fisher Wallace Stimulator Anxiety Relief Device','Fisher Wallace','The Grommet','https://thegrommet.com/product/health-wellness/fisher-wallace-stimulator','https://fisherwallace.com','support@fisherwallace.com','★★★★','FDA登録・微弱電流で不安/不眠を緩和する頭部装着デバイス'],
  [424,'ウェアラブル・ヘルス・フィットネス','Bala Power Ring Weighted Fitness Tool','Bala','The Grommet','https://thegrommet.com/product/health-wellness/bala-power-ring','https://yourbala.com','hello@yourbala.com','★★★','重量入りリング・ヨガ/ピラティスの負荷を高めるフィットネスツール'],
  [425,'ウェアラブル・ヘルス・フィットネス','RENPHO Eye Massager with Heat','RENPHO','The Grommet','https://thegrommet.com/product/health-wellness/renpho-eye-massager-heat','https://renpho.com','support@renpho.com','★★★★','空気圧+振動+温熱・目の疲れをほぐすアイマッサージャー'],
  [426,'ウェアラブル・ヘルス・フィットネス','Theragun Mini Percussive Massage Device','Therabody','The Grommet','https://thegrommet.com/product/health-wellness/theragun-mini','https://therabody.com','support@therabody.com','★★★★★','携帯サイズ・3段階速度筋膜リリースマッサージデバイス'],
  [427,'ウェアラブル・ヘルス・フィットネス','Vitruvian Trainer+ Digital Resistance Machine','Vitruvian','The Grommet','https://thegrommet.com/product/health-wellness/vitruvian-trainer-plus','https://vitruvian.com','support@vitruvian.com','★★★★','デジタル負荷調整・コンパクト収納型トレーニングマシン'],
  [428,'ウェアラブル・ヘルス・フィットネス','Nurokor MiBody Pain Relief Patch','Nurokor','The Grommet','https://thegrommet.com/product/health-wellness/nurokor-mibody-patch','https://nurokor.com','support@nurokor.com','★★★','EMS+TENS複合波形・貼るだけ疼痛緩和パッチ'],
  [429,'ウェアラブル・ヘルス・フィットネス','Wahl Clean & Cordless Groin Trimmer','Wahl','The Grommet','https://thegrommet.com/product/health-wellness/wahl-clean-cordless-trimmer','https://wahlusa.com','support@wahlusa.com','★★★','防水コードレス・デリケートゾーン用低刺激トリマー'],
  [430,'ウェアラブル・ヘルス・フィットネス','Aura Strap Smart Recovery Compression Boots','Aura Strap','The Grommet','https://thegrommet.com/product/health-wellness/aura-strap-recovery-boots','https://theaurastrap.com','support@theaurastrap.com','★★★★','段階的空気圧縮・アスリート向け脚部リカバリーブーツ'],

  // ── アウトドア・スポーツ・旅行 (10件・431-440) ──
  [431,'アウトドア・スポーツ・旅行','Biolite Firepit+ Smokeless Fire Pit','BioLite','The Grommet','https://thegrommet.com/product/outdoors-garden/biolite-firepit-plus','https://bioliteenergy.com','hello@bioliteenergy.com','★★★★★','内蔵ファンで2次燃焼・煙を大幅削減するポータブル焚き火台'],
  [432,'アウトドア・スポーツ・旅行','Nemo Switchback Ultralight Sleeping Pad','NEMO Equipment','The Grommet','https://thegrommet.com/product/outdoors-garden/nemo-switchback-sleeping-pad','https://nemoequipment.com','support@nemoequipment.com','★★★★','折りたたみ式クローズドセル・穴の心配がない軽量マット'],
  [433,'アウトドア・スポーツ・旅行','Gerber Bear Grylls Compact Multi-Tool','Gerber Gear','The Grommet','https://thegrommet.com/product/outdoors-garden/gerber-bear-grylls-multi-tool','https://gerbergear.com','support@gerbergear.com','★★★','コンパクト設計・12機能内蔵サバイバルマルチツール'],
  [434,'アウトドア・スポーツ・旅行','Aer Travel Pack 3 Modular Backpack','Aer','The Grommet','https://thegrommet.com/product/travel/aer-travel-pack-3','https://aer.com','support@aer.com','★★★★','機内持ち込み対応・PC/衣類収納最適化トラベルバックパック'],
  [435,'アウトドア・スポーツ・旅行','Coleman OneSource Rechargeable Lantern','Coleman','The Grommet','https://thegrommet.com/product/outdoors-garden/coleman-onesource-lantern','https://coleman.com','support@coleman.com','★★★★','モバイルバッテリー機能付き・充電式LEDキャンプランタン'],
  [436,'アウトドア・スポーツ・旅行','Sea to Summit Ultra-Sil Dry Sack','Sea to Summit','The Grommet','https://thegrommet.com/product/travel/sea-to-summit-ultra-sil-dry-sack','https://seatosummit.com','info@seatosummit.com','★★★','超軽量30デニール・防水コンプレッション収納袋'],
  [437,'アウトドア・スポーツ・旅行','Rumpl NanoLoft Travel Blanket','Rumpl','The Grommet','https://thegrommet.com/product/travel/rumpl-nanoloft-travel-blanket','https://rumpl.com','hello@rumpl.com','★★★★','撥水加工・機内でも使える軽量パッカブルブランケット'],
  [438,'アウトドア・スポーツ・旅行','Yeti Rambler Bottle with Chug Cap','YETI','The Grommet','https://thegrommet.com/product/outdoors-garden/yeti-rambler-chug-cap','https://yeti.com','support@yeti.com','★★★★★','二重壁ステンレス構造・保冷保温力抜群のボトル'],
  [439,'アウトドア・スポーツ・旅行','Kelty Discovery Trail Tent','Kelty','The Grommet','https://thegrommet.com/product/outdoors-garden/kelty-discovery-trail-tent','https://kelty.com','support@kelty.com','★★★','初心者向け・設営簡単な4人用ファミリーキャンプテント'],
  [440,'アウトドア・スポーツ・旅行','Matador FlatPak Soap Bar Case','Matador','The Grommet','https://thegrommet.com/product/travel/matador-flatpak-soap-bar-case','https://matadoru.com','support@matadoru.com','★★★','平らに収納できる漏れ防止石鹸ケース'],

  // ── ペット用品 (10件・441-450) ──
  [441,'ペット用品','West Paw Toppl Treat Dispensing Toy','West Paw','The Grommet','https://thegrommet.com/product/pets/west-paw-toppl-treat-dispensing-toy','https://westpaw.com','info@westpaw.com','★★★★','コーン型・おやつを詰めて長時間遊べる犬用知育トイ'],
  [442,'ペット用品','Yeti Boomer 8 Dog Bowl','YETI','The Grommet','https://thegrommet.com/product/pets/yeti-boomer-8-dog-bowl','https://yeti.com','support@yeti.com','★★★★','錆びにくいステンレス製・タフなペット用食器'],
  [443,'ペット用品','PetSafe Easy Walk Front-Clip Harness','PetSafe','The Grommet','https://thegrommet.com/product/pets/petsafe-easy-walk-harness','https://petsafe.net','support@petsafe.net','★★★★','胸元クリップ式・引っ張り軽減ドッグハーネス'],
  [444,'ペット用品','Meowingtons Cat Wall Shelves Set','Meowingtons','The Grommet','https://thegrommet.com/product/pets/meowingtons-cat-wall-shelves','https://meowingtons.com','hello@meowingtons.com','★★★','組み立て自由・壁面を活用する猫用ウォークウェイ'],
  [445,'ペット用品','Kong Wobbler Treat Dispensing Toy','KONG','The Grommet','https://thegrommet.com/product/pets/kong-wobbler-treat-dispensing-toy','https://kongcompany.com','support@kongcompany.com','★★★★','揺れる度におやつが出る早食い防止知育トイ'],
  [446,'ペット用品','Necoichi Portable Stainless Steel Cat Bowl','Necoichi','The Grommet','https://thegrommet.com/product/pets/necoichi-portable-cat-bowl','https://necoichi.jp','support@necoichi.jp','★★★','日本製・折りたたみ式携帯用ステンレス猫ボウル'],
  [447,'ペット用品','Bark Bright Enzymatic Dog Toothpaste','Bark Bright','The Grommet','https://thegrommet.com/product/pets/bark-bright-enzymatic-toothpaste','https://barkbright.com','hello@barkbright.com','★★★','酵素配合・歯磨き嫌いの犬でも使いやすい歯みがき粉'],
  [448,'ペット用品','Frisco Elevated Dog Feeder','Frisco','The Grommet','https://thegrommet.com/product/pets/frisco-elevated-dog-feeder','https://frisco.com','support@frisco.com','★★★','高さ調整可能・大型犬の消化をサポートする台座付き給餌台'],
  [449,'ペット用品','Petkit Fresh Element Solo Pet Feeder','PETKIT','The Grommet','https://thegrommet.com/product/pets/petkit-fresh-element-solo-feeder','https://petkit.com','support@petkit.com','★★★★','密閉保存・タイマー給餌対応スマートフードコンテナ'],
  [450,'ペット用品','Zippypaws Skinny Peltz No Stuffing Dog Toy','ZippyPaws','The Grommet','https://thegrommet.com/product/pets/zippypaws-skinny-peltz-toy','https://zippypaws.com','support@zippypaws.com','★★★','詰め物なし・キューキュー音が鳴る耐久犬用トイ3体セット'],

  // ── テクノロジー・ガジェット (10件・451-460) ──
  [451,'テクノロジー・ガジェット','Anker PowerConf S3 Speakerphone','Anker','The Grommet','https://thegrommet.com/product/tech/anker-powerconf-s3-speakerphone','https://anker.com','support@anker.com','★★★★','AIノイズ除去・在宅会議用コンパクトスピーカーフォン'],
  [452,'テクノロジー・ガジェット','Twelve South BookArc Laptop Stand','Twelve South','The Grommet','https://thegrommet.com/product/tech/twelve-south-bookarc-laptop-stand','https://twelvesouth.com','support@twelvesouth.com','★★★','縦置き収納・デスクスペースを節約するノートPCスタンド'],
  [453,'テクノロジー・ガジェット','Logitech MX Anywhere 3S Compact Mouse','Logitech','The Grommet','https://thegrommet.com/product/tech/logitech-mx-anywhere-3s','https://logitech.com','support@logitech.com','★★★★','どんな素材の上でも使える高精度コンパクトマウス'],
  [454,'テクノロジー・ガジェット','Elgato Key Light Air Streaming Light','Elgato','The Grommet','https://thegrommet.com/product/tech/elgato-key-light-air','https://elgato.com','support@elgato.com','★★★★','アプリ調光対応・配信/ビデオ通話用LEDライト'],
  [455,'テクノロジー・ガジェット','Anker 737 Charger GaNPrime 120W','Anker','The Grommet','https://thegrommet.com/product/tech/anker-737-charger-ganprime-120w','https://anker.com','support@anker.com','★★★★★','GaN素材・3ポート同時急速充電対応コンパクト充電器'],
  [456,'テクノロジー・ガジェット','Belkin SoundForm Rise Wireless Earbuds','Belkin','The Grommet','https://thegrommet.com/product/tech/belkin-soundform-rise-earbuds','https://belkin.com','support@belkin.com','★★★','IPX5防水・ケース付きワイヤレスイヤホン'],
  [457,'テクノロジー・ガジェット','Twelve South Curve Riser Laptop Stand','Twelve South','The Grommet','https://thegrommet.com/product/tech/twelve-south-curve-riser','https://twelvesouth.com','support@twelvesouth.com','★★★','アルミ削り出し・視線を上げる姿勢改善ラップトップスタンド'],
  [458,'テクノロジー・ガジェット','Peak Design Everyday Sling Camera Bag','Peak Design','The Grommet','https://thegrommet.com/product/tech/peak-design-everyday-sling','https://peakdesign.com','support@peakdesign.com','★★★★★','カメラ+ガジェット収納・都市型スリングバッグ'],
  [459,'テクノロジー・ガジェット','Anker Soundcore Motion Boom Plus Speaker','Anker Soundcore','The Grommet','https://thegrommet.com/product/tech/soundcore-motion-boom-plus','https://soundcore.com','support@soundcore.com','★★★★','IP67防水・大音量アウトドア用Bluetoothスピーカー'],
  [460,'テクノロジー・ガジェット','Logitech Casa Pop-Up Desk Hub','Logitech','The Grommet','https://thegrommet.com/product/tech/logitech-casa-pop-up-desk-hub','https://logitech.com','support@logitech.com','★★★','USB-C一本で複数機器接続できるポップアップデスクハブ'],

  // ── 美容・スキンケア (10件・461-470) ──
  [461,'美容・スキンケア','Vanicream Gentle Facial Cleanser','Vanicream','The Grommet','https://thegrommet.com/product/beauty/vanicream-gentle-facial-cleanser','https://vanicream.com','support@vanicream.com','★★★★','無香料・敏感肌向け低刺激洗顔料'],
  [462,'美容・スキンケア','Conair InfinitiPro Rotating Curling Iron','Conair','The Grommet','https://thegrommet.com/product/beauty/conair-infinitipro-curling-iron','https://conair.com','support@conair.com','★★★','自動回転バレル・巻き髪を簡単に作れるカーリングアイロン'],
  [463,'美容・スキンケア','Tula 24-7 Moisturizer','Tula','The Grommet','https://thegrommet.com/product/beauty/tula-24-7-moisturizer','https://tula.com','hello@tula.com','★★★★','プロバイオティクス配合・オールインワン保湿クリーム'],
  [464,'美容・スキンケア','Revlon Hair Dryer Brush Volumizer','Revlon','The Grommet','https://thegrommet.com/product/beauty/revlon-hair-dryer-brush-volumizer','https://revlon.com','support@revlon.com','★★★★','ワンステップでブロー+ボリュームを叶えるヘアドライヤーブラシ'],
  [465,'美容・スキンケア','Beautybio GloPRO Microneedling Regeneration Tool','BeautyBio','The Grommet','https://thegrommet.com/product/beauty/beautybio-glopro-microneedling-tool','https://beautybio.com','support@beautybio.com','★★★★★','家庭用マイクロニードリング・美容液浸透をサポートするツール'],
  [466,'美容・スキンケア','Naturium Niacinamide Serum','Naturium','The Grommet','https://thegrommet.com/product/beauty/naturium-niacinamide-serum','https://naturium.com','support@naturium.com','★★★★','12%ナイアシンアミド配合・毛穴/皮脂ケアセラム'],
  [467,'美容・スキンケア','Wet Brush Original Detangler Hairbrush','Wet Brush','The Grommet','https://thegrommet.com/product/beauty/wet-brush-original-detangler','https://wetbrush.com','support@wetbrush.com','★★★','IntelliFlex毛先・濡れた髪も痛くなくとかせるブラシ'],
  [468,'美容・スキンケア','Mele & Co. Lighted Makeup Vanity Mirror','Mele & Co.','The Grommet','https://thegrommet.com/product/beauty/mele-lighted-vanity-mirror','https://meleandco.com','support@meleandco.com','★★★','LEDリング付き・拡大鏡搭載卓上メイクミラー'],
  [469,'美容・スキンケア','Isle of Paradise Self-Tanning Drops','Isle of Paradise','The Grommet','https://thegrommet.com/product/beauty/isle-of-paradise-self-tanning-drops','https://isleofparadise.com','hello@isleofparadise.com','★★★★','好きな保湿剤に混ぜて使う色調整可能セルフタンニングドロップ'],
  [470,'美容・スキンケア','Conair Double Ionic Hair Straightener','Conair','The Grommet','https://thegrommet.com/product/beauty/conair-double-ionic-straightener','https://conair.com','support@conair.com','★★★','イオン技術・ツヤ髪ストレートを叶えるヘアアイロン'],

  // ── 子供・教育 (10件・471-480) ──
  [471,'子供・教育','Green Toys Recycled Plastic Dump Truck','Green Toys','The Grommet','https://thegrommet.com/product/family-kids/green-toys-dump-truck','https://greentoys.com','support@greentoys.com','★★★★','米国製リサイクルプラスチック・BPAフリー知育ダンプトラック'],
  [472,'子供・教育','Hape Pound & Tap Bench Toy','Hape','The Grommet','https://thegrommet.com/product/family-kids/hape-pound-and-tap-bench','https://hape.com','support@hape.com','★★★','木製・叩いて音を鳴らす幼児向け木琴ベンチトイ'],
  [473,'子供・教育','Learning Resources Coding Critters Ranger & Zip','Learning Resources','The Grommet','https://thegrommet.com/product/family-kids/coding-critters-ranger-and-zip','https://learningresources.com','support@learningresources.com','★★★★','ペット育成+基礎プログラミングを学べるロボットトイ'],
  [474,'子供・教育','Melissa & Doug Wooden Shape Sorting Cube','Melissa & Doug','The Grommet','https://thegrommet.com/product/family-kids/melissa-doug-shape-sorting-cube','https://melissaanddoug.com','info@melissaanddoug.com','★★★★','木製・形合わせで手指の発達を促す知育キューブ'],
  [475,'子供・教育','Educational Insights GeoSafari Jr. Talking Microscope','Educational Insights','The Grommet','https://thegrommet.com/product/family-kids/geosafari-jr-talking-microscope','https://educationalinsights.com','support@educationalinsights.com','★★★★','音声ガイド付き・子供向け顕微鏡セット'],
  [476,'子供・教育','Fat Brain Toys Dimpl Sensory Toy','Fat Brain Toys','The Grommet','https://thegrommet.com/product/family-kids/fat-brain-toys-dimpl','https://fatbraintoys.com','hello@fatbraintoys.com','★★★','ぷにぷに押して遊ぶ赤ちゃん向け感覚刺激トイ'],
  [477,'子供・教育','Melissa & Doug Magnetic Wooden Puzzle Board','Melissa & Doug','The Grommet','https://thegrommet.com/product/family-kids/melissa-doug-magnetic-puzzle-board','https://melissaanddoug.com','info@melissaanddoug.com','★★★★','磁石ピース・持ち運びできる旅行用木製パズル'],
  [478,'子供・教育','Learning Resources Pretend & Play Cash Register','Learning Resources','The Grommet','https://thegrommet.com/product/family-kids/pretend-and-play-cash-register','https://learningresources.com','support@learningresources.com','★★★','おままごとで計算/お金の概念を学ぶレジスタートイ'],
  [479,'子供・教育','Plan Toys Sustainable Wooden Balancing Toy','Plan Toys','The Grommet','https://thegrommet.com/product/family-kids/plantoys-balancing-toy','https://plantoys.com','support@plantoys.com','★★★★','ゴムの木素材・環境配慮型バランス知育玩具'],
  [480,'子供・教育','Hand2mind Base Ten Blocks Math Kit','hand2mind','The Grommet','https://thegrommet.com/product/family-kids/hand2mind-base-ten-blocks','https://hand2mind.com','support@hand2mind.com','★★★','具体物操作で位取りを学ぶ算数教材キット'],

  // ── ファッション・アクセサリー (10件・481-490) ──
  [481,'ファッション・アクセサリー','Baggallini Everywhere Crossbody Bag','Baggallini','The Grommet','https://thegrommet.com/product/fashion/baggallini-everywhere-crossbody-bag','https://baggallini.com','support@baggallini.com','★★★★','軽量ナイロン・多ポケット設計トラベルクロスボディ'],
  [482,'ファッション・アクセサリー','Sanuk Yoga Sling Sandals','Sanuk','The Grommet','https://thegrommet.com/product/fashion/sanuk-yoga-sling-sandals','https://sanuk.com','support@sanuk.com','★★★','ヨガマット素材インソール・履き心地重視サンダル'],
  [483,'ファッション・アクセサリー','Chaco Z/Cloud Sport Sandals','Chaco','The Grommet','https://thegrommet.com/product/fashion/chaco-z-cloud-sport-sandals','https://chacos.com','support@chacos.com','★★★★','調整可能ストラップ・アーチサポート付きスポーツサンダル'],
  [484,'ファッション・アクセサリー','Vera Bradley Cotton Lighten Up Travel Bag','Vera Bradley','The Grommet','https://thegrommet.com/product/fashion/vera-bradley-lighten-up-travel-bag','https://verabradley.com','support@verabradley.com','★★★','軽量コットン・鮮やかなパターンの旅行バッグ'],
  [485,'ファッション・アクセサリー','Buff Multifunctional Headwear','Buff','The Grommet','https://thegrommet.com/product/fashion/buff-multifunctional-headwear','https://buff.com','support@buff.com','★★★★','12通り以上の使い方ができるチューブ型ヘッドウェア'],
  [486,'ファッション・アクセサリー','Danner Mountain 600 Hiking Boots','Danner','The Grommet','https://thegrommet.com/product/fashion/danner-mountain-600-hiking-boots','https://danner.com','support@danner.com','★★★★★','米国製・防水設計耐久性ハイキングブーツ'],
  [487,'ファッション・アクセサリー','Toms Alpargata Classic Slip-On Shoes','TOMS','The Grommet','https://thegrommet.com/product/fashion/toms-alpargata-classic-slip-on','https://toms.com','support@toms.com','★★★','購入ごとに寄付・キャンバス地スリッポンシューズ'],
  [488,'ファッション・アクセサリー','Filson Original Briefcase Leather','Filson','The Grommet','https://thegrommet.com/product/fashion/filson-original-briefcase','https://filson.com','support@filson.com','★★★★★','米国製ブライドルレザー・生涯使えるビジネスブリーフケース'],
  [489,'ファッション・アクセサリー','Vuori Ponto Performance Jogger','Vuori','The Grommet','https://thegrommet.com/product/fashion/vuori-ponto-performance-jogger','https://vuoriclothing.com','support@vuoriclothing.com','★★★★','動きやすさ重視・タウンユースにも対応するジョガーパンツ'],
  [490,'ファッション・アクセサリー','Injinji Toe Socks Performance','Injinji','The Grommet','https://thegrommet.com/product/fashion/injinji-toe-socks-performance','https://injinji.com','support@injinji.com','★★★','5本指設計・マメ防止ランニング用トウソックス'],

  // ── クリーニング・収納・整理 (10件・491-500) ──
  [491,'クリーニング・収納・整理','OXO Good Grips Salad Spinner','OXO','The Grommet','https://thegrommet.com/product/home/oxo-good-grips-salad-spinner','https://oxo.com','info@oxo.com','★★★★','ワンプッシュブレーキ・水切れの良いサラダスピナー'],
  [492,'クリーニング・収納・整理','Simplehuman Dish Rack Steel Frame','simplehuman','The Grommet','https://thegrommet.com/product/home/simplehuman-dish-rack-steel-frame','https://simplehuman.com','hello@simplehuman.com','★★★★','錆びないスチールフレーム・シンク横に収まる水切りラック'],
  [493,'クリーニング・収納・整理','Cleancult Dish Soap Refill Pods','Cleancult','The Grommet','https://thegrommet.com/product/home/cleancult-dish-soap-refill-pods','https://cleancult.com','support@cleancult.com','★★★','紙製リフィル・プラスチック削減食器用洗剤詰め替え'],
  [494,'クリーニング・収納・整理','Household Essentials Collapsible Laundry Hamper','Household Essentials','The Grommet','https://thegrommet.com/product/home/household-essentials-collapsible-hamper','https://householdessentials.com','support@householdessentials.com','★★★','折りたたみ式・省スペース収納ランドリーハンパー'],
  [495,'クリーニング・収納・整理','Bissell CrossWave Wet Dry Vacuum','Bissell','The Grommet','https://thegrommet.com/product/home/bissell-crosswave-wet-dry-vacuum','https://bissell.com','support@bissell.com','★★★★★','掃除機+モップ一体型・硬い床とラグ両対応クリーナー'],
  [496,'クリーニング・収納・整理','Grove Collaborative Walnut Scrubber Sponge','Grove Collaborative','The Grommet','https://thegrommet.com/product/home/grove-walnut-scrubber-sponge','https://grove.co','support@grove.co','★★★','クルミ殻使用・傷つけずに汚れを落とすスポンジ'],
  [497,'クリーニング・収納・整理','Full Circle Tersus Dish Brush','Full Circle Home','The Grommet','https://thegrommet.com/product/home/full-circle-tersus-dish-brush','https://fullcirclehome.com','support@fullcirclehome.com','★★★','交換可能ブラシヘッド・立てて乾かせる食器洗いブラシ'],
  [498,'クリーニング・収納・整理','mDesign Bathroom Vanity Organizer Tray','mDesign','The Grommet','https://thegrommet.com/product/home/mdesign-bathroom-vanity-organizer-tray','https://mdesign.com','support@mdesign.com','★★★','仕切り可動式・洗面台の小物整理トレー'],
  [499,'クリーニング・収納・整理','iRobot Roomba Combo Essential Robot','iRobot','The Grommet','https://thegrommet.com/product/home/irobot-roomba-combo-essential','https://irobot.com','support@irobot.com','★★★★★','吸引+水拭き一体型・コンパクト設計ロボット掃除機'],
  [500,'クリーニング・収納・整理','Rubbermaid FreshWorks Produce Saver Set','Rubbermaid','The Grommet','https://thegrommet.com/product/home/rubbermaid-freshworks-produce-saver','https://rubbermaid.com','support@rubbermaid.com','★★★★','通気孔+フィルター設計・野菜/果物を長持ちさせる保存容器'],
];

module.exports = NEW_PRODUCTS_100;
