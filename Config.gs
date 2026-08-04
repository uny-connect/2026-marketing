/**
 * [공통 설정] 시트명, 캘린더명, 업체별 시트 ID 모음
 */
const CONFIG = {
  // 시스템에서 사용하는 시트 이름 정리
  SHEETS: {
    MASTER_SCHEDULE: "2026_Schedule",
    JAPAN_SCHEDULE: "2026スケジュール",
    CALENDAR_BOT: "캘린더봇",
    REPORT: "2026_보고서",
    MASTER_DATA: "Master_Data"
  },
  
  // 구글 캘린더 이름 정리
  CALENDARS: {
    JAPAN_TRIP: "韓国ブロガー取材",
    COVERAGE: "[알람]취재 스케줄"
  },

  // 업체별 결과 보고서 스프레드시트 고유 ID Map
  CLIENT_MAP: {
    "HILLTOP RESORT FUKUOKA様": "129NkbaVEqUxbEKnw1SxdFctVa_OYscSXCXnVn-f4bTc",

    "おんくり唐津様": "129NkbaVEqUxbEKnw1SxdFctVa_OYscSXCXnVn-f4bTc",

    "クロスライフ博多天神様": "1FBzQicNUC2AaMm0alHLGkw0g28gkuZXG-qk9pkpLi6s",
    "クロスライフ博多柳橋様": "1FBzQicNUC2AaMm0alHLGkw0g28gkuZXG-qk9pkpLi6s",

    "ホテルフォルツァ博多駅筑紫口I様": "18nlMjLbjxAIbobBquPRqeJKMCRGCG0g7ogfW2WaPjfE",
    "ホテルフォルツァ博多駅筑紫口Ⅱ様": "18nlMjLbjxAIbobBquPRqeJKMCRGCG0g7ogfW2WaPjfE",
    "ホテルフォルツァ博多駅博多口様": "18nlMjLbjxAIbobBquPRqeJKMCRGCG0g7ogfW2WaPjfE",

    "ホテルフォルツァ大阪北浜様": "1GM1vv1DlSG9z9JlvwFzGWT_UXHzTXSXPeBjv5pwcuuQ",
    
    "ホテルフォルツァ大阪なんば道頓堀様": "1L79RUIlKLTlXNzpV9Q85IKcd6PTP7VXDhuL53GCcI_0",

    "博多めんちゃんこ亭 天神店様": "1ISvbbsebm_sDnO_qa8ZznVctCGe3Xu0aAnco6oLjxDo",

    "カメリア様": "1ZgXKLbMov1s5ihWQdckBeMtDU05HUbGcO_GqHxnMJQI",

    "🔵デリバッグ様": "1ep0XITC6ru4_1-xeJpiz5LMDVXY5sKGSAG2Mg-NAVzU",

    "【藁焼き炉端と海鮮】博多海風土様": "1-Fn1nJcqNHYNRURBDo5rMk3s0WKB68Na_twVpgpvgXw",
    "【博多もつ鍋と焼き鳥】はかた風土本店様": "1-Fn1nJcqNHYNRURBDo5rMk3s0WKB68Na_twVpgpvgXw",
    "とり皮・中華そば かわのそば様": "1-Fn1nJcqNHYNRURBDo5rMk3s0WKB68Na_twVpgpvgXw",

    "焼肉割烹 YP流 宗右衛門町本店様": "18jqsmT7dh_OfkbyfIAFh-JlTmabSMmaZkgoVZS5rjjg",

    "🔵板前焼肉一光千日前店様": "1MypWeEWPfI21qiKOKmdyqLYmngvAUK77exRw6Xs0FTs",

    "マグロとご飯 黒田飯【博多店】様": "16u0Vzgdb5LliveifO7hZvAVrby_bgSIXIyJ7OwtxdLo",
    "マグロとご飯 黒田飯【福岡本店】様": "16u0Vzgdb5LliveifO7hZvAVrby_bgSIXIyJ7OwtxdLo",

    "ポーたま ももち浜店様": "1A6agu5yxEAHyoI1ZRRt8XEIWjcEm_OjQALhwi18neYo",
    "ポーたま 福岡赤坂店様": "1A6agu5yxEAHyoI1ZRRt8XEIWjcEm_OjQALhwi18neYo",
    "ポーたま 現代百貨店 板橋店": "1A6agu5yxEAHyoI1ZRRt8XEIWjcEm_OjQALhwi18neYo",

    "焼肉ニクゼン様": "1D2xCH-Cx_6JgVtzM__nl8VJrQUwchGwCUjiqSGYYr_Q",
    
    "リーガロイヤルホテル小倉様": "1JY-rsYa-dOk2hBFO6TmZT3v-nwjJ_SlPXVJgkmZhA-A",
    "小倉城周辺": "1JY-rsYa-dOk2hBFO6TmZT3v-nwjJ_SlPXVJgkmZhA-A",

    "大賀ドラッグストア 博多口店様": "17aLYprHg-nhtw3DXHh_3SW0KakwA8ncjPtbUXlpB3Y8",
    "大賀薬局 ソラリアステージ店様": "17aLYprHg-nhtw3DXHh_3SW0KakwA8ncjPtbUXlpB3Y8",

    "元祖肉肉うどん 川端店様": "1lsTdVkWfiz_uu-n0-w2RWkGh0SAYj7FVYjvyJh8InUo",
    "元祖肉肉うどん 中洲店様": "1lsTdVkWfiz_uu-n0-w2RWkGh0SAYj7FVYjvyJh8InUo",

    "博多焼肉ハチハチ博多本店様": "1YO-UtxEVivfRczGEj26AEb8yX5FcHnwylapoWix9gkA",
    "博多焼肉 HACHI HACHI 大名店様": "1YO-UtxEVivfRczGEj26AEb8yX5FcHnwylapoWix9gkA",
    "焼肉 HACHI HACHI ソラリア店様": "1YO-UtxEVivfRczGEj26AEb8yX5FcHnwylapoWix9gkA",

    "IMURI様": "1v318TLg-v2IOnNyKY8G_GKS9jbhgR7laRXXyHwqwAAk",

    "三井住友カード様": "10Vip2e7e1NFKmbYBCGQOLr7JQaZPBCOPk0CuKB5I1wE",

    "エスプリ様": "19bwHOHhq32yUHrBByYZT1fmwekL36uNRnCbIXYTzCsQ",

    "BiVi福岡様": "1mvZwxxuaK-WzwmXus40zbBx8VLHPJ5ayPurz41a3s4o",

    "京から鰻 赤坂店様": "1ImqtC2Tvgz6FVlj13Hzsh3nwUcSg6VBe8N5yAkAkiBU",
    "京から鰻 中洲店様": "1ImqtC2Tvgz6FVlj13Hzsh3nwUcSg6VBe8N5yAkAkiBU",

    "天神茶屋たばねのし様": "1TNBW9G54KaNWPyZZrH59ukdxo7HFyKBoalx7pVdWozY",

    "The Abram Aroma Shop様": "11uHpGMs9cdlwVXiiDQz6SxG0RdsbDvo5h1eN50gJWMs",

    "antiage本店サロン様": "1XufW3qcqQl3zksRTND3P-1hJSgEcAqYATk78DFKs940",

    "娘娘麻辣湯 中洲店様": "18c4cyTsYLPBsH76THJECL9mKaZo1nOfdohlCUcP6-3g",
    "娘娘麻辣湯 博多駅前店様": "18c4cyTsYLPBsH76THJECL9mKaZo1nOfdohlCUcP6-3g",

    "まわる寿司 博多魚がし 博多1番街店様": "126ctOM_kOFeRCiBM-gFLG_7ZYFc50jt6JZmlewnH2XE",

    "和牛ホルモン焼きジャック西成店様": "11q8S0cacxoHz8HgfTLx_142js9tUQ7XLiJd243C1OHM",

    "蟹かに城様": "1BkMWYb3Kj2KqEzBwLkMmyF_7TPEP-Ugk46iCgkXnU5w",
    "とらふぐ城様": "1BkMWYb3Kj2KqEzBwLkMmyF_7TPEP-Ugk46iCgkXnU5w",

    "黒門インフォメーションセンター様": "1ayiQxwuPdWus4F072kUoPyoFEWPW9YtSS0VTrcpKDIQ",

    "天神イナチカ様": "1I4bEDo5xHMIcQHPZunyrH9fOL5B7N11kTInslzXGraw",

    "CANAL Gourmet Street KUOHKA(くおうか)様": "1Kr9SLxqkbCW-56JDS-uMJf2LeRX5UrOzoWk0zCmV1t4",

    "炭焼肉 ステーキ irori by DAISUKE TAKUBO様": "1_WeMtHPCF_7jBYRa9kCXKNN_T9obRA3lAfx_mkjd9-o",

    "名代 宇奈とと 大名店様": "1-wzSBUnRWj2WwtDmxxe0ZmDlvKPrNyXhBSremY4oNHw",

    "天ぷら 雨のち晴れ 天神店様": "1GJzwt_aKGZGUjPNU5h12OEalRgij7urfblyipLUX628",

    "焼肉長介様": "1HVVqi9hTF3wAre5iSMd3n6CncEkhCHInmUahhK7fl9g",

    "土鍋と肴 海彩 miiro様": "1HdJVgaJ_ReZYWMlPCpn9Fbz3C8ZeNSXyvRYJodOirU0",

    "希代もつ鍋 万作家様": "1HjxaInfmrZoybbGBFfFCjQafCYQLiCrQ8rTNtYGRBf4",

    "鮨みつ様": "1IQBVmRCXW-YB9O4jwgAaGcCCtQeSIRXxhz5GKV4Rjt8",

    "居酒屋 天晴れ！輝想天外 博多中洲本店 海鮮×和牛様": "1JMn0nCKvbvx3XutNhnzLEWCGQBsEakWNXbdIQfGQJ78",


    "スタミナ山様": "1JtZ9Ddoej-GIoB1tg_CJuCR3dRTQWBhBAdQDkkz4G-U",
    "グルトベース 三条店様": "1JtZ9Ddoej-GIoB1tg_CJuCR3dRTQWBhBAdQDkkz4G-U",


    "コサエルコーヒーアンドチーズケーキ": "112MyvK9PEFZyK0DyWJSxlF3Zzi0VOLMpfd2wQ3hveU4",


  }
};
