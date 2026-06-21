// 文字列定義
const CANCEL = "キャンセル"
const TORIGGER = "トリガ実行"
const MANUAL = "手動実行"

// スクリプトプロパティ定義変数
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_ACCESS_TOKEN');
const API_OPENAI = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
const EMAIL_PRIMARY = PropertiesService.getScriptProperties().getProperty("EMAIL_PRIMARY");
const EMAIL_SECONDARY = PropertiesService.getScriptProperties().getProperty("EMAIL_SECONDARY");

// メール関連定義
const SUBJECT_REQUEST_CSV_STORAGE = "CSVファイル配置のお願い";
const SUBJECT_REQUEST_BUCKUP = "前年度実績のバックアップ依頼";
const SUBJECT_FAILED_IMPORT = "CSVファイルのインポートエラー通知";
const SUBJECT_SCCEED_IMPORT = "CSVファイルのインポート完了通知";
const BODY_REQUEST_CSV_STORAGE = "ワンバンクの内容を確認して、GoogleDriveにCSVファイルを配置してください。";
const BODY_REQUEST_BUCKUP = "各シートのバックアップを取得し、内容をクリアしてください。";
const BODY_FAILED_IMPORT = "CSVファイルのインポートに失敗しました。ソースまたはファイルを確認してください。"; 
const BODY_SCCEED_IMPORT = "CSVファイルのインポートが完了しました。"; 
const BODY_SCCEED_IMPORT_INCLUDE_ERROR = "CSVファイルのインポートが完了しました。\n" + "また、WARNING または ERRORを検知しました。\n" + "ログシートを確認してください。";

// シート名
const SHEET_THOUBOKANRI = "帳簿管理"
const SHEET_SYUNYUKANRI = "収入管理"
const SHEET_SISYUTUKANRI = "支出管理"
const SHEET_ONLINE = "オンライン明細"
const SHEET_LOG = "ログ"
const SHEET_SETTING = "設定"
const SHEET_THOUBOKANRI_BK = "帳簿管理履歴"
const SHEET_SYUNYUKANRI_BK = "収入管理履歴"
const SHEET_SISYUTUKANRI_BK = "支出管理履歴"
const SHEET_ONLINE_BK = "オンライン明細履歴"

// CSV関連
const NYUUKIN = "クレジットカードで入金"
const WANBANK = "ﾜﾝﾊﾞﾝｸ"
const COL_DESC = 1;     // 明細内容
const COL_OPERATOR = 2; // 操作者
const COL_TYPE = 3;     // メモ
const COL_AMOUNT = 4;   // 金額
const COL_CATEGORY = 5; // カテゴリ
const COL_VALID = 6;    // 有効フラグ
const COL_CARD = 7;     // カード種別

// 各シートのグラフ出力時の定数
const CELL_POSITIONS = {
   // 帳簿管理シートのグラフは自動更新されるため、出力不要
  /*[SHEET_THOUBOKANRI]: {
    yearCell: "A16",
    monthCell: "B16",
    labelRange: "A4:A9",
    dateRange: "B3:M3",
    totalResultRange: "A8:A10",
    chartPosition: "A30"
  },*/
  [SHEET_SISYUTUKANRI]: {
    yearCell: "A31",
    monthCell: "B31",
    labelRange: "A4:A26",
    labelRangeCustom: "A4:A25",
    dateRange: "B3:M3",
    chartPosition: "A33"
  },
  [SHEET_SYUNYUKANRI]: {
    yearCell: "A14",
    monthCell: "B14",
    labelRange: "A4:A9",
    dateRange: "B3:M3",
    chartPosition: "A16"
  },
  [SHEET_ONLINE]: {
    yearCell: "A14",
    monthCell: "B14",
    labelRange: "A4:A9",
    dateRange: "B3:M3",
    chartPosition: "A16"
  }
};

// 各シートのバックアップ範囲
const BACKUP_RANGE = {
  [SHEET_THOUBOKANRI]: {
    backupRange: "A3:N12"
  },
  [SHEET_SYUNYUKANRI]: {
    backupRange: "A3:N10"
  },
  [SHEET_SISYUTUKANRI]: {
    backupRange: "A3:N27"
  },
  [SHEET_ONLINE]: {
    backupRange: "A3:N9"
  }
};

const CLEAR_RANGE = {
  [SHEET_THOUBOKANRI]: {
    backupRange: ["B4:M5"]
  },
  [SHEET_SYUNYUKANRI]: {
    backupRange: ["B4:M9"]
  },
  [SHEET_SISYUTUKANRI]: {
    backupRange: ["B5:M7", "B9:M12", "B14:M26"]
  },
  [SHEET_ONLINE]: {
    backupRange: ["B4:M8"]
  }
};

// 支出カテゴリ
let EXPEND_CATEGORY = {
  DENKI: "電気料金",
  SUIDOU: "水道料金",
  GASS: "ガス料金",
  TUUSINHI: "通信費",
  SYOKUHI: "食品",
  KONBINI: "コンビニ",
  GAISYOKU: "外食",
  SEIKATU: "生活用品",
  TEMP_USER_A: "",
  TEMP_USER_B: "",
  BUNKATU: "分割費",
  SABSC: "サブスク",
  ENTAME: "エンタメ",
  HOSPITAL: "病院",
  SHOPPING: "ショッピング",
  SYUMI: "趣味",
  KOUTUU: "交通",
  KENKOUBIYOU: "健康/美容",
  ODEKAKE: "おでかけ",
  JUKYO: "住居",
  KYOUIKU: "教育",
  KOUKYOURYOUKIN: "公共料金",
  LOAN: "ローン",
  KOUSAIHI: "交際費",
  OTHER: "その他"
};

// CSVの明細に記載されている項目名
const ONLINE_CATEGORY = {
  AMAZON: ["ＡＭＡＺＯＮ",
           "Ａｍａｚｏｎ　Ｄｏｗｎｌｏａｄｓ"
          ],
  RAKUTEN: ["ﾗｸﾃﾝｲﾁﾊﾞ",
            "楽天",
            "サンドラッグｅ－ｓｈｏｐ",
            "ホット安心ストア",
            "ａ‐ｃｕｅｂｓｈｏｐ／エーキュ",
            "ホスピマート",
            "ＴＫ‐ＪＩＡＮＧ",
            "でん吉",
            "サンドラッグｅ‐ｓｈｏｐ",
            "Ｈａｍｅｅ（ハミィ）",
            "便利生活　マイルーム",
            "インテリアパレット",
            "Ｈａｐｐｙ　Ｅｖｅｒｙｄａｙ楽",
            "ＴＯＤＡＹ＆ＡＬＷＡＹＳ",
            "Ｈａｐｐｙ　Ｅｖｅｒｙｄａｙ楽",
            "スリーアールプラザ",
            "ＳＨＯＰＷＮＹ",
            "ＴＣＣ　Ｏｎｌｉｎｅ　Ｓｈｏｐ",
            "ＳＯＬＩＡ　ＳＨＯＰ",
            "中村商事",
            "ＤｅａｒＣａｒｄｓ（ディアカー",
            "松阪牛三重松良",
            "おいしい醤油・味噌【足立醸造】",
            "オーラルケアのＤＯＤ",
            "いまどき本舗",
            "ｂｌａｎｃｏ",
            "まごころ屋本店",
            "山善山屋オンラインショップ",
            "ｎｅｗｗａｖｅｓｔｏｒｅ",
            "【公式】いろはショップオンライ",
            "東栄Ｓｈｏｐ",
            "ギフト百花",
            "東海ストア",
            "地酒「作」＆全国銘酒専門べんの",
            "ＤＵＥＮ　ＳＨＯＰ",
            "ｓｈｉｎｓｏｕ商店",
            "ＵＬＴＯＲＡ　公式",
            "ＥＸＣＥＰＴＩＯＮ本店",
            "ｆｕｔｕｒａ　ｒｕｒｕ",
            "ｽｲﾂﾁﾎﾞﾂﾄｶﾌﾞｼｷｶｲｼﾔ",
            "アエトニクス　Ｒａｉｎ＆Ｏｕｔ",
            "赤ちゃんデパート",
            "ｒｅｌｉｅｆ１０",
            "ＭＯＮＯ　ＫＯＴＯ　ＤＥＰＴ．"
           ],
  NITORI: "ニトリネット",
  MUJIRUSI: "無印良品",
  TIKTOK: "TIKTOK SHOP",
  OTHER_ONLINE: "その他"
};

// オンラインカテゴリ
const ONLINE_MEISAI_CATEGORY = {
  AMAZON: "Amazon",
  RAKUTEN: "楽天市場",
  NITORI: "ニトリ",
  MUJIRUSI: "無印良品",
  TIKTOK: "TikTokShop",
  OTHER_ONLINE: "その他"
};

// カードカテゴリ
const CARD_CATEGORY = {
  WANBANK: "ペアカード",
  RAKUTEN_CARD: "楽天カード",
  AMAZON: "Amazon Mastercard / Amazon Prime Mastercard"
};

// 支出の計算
let calculateExpend = {
  electricity: 0,
  water: 0,
  gass: 0,
  communicationExpenses: 0,
  food: 0,
  convenience_store: 0,
  eateOut: 0,
  lifeItem: 0,
  temp_user_A: 0,
  temp_user_B: 0,
  sabsc: 0,
  entame: 0,
  hospital: 0,
  syumi: 0,
  koutuu: 0,
  kenkoubiyou: 0,
  odekake: 0,
  jyukyo: 0,
  education: 0,
  public_fee: 0,
  division: 0,
  shopping: 0,
  loan: 0,
  society: 0,
  otherexp: 0
};

// オンライン取引の計算
let calculateOnline = {
  amazon: 0,
  rakuten: 0,
  nitori: 0,
  mujirusi: 0,
  tiktok: 0,
  otheronl: 0
};

// 支出カテゴリと計算項目のマッピング
const EXPEND_MAP = {
  DENKI: "electricity",
  SUIDOU: "water",
  GASS: "gass",
  TUUSINHI: "communicationExpenses",
  SYOKUHI: "food",
  KONBINI: "convenience_store",
  GAISYOKU: "eateOut",
  SEIKATU: "lifeItem",
  SABSC: "sabsc",
  ENTAME: "entame",
  HOSPITAL: "hospital",
  SYUMI: "syumi",
  KOUTUU: "koutuu",
  KENKOUBIYOU: "kenkoubiyou",
  TEMP_USER_A: "temp_user_A",
  TEMP_USER_B: "temp_user_B",
  BUNKATU: "division",
  SHOPPING: "shopping",
  LOAN: "loan",
  KOUSAIHI: "society",
  ODEKAKE: "odekake",
  JUKYO: "jyukyo",
  KYOUIKU: "education",
  KOUKYOURYOUKIN: "public_fee",
  OTHER: "otherexp"
};

// オンラインカテゴリと計算項目のマッピング
const ONLINE_MAP = {
  AMAZON: "amazon",
  RAKUTEN: "rakuten",
  NITORI: "nitori",
  MUJIRUSI: "mujirusi",
  TIKTOK: "tiktok",
  OTHER_ONLINE: "otheronl"
};
