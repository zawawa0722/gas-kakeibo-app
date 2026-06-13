//文字列定義
const CANCEL = "キャンセル"

// メールアドレス定義
const EMAIL_NAOTO = PropertiesService.getScriptProperties().getProperty("EMAIL_NAOTO");
const EMAIL_SARA = PropertiesService.getScriptProperties().getProperty("EMAIL_SARA");

// メールメッセージ定義
const SAVINGS_MESSAGES = {
  high: "貯金率が30%を超えました！すごい！来月もこの調子で頑張りましょう！",
  midHigh: "貯金率が20%台です！来月もこの調子で30%以上を目指しましょう！",
  midLow: "貯金率が10%台です。来月は20%台を目指しましょう。",
  low: "貯金率が10%未満はちょっとまずいです。。来月は節約生活じゃー！！！"
};

// メール関連定義
const SUBJECT_FIVE_DAY = "CSVファイル配置のお願い";
const SUBJECT_TWENTY_SIX_DAY = "CSVファイル配置のお願い";
const SUBJECT_TWENTY_FIVE_DAY = "家族口座へ振込のお願い";
const BODY_FVE_DAY = "B43の内容を確認して、GoogleDriveにCSVファイルを配置してください。"; // メールの本文
const REQUEST_BUCK_UP = "各シートのバックアップを取得し、内容をクリアしてください。"; // 1/26に使用する

// シート名
const SHEET_THOUBOKANRI = "帳簿管理"
const SHEET_SYUNYUKANRI = "収入管理"
const SHEET_SISYUTUKANRI = "支出管理"
const SHEET_ONLINE = "オンライン明細"
const SHEET_THOUBOKANRI_BK = "帳簿管理履歴"
const SHEET_SYUNYUKANRI_BK = "収入管理履歴"
const SHEET_SISYUTUKANRI_BK = "支出管理履歴"
const SHEET_ONLINE_BK = "オンライン明細履歴"

// 各シートのグラフ出力時の定数
const CELL_POSITIONS = {
  [SHEET_THOUBOKANRI]: {
    yearCell: "G15",
    monthCell: "H15",
    labelRange: "A4:A9",
    dateRange: "B3:M3",
    chartPosition: "A22"
  },
  [SHEET_SISYUTUKANRI]: {
    yearCell: "A29",
    monthCell: "B29",
    labelRange: "A4:A24",
    dateRange: "B3:M3",
    chartPosition: "A31"
  },
  [SHEET_SYUNYUKANRI]: {
    yearCell: "A15",
    monthCell: "B15",
    labelRange: "A4:A10",
    dateRange: "B3:M3",
    chartPosition: "A17"
  },
  [SHEET_ONLINE]: {
    yearCell: "A13",
    monthCell: "B13",
    labelRange: "A4:A8",
    dateRange: "B3:M3",
    chartPosition: "A15"
  }
};

// 各シートのバックアップ範囲
const BACKUP_RANGE = {
  [SHEET_THOUBOKANRI]: {
    backupRange: "A3:N12"
  },
  [SHEET_SYUNYUKANRI]: {
    backupRange: "A3:N11"
  },
  [SHEET_SISYUTUKANRI]: {
    backupRange: "A3:N25"
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
    backupRange: ["B5:M10"]
  },
  [SHEET_SISYUTUKANRI]: {
    backupRange: ["B5:M7", "B9:M12", "B14:M24"]
  },
  [SHEET_ONLINE]: {
    backupRange: ["B4:M8"]
  }
};

// 支出カテゴリ
const EXPEND_CATEGORY = {
  DENKI: "電気料金",
  SUIDOU: "水道料金",
  GASS: "ガス料金",
  SYOKUHI: "食品",
  GAISYOKU: "外食",
  SEIKATU: "生活用品",
  TEMP_N: "直人一時負担",
  TEMP_S: "沙羅一時負担",
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
  OTHER: "その他"
};

const ONLINE_CATEGORY = {
  AMAZON: "ＡＭＡＺＯＮ",
  RAKUTEN: ["ﾗｸﾃﾝｲﾁﾊﾞ", "楽天"],
  NITORI: "ニトリネット",
  MUJIRUSI: "無印良品",
  OTHER_ONLINE: "その他"
};


// オンラインカテゴリ
const ONLINE_MEISAI_CATEGORY = {
  AMAZON: "Amazon",
  RAKUTEN: "楽天市場",
  NITORI: "ニトリ",
  MUJIRUSI: "無印良品",
  NETFLIX: "Netflix",
  EVERY: "Everyfrecious",
  OTHER_ONLINE: "その他"
};

// カードカテゴリ
const CARD_CATEGORY = {
  B43: "ペアカード",
  RAKUTEN_CARD: "楽天カード",
  VISA: "三井住友カード"
};

// 支出の計算
let calculateExpend = {
  electricity: 0,
  water: 0,
  gass: 0,
  food: 0,
  eateOut: 0,
  lifeItem: 0,
  tempNaoto: 0,
  tempSara: 0,
  sabsc: 0,
  entame: 0,
  hospital: 0,
  syumi: 0,
  koutuu: 0,
  kenkoubiyou: 0,
  division: 0,
  shopping: 0,
  loan: 0,
  otherexp: 0
};

// オンライン取引の計算
let calculateOnline = {
  amazon: 0,
  rakuten: 0,
  nitori: 0,
  mujirusi: 0,
  otheronl: 0
};

// 支出カテゴリと計算項目のマッピング
const EXPEND_MAP = {
  DENKI: "electricity",
  SUIDOU: "water",
  GASS: "gass",
  SYOKUHI: "food",
  GAISYOKU: "eateOut",
  SEIKATU: "lifeItem",
  SABSC: "sabsc",
  ENTAME: "entame",
  HOSPITAL: "hospital",
  SYUMI: "syumi",
  KOUTUU: "koutuu",
  KENKOUBIYOU: "kenkoubiyou",
  TEMP_N: "tempNaoto",
  TEMP_S: "tempSara",
  BUNKATU: "division",
  SHOPPING: "shopping",
  LOAN: "loan",
  OTHER: "otherexp"
};

// オンラインカテゴリと計算項目のマッピング
const ONLINE_MAP = {
  AMAZON: "amazon",
  RAKUTEN: "rakuten",
  NITORI: "nitori",
  MUJIRUSI: "mujirusi",
  NETFLIX: "netflix",
  EVERY: "every",
  OTHER_ONLINE: "otheronl"
};
