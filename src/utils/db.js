const DEFAULT_MAPPINGS = {
  "QC10001-R01": "樣品卡",
  "QC10002-R01": "入庫品檢驗收單",
  "QC10002-R02": "原物料品檢表",
  "QC10002-R03": "原料進料審查規格",
  "QC10004-R02": "QUALITY INSPECTION PLAN RECORD",
  "QC10005-R01": "押出機每日巡檢表",
  "QC10006-R01": "裝配對樣巡檢記錄表",
  "QC10006-R02": "半成品品檢表",
  "QC10007-R01": "完成品品檢表(首頁)",
  "QC10007-R02": "完成品品檢表(續頁)",
  "QC10007-R03": "零組件入庫品檢表(射出零件品檢表?)",
  "QC10008-R02": "出貨檢驗報告"
};

const STORAGE_KEY = "qc_form_mappings";

export const getMappings = () => {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (!stored) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(DEFAULT_MAPPINGS));
    return DEFAULT_MAPPINGS;
  }
  try {
    return JSON.parse(stored);
  } catch (e) {
    console.error("Failed to parse QC mappings from localStorage", e);
    return DEFAULT_MAPPINGS;
  }
};

export const saveMappings = (mappings) => {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(mappings));
};

export const resetMappings = () => {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(DEFAULT_MAPPINGS));
  return DEFAULT_MAPPINGS;
};
