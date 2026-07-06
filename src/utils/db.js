const DEFAULT_MAPPINGS = {
  "QC10002-R02": "原物料品檢表",
  "QC10004-R02": "QIP",
  "QC10006-R01": "裝配對樣巡檢記錄表",
  "QC10006-R02": "半成品品檢表",
  "QC10007-R01": "完成品品檢表(首頁)",
  "QC10007-R02": "完成品品檢表(續頁)",
  "QC10007-R03": "零組件入庫品檢表",
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
    const parsed = JSON.parse(stored);
    // 智慧檢查：若包含舊版已廢棄的預設鍵，自動重設為最新預設值
    if (parsed && (parsed["QC10001-R01"] !== undefined || parsed["QC10002-R01"] !== undefined || parsed["QC10002-R03"] !== undefined)) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(DEFAULT_MAPPINGS));
      return DEFAULT_MAPPINGS;
    }
    return parsed;
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
