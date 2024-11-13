// 读取excel表 sheet的名称
export const SHEET_NAME = "送货单明细";
// 导出表格开始行
export const TABLE_START_ROW = 10;

// excel列字段映射
export const deliveryExcelCloumnsMap = {
  送货单号: "delivery_number",
  客户编号: "customer_number",
  客户订单号: "customer_order_number",
  货号: "product_number",
  品名: "product_name",
  颜色编号: "color_number",
  颜色名称: "color_name",
  缸号: "dyeing_vat_number",
  卷号: "roll_number",
  数量: "quantity",
  "净总(KG)": "net_total_kg",
  扣库存数: "deducted_stock_number",
  单价: "unit_price",
  附件费: "accessory_fee",
  单位: "unit",
  客户货号: "customer_product_number",
  客户色号: "customer_color_number",
  备注: "remark",
  副宽: "sub_width",
  克重: "gram_weight",
  仓位: "storage_location",
};

export const deliveryExcelColumns = [
  { name: "箱號", englishName: "No. of Ctn", key: "case_number", width: 18 },
  { name: "箱/卷數", englishName: "Ctn/Roll Qty.", key: "roll_qty", width: 12 },
  { name: "單位", englishName: "Unit", key: "unit", width: 26 },
  { name: "合同號", englishName: "Order No", key: "order_no", width: 18 },
  { name: "成份", englishName: "Compostion", key: "compostion", width: 16 },
  {
    name: "落貨物料",
    englishName: "Description",
    key: "description",
    width: 16,
  },
  { name: "顏色", englishName: "Color", key: "color", width: 24 },
  { name: "封度/實用", englishName: "Width", key: "width", width: 12 },
  { name: "數量", englishName: "Qty", key: "qty", width: 12 },
  { name: "單位", englishName: "Unit", key: "length_unit", width: 12 },
  { name: "淨重", englishName: "N/W", key: "net_weight", width: 12 },
  { name: "毛重", englishName: "G/W", key: "gross_weight", width: 12 },
  { name: "尺寸 (cm)", englishName: "Meas.", key: "meas", width: 12 },
  { name: "備註", englishName: "Remark", key: "remark", width: 30 },
];

export const productNumberToCompostionMap = {
  "N3333-140R": "Recycle Nylon87% Spandex13%",
  "N3333-140P": "Nylon87% Spandex13%",
  "N3333-140": "Nylon87% Spandex13%",
  "N3333-130R": "Recycle Nylon87% Spandex13%",
  "N3333-130P": "Nylon87% Spandex13%",
  "N3333-130C": "Nylon87% Spandex13%",
  "N3333-130": "Nylon87% Spandex13%",
  "N3350-165R": "Recycle Nylon 82% Spandex 18%",
  "N3350-160": "Nylon82% Spandex18%",
  "N3350-150R": "Recycle Nylon 82% Spandex 18%",
  "N3350-150P": "Nylon82% Spandex18%",
  "N3350-150": "Nylon82% Spandex18%",
  "N3350-140R": "Recycle Nylon 82% Spandex 18%",
  "N3350-135": "Nylon79% Spandex21%",
  "N3350-130V": "Nylon82% Spandex18%",
  "N3350-130R": "Recycle Nylon 82% Spandex 18%",
  "N3350-130P": "Nylon82% Spandex18%",
  "N3350-130F": "Nylon82% Spandex18%",
  "N3350-130BIO": "BIO PA56 82% Spandex18%",
  "N3350-130": "Nylon82% Spandex18%",
  "N3350-125P": "Nylon82% Spandex18%",
  "N3350-125": "Nylon82% Spandex18%",
  N3328V: "Nylon90% Spandex10%",
  N3328RP: "Re Nylon90% Spandex10%",
  N3328R: "Recycle Nylon90% Spandex10%",
  N3328P: "Nylon90% Spandex10%",
  N3328C: "Nylon90% Spandex10%",
  "N3328-180": "Nylon90% Spandex10%",
  "N3328-170P": "Nylon90% Spandex10%",
  "N3328-170": "Nylon90% Spandex10%",
  "N3328-150P": "Nylon 90% Spandex 10%",
  "N3328-150": "Nylon90% Spandex10%",
  "N3328-140R": "Recycled Nylon 92% Spandex 8%",
  "N3328-140P": "Nylon90% Spandex10%",
  "N3328-140": "Nylon90% Spandex10%",
  "N3328-135": "Nylon90% Spandex10%",
  "N3328-120": "Nylon90% Spandex10%",
  N3328: "Nylon90% Spandex10%",
  N4387R: "Recycle Nylon76% Spandex24%",
  N4387P: "Nylon76% Spandex24%",
  "N4387-200R": "Recycle Nylon76% Spandex24%",
  "N4387-200P": "Nylon76% Spandex24%",
  "N4387-200DR": "Recycle Nylon76% Spandex24%",
  "N4387-200DP": "Nylon76% Spandex24%",
  "N4387-200D": "Nylon76% Spandex24%",
  "N4387-200A": "Nylon76% Spandex24%",
  "N4387-200": "Nylon76% Spandex24%",
  "N4387-190": "Nylon76% Spandex24%",
  "N4387-180": "Nylon76% Spandex24%",
  "N4387-170P": "Nylon76% Spandex24%",
  "N4387-170F": "Nylon76% Spandex24%",
  "N4387-170": "Nylon76% Spandex24%",
  "N4387-160": "Nylon75% Spandex25%",
  "N4387-1": "Nylon76% Spandex24%",
  N4387: "Nylon76% Spandex24%",
  "N1066R-125": "Recycle Nylon 67% Spandex 33%",
  N1066R: "Recycle Nylon 67% Spandex 33%",
  N1066P: "Nylon 67% Spandex 33%",
  N1066F: "Nylon 67% Spandex 33%",
  "N1066-140R": "Recycle Nylon 67% Spandex 33%",
  "N1066-140": "Nylon 67% Spandex 33%",
  "N1066-130E": "Nylon 67% Spandex 33%",
  "N1066-130": "Nylon 67% Spandex 33%",
  N1066: "Nylon 67% Spandex 33%",
  N1018RP: "Recycle Nylon69%,Spandex31%",
  N1018RM2: "Recycle Nylon69%,Spandex31%",
  N1018RL: "Recycle Nylon 69%,Lycra 31%",
  N1018R: "Recycle Nylon 69%,Spandex31%",
  N1018P: "Nylon69%,Spandex31%",
  N1018M2: "Nylon69%,Spandex31%",
  N1018L: "Nylon 69%,Lycra 31%",
  N1018: "Nylon69%,Spandex31%",
  N1012RM: "Recycle Nylon 57%,Spandex 43%",
  N1012R: "Recycle Nylon 57%,Spandex 43%",
  N1012M: "Nylon 57%,Spandex 43%",
  "N1012-250": "Nylon 57%,Spandex 43%",
  "N1012-230M": "Nylon 57%,Spandex 43%",
  "N1012-230": "Nylon 57%,Spandex 43%",
  "N1012-200R": "Recycle Nylon 57%,Spandex 43%",
  "N1012-200H": "Nylon 57%,Spandex 43%",
  "N1012-200BIO": "BIO Nylon（PA56) 57% Spandex 43%",
  "N1012-200": "Nylon 57%,Spandex 43%",
  "N1012-190R": "Recycle Nylon 57%,Spandex 43%",
  "N1012-190": "Nylon 57%,Spandex 43%",
  "N1012-185R": "Recycle Nylon 57%,Spandex 43%",
  "N1012-185": "Nylon 57%,Spandex 43%",
  "N1012-1": "Nylon 57%,Spandex 43%",
  N1012: "Nylon 57%,Spandex 43%",
  W2310L: "Nylon85% Lycra 15%",
  "W2310-210": "Nylon 85% Spandex15%",
  "W2310-200": "Nylon 85% Spandex15%",
  "W2310-190M": "Nylon 85% Spandex15%",
  "W2310-190": "Nylon 85% Spandex15%",
  "W2310-180": "Nylon 85% Spandex15%",
  W2310: "Nylon85% Spandex15%",
  "N1092-190": "Nylon 59 % Spandex 41 %",
  N1092: "Nylon 59 % Spandex 41 %",
  N4526RP: "Recycle Nylon 67% Spandex 33%",
  N4526R: "Recycle Nylon 67% Spandex 33%",
  N4526P: "Nylon 67% Spandex 33%",
  "N4526-1P": "Nylon 67% Spandex 33%",
  "N4526-1": "Nylon 67% Spandex 33%",
  N4526: "Nylon 67% Spandex 33%",
};