export const sheetName = '送货单明细';
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
