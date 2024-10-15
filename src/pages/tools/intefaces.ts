// 系统导出的字段
export type IDeliveryItem = {
  /**
   * 送货单号
   */
  delivery_number?: string;
  /**
   * 客户编号
   */
  customer_number?: string;
  /**
   * 客户订单号
   */
  customer_order_number?: string;
  /**
   * 货号
   */
  product_number?: string;
  /**
   * 品名
   */
  product_name?: string;
  /**
   * 颜色编号
   */
  color_number?: string;
  /**
   * 颜色名称
   */
  color_name?: string;
  /**
   * 缸号
   */
  dyeing_vat_number?: string;
  /**
   * 卷号
   */
  roll_number?: string;
  /**
   * 数量
   */
  quantity?: string;
  /**
   * 净总(KG)
   */
  net_total_kg?: string;
  /**
   * 扣库存数
   */
  deducted_stock_number?: string;
  /**
   * 单价
   */
  unit_price?: string;
  /**
   * 附件费
   */
  accessory_fee?: string;
  /**
   * 单位
   */
  unit?: string;
  /**
   * 客户货号
   */
  customer_product_number?: string;
  /**
   * 客户色号
   */
  customer_color_number?: string;
  /**
   * 备注
   */
  remark?: string;
  /**
   * 副宽
   */
  sub_width?: string;
  /**
   * 克重
   */
  gram_weight?: string;
  /**
   * 仓位
   */
  storage_location?: string;
};

export interface IDeliveryItems extends IDeliveryItem {
  name?: string;
  deliveryList?: IDeliveryItem[];
}

export type BasicInfoFieldType = {
  invoice_no?: string;
  packing_list_no?: string;
  case_number?: string;
};
