import { Input, Form, InputNumber } from "antd";
import React from "react";
import { BasicInfoFieldType } from "./_intefaces";
import styles from "./index.module.scss";

export type BasicInfoRefProps = {
  getFieldsValue: () => BasicInfoFieldType;
};

interface IBasicInfoProps {
  onFormValueChange: () => void;
}

const BasicInfo = React.forwardRef<BasicInfoRefProps, IBasicInfoProps>(
  ({ onFormValueChange }, ref) => {
    const [form] = Form.useForm<BasicInfoFieldType>();

    React.useImperativeHandle(ref, () => ({
      getFieldsValue,
    }));

    const getFieldsValue = () => {
      return form.getFieldsValue();
    };

    return (
      <div className={styles["basic-info-wrap"]}>
        <Form
          form={form}
          name="basic"
          layout="horizontal"
          className={styles["basic-info-form"]}
          labelCol={{ span: 8 }}
          wrapperCol={{ span: 16 }}
          onChange={onFormValueChange}
        >
          <div className={styles["block-title"]}>基础信息字段</div>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="invoice no."
            name="invoice_no"
            initialValue="LS20241007sea"
            rules={[{ required: true, message: "Please input invoice no.!" }]}
          >
            <Input />
          </Form.Item>

          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="packing list no."
            name="packing_list_no"
            initialValue="R037/24"
            rules={[
              {
                required: true,
                message: "Please input packing list no.!",
              },
            ]}
          >
            <Input />
          </Form.Item>

          <div className={styles["block-title"]}>表格列字段</div>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="箱號"
            name="case_number"
            tooltip="输入初始箱号，如：33S37-24-11384, 后面行会自动递增（x-x-11385， x-x-1386...）"
            initialValue="33S37-24-11384"
            rules={[
              {
                required: true,
                pattern: /^[^\s-]+-[^\s-]+-\d+$/,
                message: "请输入格式为 x-x-x 的内容, 最后x需要为数字",
              },
            ]}
          >
            <Input />
          </Form.Item>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="箱/卷數"
            name="roll_qty"
            tooltip="导出该列的所有数据都会统一使用这个值，若有个别不同需要手动修改"
            initialValue="1"
            rules={[{ required: true, message: "Please input roll qty!" }]}
          >
            <InputNumber style={{ width: "100%" }} />
          </Form.Item>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="單位"
            name="unit"
            tooltip="导出该列的所有数据都会统一使用这个值，若有个别不同需要手动修改"
            initialValue="ROLL"
            rules={[{ required: true, message: "Please input unit!" }]}
          >
            <Input />
          </Form.Item>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="合同號"
          >
            <div>自动匹配送货单的'客户订单号'</div>
          </Form.Item>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="成份"
          >
            <div>根据送货单的'货号'自动匹配</div>
          </Form.Item>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="落貨物料"
          >
            <div>自动匹配送货单的'货号'</div>
          </Form.Item>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="顏色"
          >
            <div>自动匹配送货单的'颜色名称 '</div>
          </Form.Item>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="封度/實用"
            name="width"
            tooltip="导出该列的所有数据都会统一使用这个值，若有个别不同需要手动修改"
            initialValue="150CM"
            rules={[{ required: true, message: "Please input width!" }]}
          >
            <Input />
          </Form.Item>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="數量"
          >
            <div>自动匹配送货单的'数量 '</div>
          </Form.Item>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="單位"
            tooltip="Y 会 自动转成 YD，其他不变"
          >
            <div>自动匹配送货单的'单位 '</div>
          </Form.Item>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="淨重"
          >
            <div>自动匹配送货单的'净总(KG) '</div>
          </Form.Item>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="毛重"
          >
            <div>自动匹配送货单的'净总(KG) + 0.5'</div>
          </Form.Item>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="尺寸 (cm)"
            name="meas"
            tooltip="导出该列的所有数据都会统一使用这个值，若有个别不同需要手动修改"
            initialValue="26*26*153"
            rules={[{ required: true, message: "Please input meas!" }]}
          >
            <Input />
          </Form.Item>
          <Form.Item<BasicInfoFieldType>
            className={styles["mb-10"]}
            label="備註"
            tooltip="生成规则：客户订单号，货号，颜色编号，颜色名称， 缸号，共'{'送货单卷数'}'卷
              报告 OK"
          >
            <div>自动生成</div>
          </Form.Item>
        </Form>
      </div>
    );
  }
);

export default BasicInfo;
