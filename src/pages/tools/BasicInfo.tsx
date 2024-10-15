import { Input, Button, Form } from "antd";
import React from "react";
import { BasicInfoFieldType } from './intefaces';

type BasicInfoRefProps = {
  getFieldsValue: () => BasicInfoFieldType;
};

interface IBasicInfoProps {
  onExportPackingList: () => void;
}

const BasicInfo = React.forwardRef<BasicInfoRefProps, IBasicInfoProps>(
  ({ onExportPackingList }, ref) => {
    const [form] = Form.useForm<BasicInfoFieldType>();

    React.useImperativeHandle(ref, () => ({
      getFieldsValue,
    }));

    const getFieldsValue = () => {
      return form.getFieldsValue();
    };

    return (
      <div>
        <Form
          form={form}
          name="basic"
          layout="horizontal"
          labelCol={{ span: 8 }}
          wrapperCol={{ span: 16 }}
          style={{ maxWidth: 400 }}
        >
          <Form.Item<BasicInfoFieldType>
            label="invoice no."
            name="invoice_no"
            rules={[
              { required: true, message: "Please input your invoice no.!" },
            ]}
          >
            <Input />
          </Form.Item>

          <Form.Item<BasicInfoFieldType>
            label="packing list no."
            name="packing_list_no"
            rules={[
              {
                required: true,
                message: "Please input your packing list no.!",
              },
            ]}
          >
            <Input />
          </Form.Item>

          <Form.Item<BasicInfoFieldType>
            label="No. of Ctn"
            name="case_number"
            rules={[
              { required: true, message: "Please input your No. of Ctn!" },
            ]}
          >
            <Input />
          </Form.Item>
          <div style={{ textAlign: "right" }}>
            <Button type="primary" onClick={onExportPackingList}>
              导出装箱单
            </Button>
          </div>
        </Form>
      </div>
    );
  }
);

export default BasicInfo;
