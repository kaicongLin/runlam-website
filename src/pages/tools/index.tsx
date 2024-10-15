import React, { useEffect, useRef } from "react";
import Layout from "@theme/Layout";
import { UploadOutlined } from "@ant-design/icons";
import type { UploadProps } from "antd";
import { Button, message, Table, Upload } from "antd";
import { IDeliveryItems, BasicInfoFieldType } from "./intefaces";
import { deliveryExcelCloumnsMap } from "./data";
import styles from "./index.module.css";
import BasicInfo from "./BasicInfo";
import { readDeliveryExcel, downloadFile } from "./utils";
import _ from "lodash";
import { exportLSPackingList } from "./exportPackingList";

interface basicInfoRefProps {
  getFieldsValue: () => BasicInfoFieldType;
}

export default function Hello() {
  const [data, setData] = React.useState<IDeliveryItems[]>([]);
  const [columns, setColumns] = React.useState<any[]>([]);
  const basicInfoRef = useRef<basicInfoRefProps>();

  const props: UploadProps = {
    name: "file",
    multiple: true,
    accept:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel",
    customRequest: async (options: any) => {
      const { onSuccess, file } = options;
      onSuccess({ status: "success", response: file });
    },
    onChange: async (info) => {
      const { status } = info.file;
      if (status !== "uploading") {
        console.log("uploading", info.file, info.fileList);
      }
      if (status === "done") {
        message.success(`${info.file.name} file uploaded successfully.`);

        const res = await readDeliveryExcel(
          info.fileList.map((file) => file.originFileObj)
        );
        setData(res);
      } else if (status === "error") {
        message.error(`${info.file.name} file upload failed.`);
      }
    },
  };

  const handleExportPackingList = async () => {
    const basicValues = basicInfoRef.current?.getFieldsValue();
    // 生成新的 Excel 文件并下载
    const buffer = await exportLSPackingList(basicValues, data);
    const totalRollNumber = data.reduce(
      (acc, cur) => acc + cur.deliveryList.length,
      0
    );
    downloadFile(
      buffer,
      `${basicValues.invoice_no}装箱单${totalRollNumber}卷.xlsx`
    );
  };

  useEffect(() => {
    // 遍历对象的键值对
    const columns = [];
    for (const [key, val] of Object.entries(deliveryExcelCloumnsMap)) {
      columns.push({
        title: key,
        dataIndex: val,
        key: val,
      });
    }
    setColumns(columns);
  }, []);

  const dataSource = _.flatMap(data, (item) => item.deliveryList);
  return (
    <Layout title="Hello" description="Hello React Page">
      <div className={styles["delivery-wrap"]}>
        <div className={styles["upload-button"]}>
          <Upload {...props}>
            <Button icon={<UploadOutlined />}>请上传送货单</Button>
          </Upload>
        </div>

        <Table columns={columns} dataSource={dataSource} />

        <div className={styles['export-warp']}>
          <div className={styles['excel-preview']}></div>
          <BasicInfo
            ref={basicInfoRef}
            onExportPackingList={handleExportPackingList}
          />
        </div>
      </div>
    </Layout>
  );
}
