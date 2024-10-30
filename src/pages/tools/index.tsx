import React, { useEffect, useRef } from "react";
import Layout from "@theme/Layout";
import ExcelJS from "exceljs";
import { UploadOutlined } from "@ant-design/icons";
import type { UploadFile, UploadProps } from "antd";
import { Button, message, Table, Upload } from "antd";
import { IDeliveryItems, BasicInfoFieldType } from "./intefaces";
import { deliveryExcelCloumnsMap } from "./data";
import styles from "./index.module.css";
import BasicInfo from "./BasicInfo";
import { readDeliveryExcel, downloadFile } from "./utils";
import _ from "lodash";
import { generateLSPackingList } from "./exportPackingList";
import UniverSheet from "./UniverSheet";

interface basicInfoRefProps {
  getFieldsValue: () => BasicInfoFieldType;
}

export default function Hello() {
  const [data, setData] = React.useState<IDeliveryItems[]>([]);
  const [columns, setColumns] = React.useState<any[]>([]);
  const [fileList, setFileList] = React.useState<UploadFile[]>([]);
  const [buffer, setBuffer] = React.useState<ArrayBuffer>();
  const [workbook, setWorkbook] = React.useState<ExcelJS.Workbook | null>(null);

  const fileCount = useRef<number>(0);
  const basicInfoRef = useRef<basicInfoRefProps>();

  const props: UploadProps = {
    name: "file",
    multiple: true,
    showUploadList: false,
    fileList: fileList,
    accept:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel",
    customRequest: async (options: any) => {
      const { onSuccess, file } = options;
      onSuccess({ status: "success", response: file });
    },
    onChange: async (info) => {
      // 记录已上传的文件数量
      fileCount.current++;
      // 当所有文件都上传完成后，读取 Excel 文件
      if (fileCount.current === info.fileList.length) {
        const res = await readDeliveryExcel(
          info.fileList.map((file) => file.originFileObj)
        );
        setData(res);
        setFileList(info.fileList);
        message.success(
          `total ${info.fileList.length} file uploaded successfully.`
        );

        generate(res);
      }
    },
  };

  /**
   *  生成装箱单模板
   */
  const generate = async (newData?: IDeliveryItems[]) => {
    const basicValues = basicInfoRef.current?.getFieldsValue();
    // 生成新的 装箱单的 workbook 对象
    const workbook = await generateLSPackingList(basicValues, newData ?? data)
    const buffer = await workbook.xlsx.writeBuffer();
    setBuffer(buffer);
    setWorkbook(workbook);
  };


  /**
   *  导出装箱单
   */
  const handleExportPackingList = async () => {
    const basicValues = basicInfoRef.current?.getFieldsValue();
    // workbook 转成 buffer 文件并下载
    const totalRollNumber = data.reduce(
      (acc, cur) => acc + cur.deliveryList.length,
      0
    );
    downloadFile(
      buffer,
      `${basicValues.invoice_no}装箱单${totalRollNumber}卷.xlsx`
    );
  };

  /**
   * 清空数据
   */
  const handleClearData = () => {
    setData([]);
    setFileList([]);
    fileCount.current = 0;
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
    // handleExportPackingList();
    generate();
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
        <Button onClick={handleClearData}>清空上传数据</Button>

        <Table columns={columns} dataSource={dataSource} />

        <div className={styles["export-warp"]}>
          <div className={styles["excel-preview"]}>
            <UniverSheet buffer={buffer} workBook={workbook} />
          </div>
          <BasicInfo
            ref={basicInfoRef}
            onExportPackingList={handleExportPackingList}
          />
        </div>
      </div>
    </Layout>
  );
}
