import React, { useEffect, useRef } from "react";
import Layout from "@theme/Layout";
import { UploadOutlined } from "@ant-design/icons";
import type { UploadFile, UploadProps } from "antd";
import { Button, message, Table, Upload } from "antd";
import { IDeliveryItems, BasicInfoFieldType } from "./intefaces";
import { deliveryExcelCloumnsMap } from "./data";
import styles from "./index.module.css";
import BasicInfo from "./BasicInfo";
import { readDeliveryExcel, downloadFile } from "./utils";
import _ from "lodash";
import { exportLSPackingList } from "./exportPackingList";
import LuckyExcel from "luckyexcel";

interface basicInfoRefProps {
  getFieldsValue: () => BasicInfoFieldType;
}

export default function Hello() {
  const [data, setData] = React.useState<IDeliveryItems[]>([]);
  const [columns, setColumns] = React.useState<any[]>([]);
  const [fileList, setFileList] = React.useState<UploadFile[]>([]);
  const [buffer, setBuffer] = React.useState<ArrayBuffer>();

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

        // 生成新的 Excel 文件
        const basicValues = basicInfoRef.current?.getFieldsValue();
        const buffer = await exportLSPackingList(basicValues, res);
        setBuffer(buffer);
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

  const handleClearData = () => {
    setData([]);
    setFileList([]);
    fileCount.current = 0;
  };

  const createExcel = () => {
    LuckyExcel.transformExcelToLucky(
      buffer,
      function (exportJson, lunckysheetfile) {
        console.log(exportJson);
        console.log(lunckysheetfile);
        console.log(window.luckysheet);
        if (window.luckysheet && window.luckysheet.create) {
          window.luckysheet?.create({
            container: "luckysheet", //luckysheet is the container id
            lang: "zh",
            allowCopy: false, //是否允许拷贝
            showtoolbar: false, //是否显示工具栏
            showinfobar: false, //是否显示顶部信息栏
            showsheetbar: false, //是否显示底部sheet页按钮
            showstatisticBar: false, //是否显示底部计数栏
            showstatisticBarConfig: {},
            enableAddRow: false, //是否允许添加行
            enableAddCol: false, //是否允许添加列
            // // showRowBar: false, // 是否显示行号区域
            // // showColumnBar: false, // 是否显示列号区域
            enableAddBackTop: false, //是否允需回到顶部
            sheetFormulaBar: false, //是否显示公示栏
            allowEdit: false, //是否允许编辑
            // rowHeaderWidth: 0, //纵坐标
            // columnHeaderHeight: 0, //横坐标
            // devicePixelRatio: 10, //设置比例
            data: exportJson.sheets,
            hook: {
              workbookCreateAfter: () => {
                console.log("workbookCreateAfter------------");
              },
            },
          });
        }
      }
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
        <Button onClick={handleClearData}>清空上传数据</Button>
        <Button onClick={createExcel}>初始化</Button>

        <Table columns={columns} dataSource={dataSource} />

        <div className={styles["export-warp"]}>
          <div className={styles["excel-preview"]}>
          <div id="luckysheet" className={styles['preiew-luckysheet']}></div>
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
