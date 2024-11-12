import React, { useEffect, useRef } from "react";
import Layout from "@theme/Layout";
import ExcelJS from "exceljs";
import {
  UploadOutlined,
  ClearOutlined,
  InfoCircleOutlined,
} from "@ant-design/icons";
import type { UploadFile, UploadProps } from "antd";
import { Button, message, Table, Upload, Affix } from "antd";
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
  const removeFlag = useRef<boolean>(false);

  const fileCount = useRef<number>(0);
  const basicInfoRef = useRef<basicInfoRefProps>();

  const props: UploadProps = {
    name: "file",
    multiple: true,
    // showUploadList: false,
    fileList: fileList,
    accept:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel",
    customRequest: async (options: any) => {
      const { onSuccess, file } = options;
      onSuccess({ status: "success", response: file });
      return {
        abort() {},
      };
    },
    onChange: async (info) => {
      if (removeFlag.current) {
        removeFlag.current = false;
        return;
      }
      fileCount.current++;
      const allDone = info.fileList.every((file) => file.status === "done");
      if (fileCount.current === info.fileList.length && !allDone) {
        const uploadingCount = info.fileList.filter(
          (f) => f.status === "uploading"
        ).length;
        const newFileList: UploadFile[] = info.fileList.map((f) => ({
          ...f,
          status: "done",
        }));
        setFileList(newFileList);
        fileCount.current -= uploadingCount;
      }

      // 读取 Excel 文件
      if (fileCount.current === info.fileList.length && allDone) {
        message.success(
          `total ${info.fileList.length} file load successfully.`
        );

        // 生成装箱单
        generate(info.fileList);
      }
    },
    onRemove: async (file) => {
      const newFileList = fileList.filter((f) => f.uid !== file.uid);
      setFileList(newFileList);
      message.success(`remove ${file.name} file successfully.`);
      fileCount.current--;
      removeFlag.current = true;
      // 重新生成装箱单
      generate(newFileList);
    },
  };

  /**
   *  生成装箱单模板
   */
  const generate = async (fileList?: UploadFile<any>[]) => {
    const basicValues = basicInfoRef.current?.getFieldsValue();
    let data = [];
    // 如果有文件列表，则读取文件内容
    console.log(fileList);

    if (fileList) {
      data = await readDeliveryExcel(
        fileList.map((file) => file.originFileObj)
      );
      setData(data);
    }

    // 生成新的 装箱单的 workbook 对象
    const workbook = await generateLSPackingList(basicValues, data);
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
    generate();
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
        {/* <Table columns={columns} dataSource={dataSource} /> */}
        <div className={styles["delivery-content"]}>
          <div className={styles["excel-preview"]}>
            <UniverSheet buffer={buffer} workBook={workbook} />
          </div>
          <div className={styles["deliver-right"]}>
            <div className={styles["first"]}>
              <span style={{ marginRight: "8px" }}>
                <div className={styles["step"]}>1</div>
                上传送货单
              </span>
              <Upload {...props}>
                <Button type="primary" icon={<UploadOutlined />}>
                  Upload
                </Button>
                {/* <InfoCircleOutlined /> */}
                {/* <Button icon={<ClearOutlined />} onClick={handleClearData}>
                清空上传数据
              </Button> */}
              </Upload>
            </div>
            <div className={styles['second']}>
              <div className={styles["step"]}>2</div>
              修改核对字段信息

              <BasicInfo
              ref={basicInfoRef}
              onFormValueChange={_.debounce(() => generate(fileList), 500)}
            />
            </div>

            <div className={styles["third"]}>
              <span style={{ marginRight: "8px" }}>
                <div className={styles["step"]}>3</div>
                导出装箱单
              </span>
              <Affix offsetBottom={0}>
                <Button type="primary" onClick={handleExportPackingList}>
                  导出装箱单
                </Button>
              </Affix>
            </div>
          </div>
        </div>
      </div>
    </Layout>
  );
}
