import React, { useEffect, useRef } from "react";
import Layout from "@theme/Layout";
import ExcelJS from "exceljs";
import { UploadOutlined } from "@ant-design/icons";
import type { UploadFile, UploadProps } from "antd";
import { Button, message, Upload, Affix, Table } from "antd";
import { IDeliveryItems } from "./_intefaces";
import { deliveryExcelCloumnsMap } from "./_data";
import styles from "./index.module.scss";
import BasicInfo, { BasicInfoRefProps } from "./_BasicInfo";
import { readDeliveryExcel, excelBufferToFile } from "./_utils";
import _ from "lodash";
import { generateLSPackingList, generateLSTag } from "./_generateLS";
import LuckySheet from "./_Luckysheet";
import JSZip from "jszip";
import moment from "moment";
import { saveAs } from "file-saver";
import BrowserOnly from "@docusaurus/BrowserOnly";

interface IColumns {
  title: string;
  dataIndex: string;
  key: string;
}

export default function Hello() {
  const [data, setData] = React.useState<IDeliveryItems[]>([]);
  const [columns, setColumns] = React.useState<IColumns[]>([]);
  const [fileList, setFileList] = React.useState<UploadFile[]>([]);
  const [buffer, setBuffer] = React.useState<ArrayBuffer>();
  const [workbook, setWorkbook] = React.useState<ExcelJS.Workbook | null>(null);
  const removeFlag = useRef<boolean>(false);

  const fileCount = useRef<number>(0);
  const basicInfoRef = useRef<BasicInfoRefProps>(null);

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
    const basicValues = basicInfoRef.current?.getFieldsValue() || {};
    let data: IDeliveryItems[] = [];
    // 如果有文件列表，则读取文件内容
    if (fileList) {
      data = await readDeliveryExcel(
        fileList.map((file) => file.originFileObj)
      );
    }

    // 生成新的 装箱单的 workbook 对象
    const workbook = await generateLSPackingList(basicValues, data);
    const buffer = await workbook.xlsx.writeBuffer();

    setData(data);
    setBuffer(buffer);
    setWorkbook(workbook);
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
    const columns: IColumns[] = [];
    for (const [key, val] of Object.entries(deliveryExcelCloumnsMap)) {
      columns.push({
        title: key,
        dataIndex: val,
        key: val,
      });
    }
    columns.push({
      title: "箱号",
      dataIndex: "case_number",
      key: "case_number",
    });
    setColumns(columns);
    // handleExportPackingList();
    generate();
  }, []);

  const dataSource = _.flatMap(data, (item) => item.deliveryList);

  return (
    <Layout title="Hello" description="Hello React Page">
      <BrowserOnly fallback={<div>Loading...</div>}>
        {() => {
          const Docxtemplater = require("docxtemplater");
          const PizZip = require("pizzip");
          const PizZipUtils = require("pizzip/utils/index.js");
          const expressionParser = require("docxtemplater/expressions");
          const generateDocument = (data: any, jszip: JSZip) => {
            return new Promise((resolve, reject) => {
              PizZipUtils.getBinaryContent(
                "/runlam-website/template/LStemplate.docx",
                (error: any, content: any) => {
                  if (error) {
                    throw error;
                  }
                  const zip = new PizZip(content);
                  const doc = new Docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                    parser: expressionParser,
                  });
                  doc.render(data);
                  const out = doc.getZip().generate({
                    type: "blob",
                    mimeType:
                      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  }); //Output the document using Data-URI

                  jszip.file(data.fileName + ".docx", out);
                  return resolve(out);
                  //   saveAs(out, data.fileName + ".docx");
                }
              );
            });
          };

          const handleExport = async () => {
            // 初始化一个zip打包对象
            const zip = new JSZip();
            const basicValues = basicInfoRef.current?.getFieldsValue() || {};

            // 生成灵达装箱单excel文件
            const totalRollNumber = data.reduce(
              (acc, cur) => acc + cur.deliveryList.length,
              0
            );
            const blob = excelBufferToFile(buffer);
            zip.file(
              `${basicValues.invoice_no}装箱单${totalRollNumber}卷.xlsx`,
              blob
            );

            // 生成灵达标签word文件
            const promises = data.map(async (item) => {
              const wordData = generateLSTag(item, basicValues);
              return await generateDocument(wordData, zip);
            });

            // 下载压缩包
            await Promise.all(promises);
            zip.generateAsync({ type: "blob" }).then((blob) => {
              saveAs(blob, `灵达出货${moment().format("YYYY-MM-DD")}.zip`);
            });
          };

          return (
            <div className={styles["delivery-wrap"]}>
              {/* <Table columns={columns} dataSource={dataSource} /> */}
              <div className={styles["delivery-content"]}>
                <div className={styles["excel-preview"]}>
                  <LuckySheet buffer={buffer} />
                </div>
                <div className={styles["deliver-right"]}>
                  <div className={styles["first"]}>
                    <span style={{ marginRight: "8px" }}>
                      <div className={styles["step"]}>1</div>
                      上传系统送货单
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
                  <div className={styles["second"]}>
                    <div className={styles["step"]}>2</div>
                    修改核对字段信息
                    <BasicInfo
                      ref={basicInfoRef}
                      onFormValueChange={_.debounce(
                        () => generate(fileList),
                        500
                      )}
                    />
                  </div>

                  <div className={styles["third"]}>
                    <span style={{ marginRight: "8px" }}>
                      <div className={styles["step"]}>3</div>
                      导出装箱单和标签
                    </span>
                    <Affix offsetBottom={0}>
                      <Button type="primary" onClick={handleExport}>
                        导出
                      </Button>
                    </Affix>
                  </div>
                </div>
              </div>
            </div>
          );
        }}
      </BrowserOnly>
    </Layout>
  );
}
