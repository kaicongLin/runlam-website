import React, { useEffect } from 'react';
import Layout from '@theme/Layout';
import { UploadOutlined } from '@ant-design/icons';
import type { UploadProps } from 'antd';
import { Button, message, Table, Upload } from 'antd';
import { deliveryCloumnsMap } from '../../constants/delivery';
import ExcelJS from 'exceljs';
import styles from './index.module.css';


export default function Hello() {
  const [data, setData] = React.useState<any[]>([]);
  const [columns, setColumns] = React.useState<any[]>([]);

    const props: UploadProps = {
        name: 'file',
        accept: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel',
        customRequest: (options: any) => {
          const { onSuccess, file } = options;

          // 使用FileReader读取文件内容
          const reader = new FileReader();
          // 将文件内容读取为ArrayBuffer
          reader.readAsArrayBuffer(file);
          reader.onload = async (e) => {
            // 获取arrayBuffer
            const buffer = e.target?.result as ArrayBuffer;

            const workbook = new ExcelJS.Workbook();
            // 加载Excel文件内容
            await workbook.xlsx.load(buffer);
            console.log('workbook', workbook);
            // 按照工作表名称获取工作表
            const worksheet = workbook.getWorksheet('送货单明细');
            const data = [];
            // 遍历工作表的每一行
            worksheet.eachRow((row, rowNumber) => {
              const rowData = {};
              // 遍历每一行的每一个单元格(包括空单元格)
              row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                // 读取单元格的值
                rowData[deliveryCloumnsMap[worksheet.getCell(1, colNumber).value as string]] = cell.value;
              });
              // 跳过第一行的标题行，从第二行开始读取数据
              if (rowNumber!== 1) {
                data.push(rowData);
              }
            });
            console.log('data', data);
            setData(data);
          }

          onSuccess({ status: 'success', response: file });

          return {
            abort() {}
          }
        },
      };


      useEffect(() => {
        // 遍历对象的键值对
        const columns = [];
        for (const [key, val] of Object.entries(deliveryCloumnsMap)) {
          columns.push({
            title: key,
            dataIndex: val,
            key: val,
          })
        }
        setColumns(columns);
      },[]
      )
      
  return (
    <Layout title="Hello" description="Hello React Page">
      <div className={styles['delivery-wrap']}>
          <div className={styles['upload-button']}>
            <Upload {...props}>
              <Button icon={<UploadOutlined />}>请上传送货单</Button>
            </Upload>
          </div>

        <Table columns={columns} dataSource={data} />
      </div>
    </Layout>
  );
}