import ExcelJS from "exceljs";
import { IDeliveryItem, IDeliveryItems } from "./intefaces";
import {
  SHEET_NAME,
  deliveryExcelCloumnsMap,
  deliveryExcelColumns
} from "./data";

/**
 * 读取送货Excel文件
 * @param files
 * @returns
 */
export const readDeliveryExcel = async (
  files: File[]
): Promise<IDeliveryItem[]> => {
  const allData: IDeliveryItems[] = [];
  const promises: Promise<void>[] = [];
  for (const file of files) {
    // 使用 FileReader 读取文件内容
    const reader = new FileReader();
    // 将文件内容读取为 ArrayBuffer
    reader.readAsArrayBuffer(file);
    const promise = new Promise<void>((resolve, reject) => {
      reader.onload = async (e) => {
        // 获取 arrayBuffer
        const buffer = e.target?.result as ArrayBuffer;
        // 创建一个新的 ExcelJS.Workbook 对象
        const workbook = new ExcelJS.Workbook();
        // 加载 Excel 文件内容
        await workbook.xlsx.load(buffer);
        // 按照工作表名称获取工作表
        const worksheet = workbook.getWorksheet(SHEET_NAME);
        const data: IDeliveryItem[] = [];
        // 遍历工作表的每一行
        worksheet.eachRow((row, rowNumber) => {
          const rowData: IDeliveryItem = {};
          // 遍历每一行的每一个单元格(包括空单元格)
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            rowData[
              deliveryExcelCloumnsMap[
                worksheet.getCell(1, colNumber).value as string
              ]
            ] = cell.value;
          });
          // 跳过第一行的标题行，从第二行开始读取数据
          if (rowNumber !== 1) {
            data.push(rowData);
          }
        });
        const chineseCharacters = file.name.match(/[\u4e00-\u9fa5]+/);
        allData.push({
          name: chineseCharacters ? chineseCharacters[0] : file.name,
          deliveryList: data,
        });
        resolve();
      };
      reader.onerror = reject;
    });
    promises.push(promise);
  }
  await Promise.all(promises);
  return allData;
};

/**
 * 下载文件
 * @param data
 * @param filename
 */
export const downloadFile = (data: any, filename?: string) => {
  const blob = new Blob([data], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8",
  }) as any;
  const url = URL.createObjectURL(blob);
  const aLink = document.createElement("a");
  aLink.setAttribute(
    "download",
    filename ? filename : `${new Date().getTime()}.xlsx`
  );
  aLink.setAttribute("href", url);
  document.body.appendChild(aLink);
  aLink.click();
  document.body.removeChild(aLink);
  URL.revokeObjectURL(blob);
};

/**
 * buffer流转文件流
 * @param buffer 
 * @param fileName 
 * @param mimeType 
 * @returns 
 */
export const bufferToFile = (buffer: ArrayBuffer, fileName: string, mimeType?: string) => {
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return new File([blob], fileName, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
};

/**
 * 字母转数字
 * @param letters
 * @returns
 */
export function letterToNumber(letters: string) {
  let result = 0;
  const base = 26;
  for (let i = 0; i < letters.length; i++) {
    const charCode = letters.charCodeAt(i) - 64;
    result += charCode * Math.pow(base, letters.length - i - 1);
  }
  return result;
}

/**
 * 数字转字母
 * @param number
 * @returns
 */
export function numberToLetter(number: number) {
  let result = '';
  let n = number;
  const base = 26;
  while (n > 0) {
    n--;
    let remainder = n % base;
    result = String.fromCharCode(remainder + 65) + result;
    n = Math.floor(n / base);
  }
  return result || 'A';
}

/**
 * 获取列字母(默认表格从A开始)
 * @param key
 * @returns
 */
export function getColumnsLetter(key: string) {
  const number = deliveryExcelColumns.findIndex((item) => item.key === key) + 1;
  return numberToLetter(number);
}