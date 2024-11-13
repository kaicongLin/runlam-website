import ExcelJS from "exceljs";
import { IDeliveryItem, IDeliveryItems } from "./_intefaces";
import {
  SHEET_NAME,
  deliveryExcelCloumnsMap,
  deliveryExcelColumns,
} from "./_data";
import {
  Workbook as ExceljsWorkbook,
  Worksheet as ExceljsWorksheet,
} from "exceljs";
import { BooleanNumber, LocaleType, SheetTypes } from "@univerjs/core";

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
export const bufferToFile = (
  buffer: ArrayBuffer,
  fileName: string,
  mimeType?: string
) => {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  return new File([blob], fileName, {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
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
  let result = "";
  let n = number;
  const base = 26;
  while (n > 0) {
    n--;
    let remainder = n % base;
    result = String.fromCharCode(remainder + 65) + result;
    n = Math.floor(n / base);
  }
  return result || "A";
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

function getUniverRgbHex(argbString?: string) {
  if (!argbString) {
    return "";
  }

  if (
    argbString.startsWith("#") &&
    (argbString.length === 7 || argbString.length === 9)
  ) {
    return `#${argbString.slice(
      argbString.length === 7 ? 1 : 3,
      argbString.length === 7 ? 7 : 9
    )}`;
  }

  if (
    !argbString.startsWith("#") &&
    (argbString.length === 6 || argbString.length === 8)
  ) {
    return `#${argbString.slice(
      argbString.length === 6 ? 0 : 2,
      argbString.length === 6 ? 6 : 8
    )}`;
  }

  return "";
}

/**
 * exceljs转univer
 * @param exceljsWorkbook
 * @returns
 */
export const convertExceljsToUniver = (
  exceljsWorkbook: ExceljsWorkbook
): any => {
  const sheets = {};
  const sheetOrder = [];
  // 处理工作表顺序
  exceljsWorkbook.eachSheet((worksheet, sheetId) => {
    const cellData = {};
    const mergeData = [];

    worksheet.eachRow((row, rowNumber) => {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const cl_rgb = getUniverRgbHex(cell.font?.color?.argb);
        const bg_rgb = getUniverRgbHex((cell.fill as any)?.fgColor?.argb);

        if (cell !== null && cell !== undefined) {
          if (!cellData[rowNumber - 1]) {
            cellData[rowNumber - 1] = [];
          }

          cellData[rowNumber - 1][colNumber - 1] = {
            v: cell.text, // 文本
            f: (cell.value as any)?.formula ? `=${(cell.value as any)?.formula}` : ``,
            s: {
              fs: cell.font?.size, // 字体大小
              ff: cell.font?.name, // 字体名称
              bl: cell.font?.bold ? 1 : 0, // bl 是一个布尔数字，0 表示不加粗，1 表示加粗。
              tb: cell.alignment?.wrapText ? 3 : 1, // 截断溢出: 3 换行，1 溢出， 2截断
              ht: cell.alignment?.horizontal === "center" ? 2 : 1, //  水平对齐， 1 表示左对齐，2 表示居中，3 表示右对齐
              vt: cell.alignment?.vertical === "middle" ? 2 : 1, // 垂直对齐，1 表示顶部对齐，2 表示居中对齐，3 表示底部对齐
              cl: {
                rgb: cl_rgb, // 字体颜色
              },
              bg: {
                rgb: bg_rgb, // 背景颜色
              },
              bd: {
                // 边框
                t: cell.border?.top
                  ? {
                      s: 1, // 边框样式
                      cl: {
                        rgb: "#000000",
                      },
                    }
                  : null,
                b: cell.border?.bottom
                  ? {
                      s: 1,
                      cl: {
                        rgb: "#000000",
                      },
                    }
                  : null,
                l: cell.border?.left
                  ? {
                      s: 1,
                      cl: {
                        rgb: "#000000",
                      },
                    }
                  : null,
                r: cell.border?.right
                  ? {
                      s: 1,
                      cl: {
                        rgb: "#000000",
                      },
                    }
                  : null,
              },
            },
          };
        }
      });
    });

    if (worksheet.hasMerges) {
      Object.values((worksheet as any)._merges).forEach((item: any) => {
        mergeData.push({
          startRow: item.model.top - 1,
          startColumn: item.model.left - 1,
          endRow: item.model.bottom - 1,
          endColumn: item.model.right - 1,
        });
      });
    }

    // if (worksheet.columns.length > 0) {
    //   columnHeader = worksheet.columns.map((c: any) => ({
    //     height: c.width
    //         }));
    // }

    sheets[sheetId.toString()] = {
      id: sheetId.toString(),
      name: worksheet.name,
      tabColor: worksheet.properties.tabColor,
      hidden: BooleanNumber.FALSE,
      rowCount: worksheet.rowCount + 30,
      columnCount: worksheet.columnCount + 10,
      zoomRatio: 1,
      cellData: cellData,
      mergeData: mergeData,
      showGridlines: 1,
    };
    sheetOrder.push(sheetId.toString());
  });

  return {
    id: "univer_id",
    locale: LocaleType.ZH_CN,
    name: "univer_sheet",
    appVersion: "3.0.0-alpha",
    sheetOrder: sheetOrder,
    sheets: sheets,
  };
};
