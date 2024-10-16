import ExcelJS from "exceljs";
import moment from "moment";
import { IDeliveryItems, BasicInfoFieldType } from "./intefaces";
import {
  deliveryExcelColumns,
  TABLE_START_ROW,
  productNumberToCompostionMap,
} from "./data";
import { getColumnsLetter } from "./utils";

const borderStyle: ExcelJS.Borders = {
  top: { style: "thin" },
  left: { style: "thin" },
  bottom: { style: "thin" },
  right: { style: "thin" },
  diagonal: { style: "thin" },
};

/**
 * 构建excel模板信息
 * @param worksheet 
 * @param basicValues 
 */
const fillBasicInfo = (worksheet: ExcelJS.Worksheet, basicValues: BasicInfoFieldType) => {
  // 基础信息
  worksheet.getCell("A1").value = "Lintas Company Limited";
  worksheet.getCell("A2").value =
    "2A, Dragon Industrial Bldg., 93 King Lam Street, Cheung Sha Wan, Kowloon, Hong Kong.";
  worksheet.getCell("A3").value =
    "Tel : (852) 2311 8028 / Fax : (852) 2311 2262";
  worksheet.getCell("A4").value = "DELIVERY TO : Lintas Bangladesh Co. Ltd.";
  worksheet.getCell("A5").value =
    "SFB # 5, Adamjee Export Processing Zone;Siddhirganj, Narayanganj, Bangladesh ";
  worksheet.getCell("A6").value =
    "VAT + BIN no. 000150053-0305      ATTN: Mr. Sujit        Tel no. 8801774-950646";
  worksheet.getCell("A7").value = "TEL : (852) 2311 8028 ";
  worksheet.getCell("A8").value = "***packing list ***(UPDATE)";
  worksheet.mergeCells("A8:N8");

  worksheet.getCell("D9").value = "Invoice no.";
  worksheet.getCell("E9").value = basicValues.invoice_no;
  worksheet.getCell("F9").value = "packing list no.";
  worksheet.getCell("G9").value = basicValues.packing_list_no;
  worksheet.getCell("H9").value = "Date";
  worksheet.getCell("I9").value = new Date();
  worksheet.getCell("I9").numFmt = "d-mmm-yy";

  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      cell.font = {
        name: "Calibri",
        size: 11,
      };
    });
  });

    // 构建表头，表格从第10行开始
    worksheet.getRow(TABLE_START_ROW).values = deliveryExcelColumns.map(
      (item) => item.englishName
    );
    worksheet.getRow(TABLE_START_ROW).font = {
      name: "Tahoma",
      size: 11,
    };
    worksheet.getRow(TABLE_START_ROW + 1).values = deliveryExcelColumns.map(
      (item) => item.name
    );
    worksheet.getRow(TABLE_START_ROW + 1).font = {
      name: "Tahoma",
      size: 11,
    };
}

export const exportLSPackingList = async (
  basicValues: BasicInfoFieldType,
  data: IDeliveryItems[]
) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(moment().format("YYYYMMDD"));

  // 注意：这些列结构仅是构建工作簿的方便之处，除了列宽之外，它们不会完全保留
  const columns = deliveryExcelColumns;
  worksheet.columns = columns;

  // 填充基础信息
  fillBasicInfo(worksheet, basicValues);

  // 箱号，自动递增初始值
  let case_number_index = -1;

  data.forEach((item) => {
    const addRowsData = item.deliveryList.map((delivery) => {
      // 缸号去重
      const dyeingVats = [
        ...new Set(
          item.deliveryList.map((delivery) => delivery.dyeing_vat_number)
        ),
      ];
      // 分割箱号，取最后数值部分自动递增
      const parts = basicValues.case_number.split("-");
      const firstPart = parts[0] + "-" + parts[1] + "-";
      const secondPart = parts[2];
      case_number_index++;

      // 产品号映射成分
      const compostion = productNumberToCompostionMap[delivery.product_number];
      return {
        case_number: `${firstPart}${Number(secondPart) + case_number_index}`,
        roll_qty: Number(basicValues.roll_qty),
        unit: basicValues.unit,
        order_no: delivery.customer_order_number ?? "error",
        compostion: compostion ?? "error",
        description: delivery.product_number ?? "error",
        color: delivery.color_name ?? "error",
        width: basicValues.width,
        qty: delivery.quantity ?? "error",
        length_unit: delivery.unit === "Y" ? "YD" : delivery.unit ?? "error",
        net_weight: delivery.net_total_kg ?? "error",
        gross_weight: delivery.net_total_kg
          ? delivery.net_total_kg + 0.5
          : "error",
        meas: basicValues.meas,
        remark: `${delivery.customer_order_number ?? ""}, ${
          delivery.product_number ?? ""
        }, ${delivery.color_number ?? ""}, ${
          delivery.color_name ?? ""
        }\n\n缸号${dyeingVats.map((vat) => vat).join("、")}, 共${
          item.deliveryList.length ?? 0
        }卷 报告OK`,
      };
    });
    const beforeNumber = worksheet.lastRow.number;
    const addRows = worksheet.addRows(addRowsData);
    addRows.map((row) => {
      row.font = {
        name: "Tahoma",
        size: 10,
      };
      // 毛重列数（数字）
      const gross_weight_number =
        columns.findIndex((item) => item.key === "gross_weight") + 1;
      // 净重列数（字母）
      const net_weight_letter = getColumnsLetter("net_weight");
      // 设置毛重公式
      row.getCell(gross_weight_number).value = {
        formula: `${net_weight_letter}${row.number}+0.5`,
      };
    });
    // 合并备注列
    const remark_letter = getColumnsLetter("remark");
    worksheet.mergeCells(
      `${remark_letter}${beforeNumber + 1}:${remark_letter}${
        beforeNumber + addRowsData.length
      }`
    );

    //  --------送货单汇总行 --------
    const summary = {
      roll_qty: item.deliveryList.length,
      qty: item.deliveryList.reduce(
        (acc, cur) => acc + Number(cur.quantity ?? 0),
        0
      ),
      net_weight: item.deliveryList.reduce(
        (acc, cur) => acc + Number(cur.net_total_kg ?? 0),
        0
      ),
      gross_weight: item.deliveryList.reduce(
        (acc, cur) => acc + Number(cur.net_total_kg ?? 0) + 0.5,
        0
      ),
      remark: `合计`,
    };

    const summaryRow = worksheet.addRow(summary);
    summaryRow.eachCell({ includeEmpty: true }, (cell: any) => {
      if (cell._column._key === "roll_qty") {
        const roll_qty_letter = getColumnsLetter("roll_qty");
        cell.value = {
          formula: `SUM(${roll_qty_letter}${
            beforeNumber + 1
          }:${roll_qty_letter}${summaryRow.number - 1})`,
          result: summary.roll_qty,
        };
      }
      if (cell._column._key === "qty") {
        const qty_letter = getColumnsLetter("qty");
        cell.value = {
          formula: `SUM(${qty_letter}${beforeNumber + 1}:${qty_letter}${
            summaryRow.number - 1
          })`,
          result: summary.qty,
        };
      }
      if (cell._column._key === "net_weight") {
        const net_weight_letter = getColumnsLetter("net_weight");
        cell.value = {
          formula: `SUM(${net_weight_letter}${
            beforeNumber + 1
          }:${net_weight_letter}${summaryRow.number - 1})`,
          result: summary.net_weight,
        };
      }
      if (cell._column._key === "gross_weight") {
        const gross_weight_letter = getColumnsLetter("gross_weight");
        cell.value = {
          formula: `SUM(${gross_weight_letter}${
            beforeNumber + 1
          }:${gross_weight_letter}${summaryRow.number - 1})`,
          result: summary.gross_weight,
        };
      }
      cell.border = borderStyle;
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF00" },
      };
      cell.font = {
        name: "Tahoma",
        size: 10,
        bold: true,
        color: { argb: "FF0000" },
      };
    });
  });

  // 送货单整体汇总
  worksheet.getRow(worksheet.lastRow.number + 1).height = 80;

  const summary_start_number = worksheet.lastRow.number + 1;
  let subSummaryNumber = TABLE_START_ROW + 1;
  data.forEach((item, index) => {
    const row = worksheet.getRow(summary_start_number + index);
    const delivery = item.deliveryList[0];
    subSummaryNumber += item.deliveryList.length + 1;
    row.getCell(1).value = delivery.customer_order_number;
    row.getCell(2).value = delivery.product_number;
    row.getCell(3).value = delivery.color_name;
    row.getCell(4).value = "SUM:";
    row.getCell(5).value = "rolls";
    row.getCell(6).value = { formula: `B${subSummaryNumber}` };
    row.getCell(7).value = delivery.unit === "Y" ? "YD" : delivery.unit;
    row.getCell(8).value = { formula: `I${subSummaryNumber}` };
    row.getCell(9).value = "NET:";
    row.getCell(10).value = { formula: `K${subSummaryNumber}` };
    row.getCell(11).value = "GROSS:";
    row.getCell(12).value = { formula: `L${subSummaryNumber}` };
    row.getCell(13).value = { formula: `N${subSummaryNumber - item.deliveryList.length}` };

    row.font = {
      name: "Tahoma",
      size: 10,
    }
    worksheet.mergeCells(`M${row.number}:N${row.number}`);
  });
  
  // 总汇总
  const totalSummaryRow = worksheet.getRow(worksheet.lastRow.number + 1)
  totalSummaryRow.getCell(4).value = "TOTAL:";
  totalSummaryRow.getCell(5).value = "rolls";
  totalSummaryRow.getCell(6).value = { formula: `SUM(F${summary_start_number}:F${summary_start_number + data.length - 1})` };
  totalSummaryRow.getCell(7).value = data[0].deliveryList[0].unit === "Y" ? "YD" : data[0].deliveryList[0].unit;
  totalSummaryRow.getCell(8).value = { formula: `SUM(H${summary_start_number}:H${summary_start_number + data.length - 1})`}
  totalSummaryRow.getCell(9).value = "NET:";
  totalSummaryRow.getCell(10).value = { formula: `SUM(J${summary_start_number}:J${summary_start_number + data.length - 1})` };
  totalSummaryRow.getCell(11).value = "GROSS:";
  totalSummaryRow.getCell(12).value = { formula: `SUM(L${summary_start_number}:L${summary_start_number + data.length - 1})` };
  totalSummaryRow.eachCell((cell) => {
    cell.border = borderStyle;
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF00" },
    };
    cell.font = {
      name: "Tahoma",
      size: 10,
      bold: true,
      color: { argb: "FF0000" },
    };
  })

  // 设置excel字体样式
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      // 设置水平和垂直居中
      if (rowNumber > 7) {
        cell.alignment = {
          horizontal: "center",
          vertical: "middle",
          wrapText: true,
        };
      }

      // 设置边框
      if (rowNumber > TABLE_START_ROW - 1 && colNumber < columns.length + 1) {
        cell.border = borderStyle;
      }
    });
    // 设置行高
    if (rowNumber > TABLE_START_ROW + 1) {
      row.height = 80;
    }
  });

  // 生成新的 Excel 文件
  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
};
