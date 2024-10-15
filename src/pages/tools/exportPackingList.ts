import ExcelJS from "exceljs";
import moment from "moment";
import { IDeliveryItems, BasicInfoFieldType } from "./intefaces";
import { deliveryExcelColumns, TABLE_START_ROW } from "./data";

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

  const borderStyle: ExcelJS.Borders = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
    diagonal: { style: "thin" },
  };
  // 表格从第10行开始
  worksheet.getRow(TABLE_START_ROW).values = columns.map(
    (item) => item.englishName
  );
  worksheet.getRow(TABLE_START_ROW).font = {
    name: "Tahoma",
    size: 11,
  };
  worksheet.getRow(TABLE_START_ROW + 1).values = columns.map(
    (item) => item.name
  );
  worksheet.getRow(TABLE_START_ROW + 1).font = {
    name: "Tahoma",
    size: 11,
  };

  data.forEach((item) => {
    const addRowsData = item.deliveryList.map((delivery) => {
      // 缸号去重
      const dyeingVats = [
        ...new Set(
          item.deliveryList.map((delivery) => delivery.dyeing_vat_number)
        ),
      ];
      return {
        case_number: basicValues.case_number ?? "",
        roll_qty: "1",
        unit: "ROLL",
        order_no: delivery.customer_order_number,
        compostion: "Nylon82% Spandex18%",
        description: delivery.product_number,
        color: delivery.color_name,
        width: "150CM",
        qty: delivery.quantity,
        length_unit: delivery.unit === "Y" ? "YD" : delivery.unit,
        net_weight: delivery.net_total_kg,
        gross_weight: delivery.net_total_kg + 0.5,
        meas: "26*26*153",
        remark: `${delivery.customer_order_number}, ${
          delivery.product_number
        }, ${delivery.color_number}, ${delivery.color_name}\n\n缸号${dyeingVats
          .map((vat) => vat)
          .join("、")}, 共${item.deliveryList.length}卷 报告OK`,
      };
    });
    const beforeNumber = worksheet.lastRow.number;
    const addRows = worksheet.addRows(addRowsData);
    addRows.map((row) => {
      row.font = {
        name: "Tahoma",
        size: 10,
      };
    });
    // 合并备注列
    worksheet.mergeCells(
      `N${beforeNumber + 1}:N${beforeNumber + addRowsData.length}`
    );

    //  --------汇总行 start --------
    const summary = {
      roll_qty: item.deliveryList.length,
      qty: item.deliveryList.reduce(
        (acc, cur) => acc + Number(cur.quantity),
        0
      ),
      net_weight: item.deliveryList.reduce(
        (acc, cur) => acc + Number(cur.net_total_kg),
        0
      ),
      gross_weight: item.deliveryList.reduce(
        (acc, cur) => acc + Number(cur.net_total_kg) + 0.5,
        0
      ),
      remark: `合计`,
    };

    const summaryRow = worksheet.addRow(summary);
    summaryRow.eachCell({ includeEmpty: true }, (cell) => {
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
  //  --------汇总行 end --------

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

  // 生成新的 Excel 文件并下载
  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
};
