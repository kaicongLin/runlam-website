import "@univerjs/design/lib/index.css";
import "@univerjs/ui/lib/index.css";
import "@univerjs/docs-ui/lib/index.css";
import "@univerjs/sheets-ui/lib/index.css";
import "@univerjs/sheets-formula-ui/lib/index.css";

import { LocaleType, Univer, UniverInstanceType } from "@univerjs/core";
import { defaultTheme } from "@univerjs/design";

import { UniverFormulaEnginePlugin } from "@univerjs/engine-formula";
import { UniverRenderEnginePlugin } from "@univerjs/engine-render";

import { UniverUIPlugin } from "@univerjs/ui";

import { UniverDocsPlugin } from "@univerjs/docs";
import { UniverDocsUIPlugin } from "@univerjs/docs-ui";

import { UniverSheetsPlugin } from "@univerjs/sheets";
import { UniverSheetsFormulaPlugin } from "@univerjs/sheets-formula";
import { UniverSheetsFormulaUIPlugin } from "@univerjs/sheets-formula-ui";
import { UniverSheetsUIPlugin } from "@univerjs/sheets-ui";
import { zhCN, enUS } from "univer:locales";
import { FUniver } from "@univerjs/facade";
 
import React, { useEffect, useRef } from "react";
import { convertExceljsToUniver } from "./_utils";
import styles from "./index.module.css";

import ExcelJS from "exceljs";

const UniverSheet = ({ workBook }: { workBook: ExcelJS.Workbook }) => {
  const containerRef = useRef(null);

  const init = (workBook: ExcelJS.Workbook) => {
    if (!containerRef.current) {
      throw Error("container not initialized");
    }

    const univer = new Univer({
      theme: defaultTheme,
      locale: LocaleType.ZH_CN,
      locales: {
        [LocaleType.ZH_CN]: zhCN,
        [LocaleType.EN_US]: enUS,
      },
    });

    univer.registerPlugin(UniverRenderEnginePlugin);
    univer.registerPlugin(UniverFormulaEnginePlugin);

    univer.registerPlugin(UniverUIPlugin, {
      container: containerRef.current,
    });

    univer.registerPlugin(UniverDocsPlugin, {
      hasScroll: false,
    });
    univer.registerPlugin(UniverDocsUIPlugin);

    univer.registerPlugin(UniverSheetsPlugin);
    univer.registerPlugin(UniverSheetsUIPlugin);
    univer.registerPlugin(UniverSheetsFormulaPlugin);
    univer.registerPlugin(UniverSheetsFormulaUIPlugin);
    
    if (workBook) {
      // univerAPI.value.importXLSXToSnapshot(file).then((res) => {
      // })
      const excelData = convertExceljsToUniver(workBook);

      univer.createUnit(UniverInstanceType.UNIVER_SHEET, excelData);
      // 设置其他属性
      const univerAPI = FUniver.newAPI(univer);
      const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet();
      const sheetName = sheet.getSheetName();

      const ExceljsWorkSheet = workBook.getWorksheet(sheetName)
      ExceljsWorkSheet.columns.forEach((c) => {
        sheet?.setColumnWidth(c.number - 1, c.width * 6);
      });
      (ExceljsWorkSheet as any)._rows.forEach((r) => {
        sheet?.setRowHeight(r.number - 1, r.height);
      })
      // sheet?.setRowHeights(11, 1000, 80);
    }
  };

  useEffect(() => {
    init(workBook);
  }, [workBook]);

  return <div ref={containerRef} className={styles["univer-container"]} />;
};

export default UniverSheet;
