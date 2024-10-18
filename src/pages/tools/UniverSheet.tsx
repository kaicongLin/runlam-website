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

import { FUniver } from "@univerjs-pro/facade";

import React, { useEffect, useRef } from "react";
import { DEFAULT_WORKBOOK_DATA } from "./default-workbook-data";
import { bufferToFile } from "./utils";
import styles from "./index.module.css";

import ExcelJS from "exceljs";

const UniverSheet = ({ buffer }: { buffer: ArrayBuffer }) => {
  const containerRef = useRef(null);

  const init = (buffer: ArrayBuffer) => {
    if (!containerRef.current) {
      throw Error("container not initialized");
    }

    console.log("init univer", zhCN);
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

    const file = bufferToFile(buffer, "test.xlsx");
    console.log(file);
    const univerAPI = FUniver.newAPI(univer);
    if (file) {
      // univerAPI.value.importXLSXToSnapshot(file).then((res) => {
        // console.log(res);
      // })
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet1");
      worksheet.getCell("A1").value = "Hello World";
      // const fdata = workbook.xlsx.writeFile("new_example.xlsx");
      console.log('worksheet', worksheet);
    
      console.log(workbook, 'workbook');
      
    }

    univer.createUnit(UniverInstanceType.UNIVER_SHEET, DEFAULT_WORKBOOK_DATA);
  };

  useEffect(() => {
    init(buffer);
  }, [buffer]);

  return <div ref={containerRef} className={styles["univer-container"]} />;
};

export default UniverSheet;
