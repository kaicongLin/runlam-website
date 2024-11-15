import React, { useEffect } from "react";
import LuckyExcel from "luckyexcel";
import styles from "./index.module.scss";

declare global {
  interface Window {
    luckysheet: any;
  }
}

const LuckySheet = ({ buffer }: { buffer: ArrayBuffer }) => {
  const init = () => {
    LuckyExcel.transformExcelToLucky(buffer, (exportJson, luckysheetfile) => {
      exportJson.sheets[0].zoomRatio = 1;
      console.log("exportJson", exportJson);
      console.log("window.luckysheet", window.luckysheet);
      if (window.luckysheet && window.luckysheet.create) {
        console.log("inner", window.luckysheet.create);

        window.luckysheet?.create({
          container: "excel", //luckysheet is the container id
          lang: "zh",
          // showtoolbar: false, //是否显示工具栏
          showinfobar: false, //是否显示顶部信息栏
          // showsheetbar: false, //是否显示底部sheet页按钮
          // allowCopy: false, //是否允许拷贝
          // allowEdit: false, //是否允许编辑
          // showstatisticBar: false,//是否显示底部计数栏
          // sheetFormulaBar: false, //是否显示公示栏
          // enableAddRow: false, //是否允许添加行
          enableAddBackTop: false, //是否允需回到顶部
          // devicePixelRatio: 10, //设置比例
          data: exportJson.sheets,
          // title: exportJson.info.name,
          // userInfo: exportJson.info.name.creator,
          hook: {
            workbookCreateAfter: () => {
              console.log("workbookCreateAfter------------");
            },
          },
        });
      }
    });
  };
  useEffect(() => {
    init();

    setTimeout(() => {
      const dom = document.getElementsByClassName('luckysheet-grid-window-2')[0];
      if (dom) {
        const tdDom = dom.getElementsByClassName('luckysheet-paneswrapper');
        for (let i = 0; i < tdDom.length; i++) {
          (tdDom[i] as HTMLElement).style.padding = '0px';
        }
      }

      const scrollbar = document.getElementsByClassName('luckysheet-scrollbar-ltr');
      for (let i = 0; i < scrollbar.length; i++) {
        (scrollbar[i] as HTMLElement).style.zIndex = '1';
      }
    }, 100)
    

  }, [buffer]);
  return (
    <div id="excel" className={styles["preiew-luckysheet"]}>
      {/* <Button onClick={init}>测试</Button> */}
    </div>
  );
};

export default LuckySheet;
