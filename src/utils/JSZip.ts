import JSZip from "jszip";

  // 遍历Zip文件
export const iterateZipFile = async (data: ArrayBuffer, iterationFn: Function) => {
    if (typeof iterationFn !== "function") {
      throw new Error("iterationFn 不是函数类型");
    }
    let zip;
    try {
      zip = await JSZip.loadAsync(data);
      zip.forEach(iterationFn);
      return zip;
    } catch (error) {
      throw new error();
    }
}
