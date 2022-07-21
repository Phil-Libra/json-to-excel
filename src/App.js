import { useState } from "react";

import styles from './App.module.css';

const XLSX = window.XLSX;

const App = () => {
  // 要生成的网页表格内容
  const [html, setHTML] = useState();
  // 转换出来的JSON数据
  const [excelData, setExcelData] = useState();
  // 设定文件名
  const [fileName, setFileName] = useState();
  // 设定下载按钮显示状态
  const [dlStatus, setDLStatus] = useState(false);

  // JSON转换为表格函数
  const JSONtoTable = (file) => {
    const fileReader = new FileReader();

    fileReader.onload = (e) => {
      try {
        const { result } = e.target;

        const data = JSON.parse(result)
        // 获取JSON key
        const headerArray = Object.keys(data[0]);
        const sheetData = XLSX.utils.json_to_sheet(data, { header: headerArray });

        // 生成表格预览
        setHTML(XLSX.utils.sheet_to_html(sheetData));

        // 生成表格数据
        setExcelData(sheetData);

      } catch (e) {
        console.log(e);
        return;
      };
    };

    // 以文本格式打开文件
    fileReader.readAsText(file, 'UTF-8');
  };

  const handleUpload = async (e) => {
    // 获取文件并赋值给state
    const files = e.target.files;

    // 判定是否有文件上传及文件格式是否正确
    if (files.length === 0) {
      return;
    } else if (files.length !== 0 && files[0].type !== 'application/json') {
      alert('文件格式错误');
      return;
    }

    // 设置上传的文件名为默认下载文件名
    let fileNameArray = files[0].name.split('.');

    // 如果上传的文件只有扩展名，则加入默认的文件名'workbook'
    if (!fileNameArray[0]) {
      fileNameArray.splice(0, 1, 'workbook');
    }

    const fileName = fileNameArray.slice(0, fileNameArray.length - 1).join('.');

    // 设置文件名state
    setFileName(fileName);

    // 将JSON数据转换为表格
    JSONtoTable(files[0]);

    // 显示下载按钮
    setDLStatus(true);
  };

  const handleDownload = () => {
    // 防止未生成数据时直接被点击导致报错
    if (!excelData) {
      return;
    }

    // 生成工作表数据
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, excelData, "Sheet1");

    // 生成文件并下载
    XLSX.writeFile(workbook, `${fileName}.xlsx`);
  }

  return (
    <>
      <fieldset>
        <legend>说明</legend>
        <ul>
          <li><a href="https://github.com/Phil-Libra/json-to-excel">源代码</a></li>
          <li>生成的文件名格式：源文件名.xlsx</li>
          <li>仅支持JSON文件上传。</li>
          <li>由于限制，暂时只支持key value为简单数据类型的JSON，否则无法获取到相应的key value，且转换出的数据可能有bug。</li>
          <li>暂时获取第一个子项的key为表头，不排除以后提供选择表头的功能。</li>
          <li><strong>因{'<input />'}标签本身限制，连续上传同一个文件请刷新页面后再上传。</strong></li>
        </ul>
      </fieldset>

      <fieldset id='file'>
        <legend>上传文件</legend>
        <label htmlFor="upload">
          <input type="file" name="upload" id="upload" onChange={handleUpload} />
        </label>
        <button style={{ display: dlStatus ? '' : 'none' }} onClick={handleDownload}>下载工作表</button>
      </fieldset>

      <fieldset>
        <legend>表格数据预览</legend>
        <div
          className={styles.tablePreview}
          dangerouslySetInnerHTML={{ __html: html }}
        />
      </fieldset>
    </>
  )
};

export default App;