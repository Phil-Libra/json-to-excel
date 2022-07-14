import { useState } from "react";

import styles from './App.module.css';

const XLSX = window.XLSX;

function App() {
  // 要生成的网页表格内容
  const [html, setHTML] = useState('');
  // 转换出来的JSON数据
  const [excelData, setExcelData] = useState({});
  // 设定文件名
  const [fileName, setFileName] = useState('');

  // JSON转换为excel函数
  const JSONtoExcel = (file) => {
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

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, sheetData, "Sheet1");
        setExcelData(workbook);

      } catch (e) {
        console.log(e);
        alert('文件类型不正确');
        return;
      };
    };

    // 以二进制方式打开文件
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
    const fileNameArray = files[0].name.split('.');
    const fileName = fileNameArray.slice(0, fileNameArray.length - 1).join('.');

    setFileName(fileName);

    // 将表格转换为JSON数据
    JSONtoExcel(files[0]);
  };

  const generateExcel = () => {
    XLSX.writeFile(excelData, `${fileName}.xlsx`);
  }

  return (
    <>
      <fieldset>
        <legend>说明</legend>
        <p><a href="https://github.com/Phil-Libra/json-to-excel">源代码</a></p>
        <br />
        <p>生成的文件名格式：源文件名.xlsx</p>
        <br />
        <p>仅支持JSON文件上传。</p>
        <br />
        <p>由于限制，暂时只支持key value为简单类型值的JSON，否则无法获取到相应的key value，且转换出的数据可能有bug。</p>
        <br />
        <p>暂时获取第一个子项的key为表头，不排除以后提供选择表头的功能。</p>
      </fieldset>

      <fieldset id='file'>
        <legend>上传文件</legend>
        <input type="file" name="json-file" id="json-file" onChange={handleUpload} />
        <button onClick={generateExcel}>下载工作表</button>
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