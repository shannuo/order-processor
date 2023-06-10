import logo from './logo.jpg';
import React, { useState } from 'react';
import './App.css';
import * as XLSX from 'xlsx';
import _ from 'lodash';
import moment from 'moment';

function App() {
  const [chefTime, setChefTime] = useState('');
  const [time, setTime] = useState('');
  const handleChange = (event) => {
    setChefTime(event.target.value);
  };
  const handleTimeChange = (event) => {
    setTime(event.target.value);
  };
  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <p>
          这是一个订单分解器，根据微信小程序下单助手官方生成的订单，分解成厨师、配送员、分拣员方便查看的三个订单数据。
        </p>
        <label htmlFor="input">输入时间段:</label>
        <input
          id="input"
          type="text"
          value={time}
          onChange={handleTimeChange}
        />
        <p>上传订单明细</p>
        <input type='file' accept='.xlsx, .xls' onChange={e => onImportExcel(e, time)} />
        <label htmlFor="input">输入时间段:</label>
        <input
          id="input"
          type="text"
          value={chefTime}
          onChange={handleChange}
        />
        <p>上传商品表格</p>
        <input type='file' accept='.xlsx, .xls' onChange={e => onImportProductsExcel(e, chefTime)} />
      </header>
    </div>
  );
}

const ORDER_STATUS_NAMES = {
  0: '无效订单',
  1: '有效订单',
};

const ORDER_EXTRA_INFO = {
  ORDER_STATUS: `订单是否有效`,
  REASON: `订单无效原因`,
  ADDRESSPART: `送餐地区`,
  ADDRESS: `送餐详细地址`,
  DATE: `送餐日期`,
  TIME:  `送餐时间`,
}

const SHEETHEADER = [
  { name: '订单编号', cellWidth: 4, cellMergeNumber: 1, isMerge: 1 },
  { name: '顾客备注', cellWidth: 20, cellMergeNumber: 1, isMerge: 1 },
  { name: '商品', cellWidth: 10, cellMergeNumber: 1, isMerge: 0 },
  { name: '商品单价（元）', cellWidth: 6, cellMergeNumber: 1, isMerge: 0 }, 
  { name: '购买数量', cellWidth: 4, cellMergeNumber: 1, isMerge: 0 },
  { name: ORDER_EXTRA_INFO.ADDRESSPART, cellWidth: 10, cellMergeNumber: 1, isMerge: 1 },
  { name: ORDER_EXTRA_INFO.ADDRESS, cellWidth: 10, cellMergeNumber: 1, isMerge: 1 },
  { name: '微信昵称/备注名', cellWidth: 8, cellMergeNumber: 1, isMerge: 1 },
  { name: '顾客电话', cellWidth: 6, cellMergeNumber: 1, isMerge: 1 },
  { name: '付款状态', cellWidth: 4, cellMergeNumber: 1, isMerge: 1 },
  { name: ORDER_EXTRA_INFO.ORDER_STATUS, cellWidth: 6, cellMergeNumber: 1, isMerge: 1 },
  { name: ORDER_EXTRA_INFO.REASON, cellWidth: 10, cellMergeNumber: 1, isMerge: 1 },
  { name: ORDER_EXTRA_INFO.DATE, cellWidth: 5, cellMergeNumber: 1, isMerge: 1 },
  { name: ORDER_EXTRA_INFO.TIME, cellWidth: 10, cellMergeNumber: 1, isMerge: 1 },
  { name: '订购时间', cellWidth: 10, cellMergeNumber: 1, isMerge: 1 },
  { name: '订单类型', cellWidth: 5, cellMergeNumber: 1, isMerge: 1 },
  { name: '顾客姓名', cellWidth: 4, cellMergeNumber: 1, isMerge: 1 },
  { name: '顾客地址', cellWidth: 4, cellMergeNumber: 1, isMerge: 1 },
  { name: '自提地址', cellWidth: 15, cellMergeNumber: 1, isMerge: 1 },
  { name: '自提时间', cellWidth: 10, cellMergeNumber: 1, isMerge: 1 },
  { name: '订单金额（元）', cellWidth: 6, cellMergeNumber: 1, isMerge: 1 },
];

const SHEETHEADER_FOR_CHEF = [
  { name: '商品销售统计', cellWidth: 15, cellMergeNumber: 1, isMerge: 1 },
  { name: '数量', cellWidth: 4, cellMergeNumber: 1, isMerge: 1 },
];

const MERGES = [];

const SHEETINFOS = {
  '分餐员表格': { 
    useless: ['顾客地址', '自提地址', '自提时间'],
  },
  '配送员表格': { 
    useless: ['顾客地址', '自提地址', '自提时间', '订购类型', '订购时间', '付款状态', ORDER_EXTRA_INFO.REASON, '订单类型'],
  },
};

// 处理订单数据
const dealData = (data, time) => {
  let orderIndex = 0; // 当前具有订单编号的对象的下标
  let currentMergeStart = 1; // 当前订单开始行
  let count = 0; // 合并的行数（商品条数）
  for (let i = 0; i < data.length; i++) {
    if (data[i][`订单编号`]) {
      if (count) {
        MERGES.push({ s: { r: currentMergeStart }, e: { r: currentMergeStart + count - 1}});
      }
      currentMergeStart = i+1;
      count = 0;
      orderIndex = i;
      data[i][ORDER_EXTRA_INFO.ORDER_STATUS] = ORDER_STATUS_NAMES[1];
      data[i][ORDER_EXTRA_INFO.REASON] = '';
      data[i][ORDER_EXTRA_INFO.ADDRESSPART] = '';
      data[i][ORDER_EXTRA_INFO.ADDRESS] = '';
      data[i][ORDER_EXTRA_INFO.DATE] = '';
      data[i][ORDER_EXTRA_INFO.TIME] = '';
      if (data[i][`订单金额（元）`] < 18.5) {
        data[i][ORDER_EXTRA_INFO.ORDER_STATUS] = ORDER_STATUS_NAMES[0];
        data[i][ORDER_EXTRA_INFO.REASON] = '订单金额不满18.5元';
      }
    }
    if (data[i][`商品`]) {
      count ++;
    }
    if (data[i][`商品`] && data[i][`商品`].indexOf('送餐') !== -1) { // 处理送餐时间和送餐地址
      if (data[orderIndex][ORDER_EXTRA_INFO.ADDRESSPART]) {
        data[orderIndex][ORDER_EXTRA_INFO.ORDER_STATUS] = ORDER_STATUS_NAMES[0];
        data[orderIndex][ORDER_EXTRA_INFO.REASON] = '选择多个送餐地址';
      } else {
        let str = data[i][`商品`];
        const splitParts = _.split(str, '('); // 拆分成两部分，以'('作为分隔符
        const addressPart = _.trim(splitParts[0]); // 获取地址部分并去除多余的空格
        const infoPart = _.trimEnd(splitParts[1], ')'); // 获取信息部分并去除多余的空格和括号
        const splitInfo = _.split(infoPart, ',')
        data[orderIndex][ORDER_EXTRA_INFO.ADDRESSPART] = addressPart;
        const pattern = /\b\d{1,2}:\d{2}[ap]m-\d{1,2}:\d{2}[ap]m\b/; // 判断是否是配送时间格式
        let timeInfo = '';
        if (pattern.test(splitInfo[0])) {
          data[orderIndex][ORDER_EXTRA_INFO.ADDRESS] = splitInfo[1];
          timeInfo = splitInfo[0];
        } else {
          data[orderIndex][ORDER_EXTRA_INFO.ADDRESS] = splitInfo[0];
          timeInfo = splitInfo[1];
        }
        data[orderIndex][ORDER_EXTRA_INFO.TIME] = _.replace(timeInfo, '次日', '');
        const date = new Date(data[orderIndex][`订购时间`]); // 假设获取的日期为 2023/6/2 22:18:44
        // 使用 Moment.js 解析日期对象
        let momentDate = moment(date);
        if (timeInfo.indexOf('次日') !== -1) {
          momentDate = momentDate.add(1, 'day')
        }
        // 格式化为 "几月几日" 的字符串
        const formattedDate = momentDate.format('M月D日');
        data[orderIndex][ORDER_EXTRA_INFO.DATE] = formattedDate;
      }
    }
  }
  // 处理最后一条订单
  MERGES.push({ s: { r: currentMergeStart }, e: { r: currentMergeStart + count - 1}});
  downloadFileToExcel(data, `分餐员表格`, '', time);
  downloadFileToExcel(data, `配送员表格`, '', time );
}

const dealProductsData = (data, time) => {
  let map = {};
  for (let i = 0; i < data.length; i++) {
    let name = data[i][`商品`];
    let count = data[i][`购买数量`];
    if (!map[name]) {
      map[name] = count;
    } else {
      map[name] = map[name] + count;
    }
  }
  let dataForChef = _.map(_.keys(map), (key => {
    return {
      '商品销售统计': key,
      '数量': map[key],
    };
  }));
  downloadFileToExcel(dataForChef, `厨师表格`, SHEETHEADER_FOR_CHEF, time);
}

const downloadFileToExcel = (data, sheetName, sheetheader, time) => {
  // sheetheader存在时数据和表头不处理
  // 数据
  const dataForCaterer = sheetheader ? data : _.map(data, item => _.pick(item, _.difference(Object.keys(item), SHEETINFOS[sheetName].useless)));
  // 表头
  const sheetData = sheetheader || _.filter(SHEETHEADER, item => !SHEETINFOS[sheetName].useless.includes(item.name));
  // 创建工作簿和工作表
  const workbook = XLSX.utils.book_new(`${sheetName}(${time})`);
  const worksheet = XLSX.utils.json_to_sheet(dataForCaterer, { header: _.map(sheetData, 'name') });

  // 设置行宽
  worksheet['!cols'] = _.map(sheetData, obj => ({ width: obj.cellWidth * 2.3 }));
  let merges = [];
  let mergeCol = [];
  if (sheetName.indexOf('厨师') === -1) {
    _.map(sheetData, (obj, index) => {
      if (obj.isMerge) {
        mergeCol.push(index);
      }
    });
    _.map(mergeCol, col => {
      merges = [...merges, ..._.map(MERGES, obj => ({ s: { ...obj.s, c: col }, e: { ...obj.e, c: col }}))];
    })
    // 合并单元格
    worksheet['!merges'] = merges;
  }
  // 将工作表添加到工作簿
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

  // 导出 Excel 文件
  XLSX.writeFile(workbook, `${sheetName}(${time}).xlsx`);
};

const onImportProductsExcel = (file, time) => {
  // 获取上传的文件对象
  const { files } = file.target;
  // 通过FileReader对象读取文件
  const fileReader = new FileReader();
  fileReader.onload = event => {
    try {
      const { result } = event.target;
      // 以二进制流方式读取得到整份excel表格对象
      const workbook = XLSX.read(result, { type: 'binary' });
      let data = []; // 存储获取到的数据
      // 遍历每张工作表进行读取（这里默认只读取第一张表）
      for (const sheet in workbook.Sheets) {
        if (workbook.Sheets.hasOwnProperty(sheet)) {
          // 利用 sheet_to_json 方法将 excel 转成 json 数据
          data = data.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]), { raw: true });
          break; // 如果只取第一张表，就取消注释这行
        }
      }
      dealProductsData(data, time);
    } catch (e) {
      console.log(e);
      // 这里可以抛出文件类型错误不正确的相关提示
      window.alert(`'文件类型不正确' + ${e}`);
      return;
    }
  };
  // 以二进制方式打开文件
  fileReader.readAsBinaryString(files[0]);
};

const onImportExcel = (file, time) => {
  // 获取上传的文件对象
  const { files } = file.target;
  // 通过FileReader对象读取文件
  const fileReader = new FileReader();
  fileReader.onload = event => {
    try {
      const { result } = event.target;
      // 以二进制流方式读取得到整份excel表格对象
      const workbook = XLSX.read(result, { type: 'binary' });
      let data = []; // 存储获取到的数据
      // 遍历每张工作表进行读取（这里默认只读取第一张表）
      for (const sheet in workbook.Sheets) {
        if (workbook.Sheets.hasOwnProperty(sheet)) {
          // 利用 sheet_to_json 方法将 excel 转成 json 数据
          data = data.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]), { raw: true });
          break; // 如果只取第一张表，就取消注释这行
        }
      }
      dealData(data, time);
    } catch (e) {
      console.log(e);
      // 这里可以抛出文件类型错误不正确的相关提示
      window.alert(`'文件类型不正确' + ${e}`);
      return;
    }
  };
  // 以二进制方式打开文件
  fileReader.readAsBinaryString(files[0]);
};

export default App;
