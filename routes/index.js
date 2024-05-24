const express = require("express");
const fs = require("fs");
const router = express.Router();
const multer = require("multer");
const xlsx = require("node-xlsx");
const ExcelJS = require("exceljs");
const moment = require("moment");
// const fileid = {};
// const btn = ["btn-primary","btn-success","btn-danger","btn-warning","btn-info","btn-dark"];

// function rand(){
//   return Math.floor(Math.random()*6)
// }

// 文件上传
const upload = multer({
  storage: multer.diskStorage({
    limits: {
      fileSize: 10 * 10000000,
    },
    destination: function (req, file, cb) {
      cb(null, "./public/file");
    },
    filename: function (req, file, cb) {
      // console.log(file,'file');
      // const time = new Date().getTime()
      const changedName = file.originalname;
      // const changedName = fileid[file.originalname]+'.'+file.originalname.split(".")[file.originalname.split(".").length-1];
      cb(null, changedName);
    },
  }),
});

// 单个文件上传
router.post("/upload", upload.single("file"), (req, res) => {
  try {
    res.status(200).json({
      fileName: req.file.filename,
    });
  } catch (error) {
    res.status(500).json({
      error: error.message,
    });
  }
});

router.post("/uploads", upload.array("file",2), (req, res) => {
  try {
    res.status(200).json({
      fileName: req.files,
    });
  } catch (error) {
    res.status(500).json({
      error: error.message,
    });
  }
});

// 下载筛选数据
router.get("/download", async (req, res, next) => {
  const lastData = req.query.lastData;
  const fileName = req.query.fileName;
  const moduleNumber = req.query.moduleNumber;
  let OperatorsNum = []; // 运营商代号
  if (moduleNumber.includes("中国移动")) {
    OperatorsNum = OperatorsNum.concat([
      "898600",
      "898602",
      "898604",
      "898607",
    ]);
  }
  if (moduleNumber.includes("中国联通")) {
    OperatorsNum = OperatorsNum.concat(["898601", "898606", "898609"]);
  }
  if (moduleNumber.includes("中国电信")) {
    OperatorsNum = OperatorsNum.concat(["898603", "898611"]);
  }
  console.log(OperatorsNum, "OperatorsNum");

  const path = "./public/file/" + fileName; // 文件地址
  // 读取
  const workSheetsFromFile = xlsx.parse(path);
  const data = workSheetsFromFile[0].data;
  // 筛选满足的行
  // const filteredData = data.filter((row) => row[19] == 1020); row[26]
  const filteredData = data.filter((row) => {
    return (
      moment(row[19], ["YYYY-MM-DD HH:mm:ss"]).isAfter(lastData) &&
      OperatorsNum.indexOf(row[26].slice(0, 6)) !== -1
    );
  });
  // 日期比较
  // function compareMoments(moment1, moment2) {
  //   if (moment1.isBefore(moment2)) {
  //     return -1; // moment1在moment2之前
  //   } else if (moment1.isAfter(moment2)) {
  //     return 1; // moment1在moment2之后
  //   } else {
  //     return 0; // moment1和moment2相同
  //   }
  // }
  // 创建一个新的Excel工作簿
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");
  // 添加表头
  worksheet.columns = [
    { header: "小区名" },
    { header: "楼栋号" },
    { header: "单元号" },
    { header: "门牌号" },
    { header: "户表编号" },
    { header: "供水温度(℃)" },
    { header: "回水温度(℃)" },
    { header: "供水压力" },
    { header: "回水压力" },
    { header: "累计工作时间(小时)" },
    { header: "流速(m3/h)" },
    { header: "累计流量(m3)" },
    { header: "阀门状态" },
    { header: "无线模块" },
    { header: "发送间隔" },
    { header: "发送序号" },
    { header: "信号强度" },
    { header: "频段" },
    { header: "采集时间" },
    { header: "保存时间" },
    { header: "供电电压(V)" },
    { header: "上传总次数" },
    { header: "上传成功次数" },
    { header: "抄表总次数" },
    { header: "抄表成功次数" },
    { header: "NB模块IMEI" },
    { header: "SIM卡ICCID" },
  ];
  filteredData.forEach((item) => {
    worksheet.addRow(item);
  });
  // 设置响应头，定义文件名并生成Excel文件
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=UTF-8"
  );
  res.setHeader(
    "Content-Disposition",
    "attachment; filename=filtered_data.xlsx"
  );
  await workbook.xlsx.write(res);
  res.end();
});

// 两个excel做差集
router.get("/twoDownload", async (req, res) => {
  const firstData = req.query.firstData
  const secondData = req.query.secondData
  console.log(firstData,'firstData');
  console.log(secondData,'secondData');
  if (!firstData) {
    return res.status(400).send("请上传两个Excel文件");
  }
  if (!secondData) {
    return res.status(400).send("请上传两个Excel文件");
  }
  try {
    const firstPath = "./public/file/" + firstData; // 文件地址
    const secondPath = "./public/file/" + secondData; // 文件地址
    // console.log(firstPath,'firstPath');
    // console.log(secondPath,'secondPath');
    // 读取第一个Excel文件
    const workbook1 = new ExcelJS.Workbook();
    await workbook1.xlsx.readFile(firstPath);
    const worksheet1 = workbook1.getWorksheet(1);
    const headers = worksheet1.getRow(1).values.slice(1); // 获取表头
    const data1 = new Set();
    worksheet1.eachRow((row, rowNumber) => {
      if(rowNumber > 1){
        data1.add(row.values.slice(1).join(",")); // 假设每行数据是唯一的
      }
    });

    // 读取第二个Excel文件
    const workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.readFile(secondPath);
    const worksheet2 = workbook2.getWorksheet(1);
    const data2 = new Set();
    worksheet2.eachRow((row, rowNumber) => {
      if(rowNumber > 1){
        data2.add(row.values.slice(1).join(",")); // 假设每行数据是唯一的
      }
    });

    // 计算差集
    const difference1 = [...data1].filter((item) => !data2.has(item));
    const difference2 = [...data2].filter((item) => !data1.has(item));

    // 创建一个新的Excel工作簿
    const resultWorkbook = new ExcelJS.Workbook();
    const resultWorksheet = resultWorkbook.addWorksheet("Difference");
    // 添加表头
    resultWorksheet.addRow(headers)

    const differenceAll = [...difference1,...difference2]
    // 添加差集数据到新的工作簿
    differenceAll.forEach((item, index) => {
      resultWorksheet.addRow(item.split(","));
    });

    // 设置响应头，定义文件名并生成Excel文件
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=difference.xlsx"
    );

    await resultWorkbook.xlsx.write(res);
    res.end();
  } catch (error) {
    res.status(500).json({
      error: error.message,
    });
  }
});

/* GET users listing. */
router.get("/", function (req, res, next) {
  res.render("index");
});

module.exports = router;
