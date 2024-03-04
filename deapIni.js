const fs = require("fs");
const { exec } = require("child_process");
const ini = require("ini");

const iniData = fs
  .readFileSync(`${process.cwd()}/AutoDeapSettings.ini`, "utf-8")
  .toString();
const setting = ini.parse(iniData);
const excelDataFileName = setting["General"].ExcelDataFileName;
const excelOutputFileName = setting["General"].ExcelOutputFileName;
const outputCategoryList = setting["General"].OutputCategoryList.split("|");
const insObj = setting["instruction"];

const excelDataFilePath = `${process.cwd()}/${excelDataFileName}`;
let data = fs.readFileSync(excelDataFilePath, "utf-8").toString();

data = data.split("\r\n").map((v) => v.replace(" ", "").split("\t"));
let categoryList = new Set();

let categoryOld = "";
for (const dataArr of data) {
  let category = dataArr[0];
  if (category === "") break;
  categoryList.add(category);
  const dataFilePath = `${process.cwd()}/${category}.txt`;
  const insFilePath = `${process.cwd()}/${category}`;
  const firmsReg = new RegExp(`${category},`, "g");
  const firms = data.toString().match(firmsReg)?.length;
  insObj.Firms = firms;
  if (category !== categoryOld) {
    fs.writeFileSync(dataFilePath, "");
    const ins =
      `${category}.txt\n` +
      `${category}.yaml\n` +
      `${insObj.Firms}\n` +
      `${insObj.TimePeriods}\n` +
      `${insObj.Outputs}\n` +
      `${insObj.Inputs}\n` +
      `${insObj.InputOrOutput}\n` +
      `${insObj.CRS_VRS}\n` +
      `${insObj.Model}`;
    fs.writeFileSync(insFilePath, ins);
  }
  const dataStr = dataArr.toString().slice(5).replaceAll(",", " ");
  fs.appendFileSync(dataFilePath, `${dataStr}\n`);
  categoryOld = category;
}

// 对所有数据执行 deap.exe
const exePath = `${process.cwd()}/DEAP.EXE`;

categoryList = Array.from(categoryList);
for (const category of categoryList) {
  while (true) {
    const index = categoryList.indexOf(category);
    if (
      fs.existsSync(`${process.cwd()}/${category}`) &&
      fs.existsSync(`${process.cwd()}/${category}.txt`)
    ) {
      if (index === 0) break;
      if (fs.existsSync(`${process.cwd()}/${categoryList[index - 1]}.yaml`)) {
        break;
      }
    }
  }
  // 创建子进程
  const childProcess = exec(exePath, (error, stdout, stderr) => {
    if (error) {
      console.error(`exec error: ${error}`);
      return;
    }
  });

  // 向子进程输入
  const inputText = `${category}\n`;
  childProcess.stdin.write(inputText);
}

// 从 out 文件找到所需数据
const excelOutputFilePath = `${process.cwd()}/${excelOutputFileName}`;
fs.writeFileSync(excelOutputFilePath, "");
for (const category of categoryList) {
  const outputFilePath = `${process.cwd()}/${category}.yaml`;
  while (true) {
    if (
      fs.existsSync(outputFilePath) &&
      fs.readFileSync(outputFilePath, "utf-8").toString() !== ""
    ) {
      break;
    }
  }

  let outputData = fs.readFileSync(outputFilePath, "utf-8").toString();
  outputData = outputData.slice(1).replaceAll(`\r\n`, `\n`);
  outputData = outputData.replaceAll(/\n +/g, "\n");
  outputData = outputData.replaceAll(/ +\n/g, "\n");
  let allCategoryData = outputData.split(`\n\n\n`);

  let data = [];
  for (const column of outputCategoryList) {
    for (const category of allCategoryData) {
      if (category.slice(0, column.length) === column) {
        let categoryData = category.split(`\n`);
        categoryData = categoryData.filter((v) => !isNaN(Number(v[0])));
        data.push(categoryData);
      }
    }
  }

  // 整理数据成 csv 文件
  for (let i = 0; i < insObj.Firms; i++) {
    let element = `${category} ${data[0][i]}`;
    for (let y = 1; y < data.length; y++) {
      const eachDataElement = data[y];
      element += " " + eachDataElement[i];
    }
    element = element.replaceAll(/ +/g, ",");
    fs.appendFileSync(`${outputFilePath.slice(0, -4)}csv`, `${element}\n`);
    fs.appendFileSync(excelOutputFilePath, `${element}\n`);
  }
}

// 整理目录
// 获取当前时间
const date = new Date().toLocaleString("zh-CN");
// 此次运行目录
const dirName = date
  .replaceAll("/", "-")
  .replaceAll(" ", "_")
  .replaceAll(":", "-");

fs.mkdirSync(`${dirName}`);
fs.mkdirSync(`${dirName}/data`);
// 移动文件
for (const category of categoryList) {
  const fileNameObj = {
    dtaName: `${category}.txt`,
    insName: `${category}`,
    outName: `${category}.yaml`,
    csvName: `${category}.csv`,
  };
  for (const fileCategory in fileNameObj) {
    const fileName = fileNameObj[fileCategory];
    while (true) {
      if (fs.existsSync(`${process.cwd()}/${fileName}`)) {
        break;
      }
    }
    fs.renameSync(
      `${process.cwd()}/${fileName}`,
      `${process.cwd()}/${dirName}/data/${fileName}`
    );
  }
}
fs.copyFileSync(
  `${process.cwd()}/${excelDataFileName}`,
  `${process.cwd()}/${dirName}/${excelDataFileName}`
);
fs.renameSync(
  excelOutputFilePath,
  `${process.cwd()}/${dirName}/${excelOutputFileName}`
);
console.log(`文件已保存于目录：${process.cwd()}\\${dirName}`);
