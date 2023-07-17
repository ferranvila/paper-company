const createReport = require("docx-templates").default;
const readXlsxFile = require("read-excel-file/node");
const fs = require("fs").promises;

async function readFile(filename) {
  console.log(`Reading file: ${filename}`);
  return await fs.readFile(filename);
}

async function writeFile(filename, buffer) {
  console.log(`Writing file ${filename}`);
  await fs.writeFile(filename, buffer);
}

async function readData(filename) {
  return await readXlsxFile(filename);
}

async function replaceVariables(template, data) {
  var replaceData = {};
  data.forEach((element) => {
    replaceData[element[0]] = element[1];
  });
  return await createReport({
    template,
    cmdDelimiter: "##",
    data: replaceData,
  });
}

readFile("input/template.docx").then((input) => {
  readData("input/data.xlsx").then((data) => {
    replaceVariables(input, data).then((output) => {
      writeFile("output/output.docx", output).then(() => {
        console.log("Process finished");
      });
    });
  });
});
