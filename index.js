const rp = require("request-promise");
const cheerio = require("cheerio");
const XLSX = require("xlsx");
const fs = require("fs");

async function getSpecies() {
  let url = "";
  fs.readFile("url.txt", "utf8", (err, data) => {
    if (err) throw err;
    url = data.trim();
    console.log(url);
    rp(url).then(function (html) {
      const $ = cheerio.load(html);
      const speciesNames = [];

      $("em").each(function (i, elem) {
        const species = $(this).text().split(" ")[1];
        speciesNames.push({ Species: species });
      });

      let workbook;
      try {
        workbook = XLSX.readFile("existing_file.xlsx");
      } catch (error) {
        console.log("newfile");
        workbook = XLSX.utils.book_new();
      }

      let worksheet = workbook.Sheets["Sheet1"];
      if (!worksheet) {
        worksheet = XLSX.utils.json_to_sheet(
          [{ Species: "Species" }].concat(speciesNames)
        );
        XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
      } else {
        const range = XLSX.utils.decode_range(worksheet["!ref"]);
        let lastRow = 0;
        for (let i = 1; i <= 700; i++) {
          //   console.log(worksheet[`B${i}`]);
          if (!worksheet[`F${i}`]) {
            lastRow = i;
            console.log("broke at ", i);
            break;
          }
        }
        XLSX.utils.sheet_add_json(worksheet, speciesNames, {
          header: ["Species"],
          origin: `F${lastRow}`,
        });
      }

      try {
        XLSX.writeFile(workbook, "existing_file.xlsx");
        console.log("success");
      } catch (error) {
        console.error("HERE IS THE ERROR", error);
      }
    });
  });
}

getSpecies();
