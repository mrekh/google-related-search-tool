const got = require("got");
const xlsx = require("xlsx");
const convert = require("xml-js");

// Target keywords
let keywords = [];
let queryResults = []; // Response body for each keyword
let xlsxResult = []; // Value we want to write in our xlsx

// Reading main.xlsx
let workbook = xlsx.readFile("main.xlsx");
let first_sheet_name = workbook.SheetNames[0]; // First sheet in workbook - Input
let inputWorksheet = workbook.Sheets[first_sheet_name]; // Get Input worksheet
let outputWorksheet = workbook.Sheets[workbook.SheetNames[1]]; // Get Output wordsheet

// Loop over column A of input sheet
for (let i = 2;;i++) {
  let address_of_cell = `A${i}`; // Get keywords from values of column A
  let desired_cell = inputWorksheet[address_of_cell]; // Find desired cell
  let disired_value = (desired_cell ? desired_cell.v : undefined); // Get the value
  if (disired_value !== undefined) {
    keywords.push(disired_value);
  } else break;
}

// Trim spaces and replace space with '+' between multi words keywords
function keywordCleaner(inputKeyword) {
  let kw = inputKeyword.trim();
  if (kw.charCodeAt(0) >= 0x0600 && kw.charCodeAt(0) <= 0x06FF) {
    kw = kw.match(/[^\s][\u0600-\u06FF]*/g); // Persian keywrods regex
  } else {
    kw = kw.match(/\b[^\s][a-z0-9]*\b/gi); // English keywrods regex
  }
  let str = "";
  for (let i = 0; i < kw.length; i++) {
    str = str + kw[i] + "+";
  }
  return str.slice(0, str.length - 1);
}

// Creating a get request
(async () => {
  try {
    for (let i = 0; i < keywords.length; i++) {
      let URL = `https://google.com/complete/search?output=toolbar&hl=en&q=${keywordCleaner(keywords[i])}`;
      const response = await got(URL);
      queryResults[i] = await convert.xml2js(response.body, {compact: true, spaces: 2});

      for (let num = 0; num < await queryResults[i].toplevel.CompleteSuggestion.length - 1; num++) {
        xlsxResult.push([queryResults[i].toplevel.CompleteSuggestion[num].suggestion._attributes.data]);
      }
    }

    // Data of created sheet
    let ws_data = await xlsxResult;
    xlsx.utils.sheet_add_aoa(outputWorksheet, ws_data); // Append data to a sheet
    xlsx.writeFile(workbook, "main.xlsx"); // Write data to xlsx file

    console.log("Let's enjoy! 🥳")
  } catch (error) {
    console.log(error.response.body);
  }
})();