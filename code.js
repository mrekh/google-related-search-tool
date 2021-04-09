const got = require("got");
const xlsx = require("xlsx");
const convert = require("xml-js");

let keywords = [];
// Reading main.xlsx
let workbook = xlsx.readFile("main.xlsx");
// Get input sheet
let first_sheet_name = workbook.SheetNames[0];
// Get worksheet
let worksheet = workbook.Sheets[first_sheet_name];

// Loop over column A of input sheet
for (let i = 2;;i++) {
  // Get keywords from values of column A
  let address_of_cell = `A${i}`;
  // Find desired cell
  let desired_cell = worksheet[address_of_cell];
  // Get the value
  let disired_value = (desired_cell ? desired_cell.v : undefined);
  if (disired_value !== undefined) {
    keywords.push(disired_value);
  } else break;
}
console.log(keywords);

// Trim spaces and replace space with '+' between multi words keywords
function keywordCleaner(inputKeyword) {
  let kw = inputKeyword.trim().match(/\b[^\s][a-z0-9]*\b/gi);
  let str = "";
  for (let i = 0; i < kw.length; i++) {
    str = str + kw[i] + "+";
  }
  return str.slice(0, str.length - 1);
}

let URL = `https://google.com/complete/search?output=toolbar&hl=en&q=${keywordCleaner(keywords[0])}`;
let queryResults = [];
let xlsxResult = [];

// Name of sheet we want to create in our xlsx file
let ws_name = "Output";

// Creating a get request
(async () => {
  try {
    const response = await got(URL);
    queryResults[0] = await convert.xml2js(response.body, {compact: true, spaces: 2});

    for (let num = 0; num < await queryResults[0].toplevel.CompleteSuggestion.length - 1; num++) {
      xlsxResult.push([queryResults[0].toplevel.CompleteSuggestion[num].suggestion._attributes.data]);
    }

    // Data of created sheet
    let ws_data = await xlsxResult;
    // Create a new sheet with our data
    let ws = xlsx.utils.aoa_to_sheet(ws_data);
    // Appending created sheet to xlsx file
    xlsx.utils.book_append_sheet(workbook, ws, ws_name);
    // Write data to xlsx file
    xlsx.writeFile(workbook, "main.xlsx");

    console.log("Let's enjoy! 🥳")
  } catch (error) {
    console.log(error.response.body);
  }
})();