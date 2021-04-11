const got = require("got");
const xlsx = require("xlsx");
const convert = require("xml-js");

// Target keywords
let keywords = [];

// Reading main.xlsx
let workbook = xlsx.readFile("main.xlsx");
let first_sheet_name = workbook.SheetNames[0]; // Get input sheet
let worksheet = workbook.Sheets[first_sheet_name]; // Get worksheet

// Loop over column A of input sheet
for (let i = 2;;i++) {
  let address_of_cell = `A${i}`; // Get keywords from values of column A
  let desired_cell = worksheet[address_of_cell]; // Find desired cell
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

let URL = `https://google.com/complete/search?output=toolbar&hl=en&q=${keywordCleaner(keywords[2])}`;
console.log(URL);
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
    let ws = xlsx.utils.aoa_to_sheet(ws_data); // Create a new sheet with our data
    xlsx.utils.book_append_sheet(workbook, ws, ws_name); // Appending created sheet to xlsx file
    xlsx.writeFile(workbook, "main.xlsx"); // Write data to xlsx file

    console.log("Let's enjoy! 🥳")
  } catch (error) {
    console.log(error.response.body);
  }
})();