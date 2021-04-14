const got = require("got");
const xlsx = require("xlsx");
const convert = require("xml-js");

let keywords = []; // Target keywords
let queryResults = []; // Response body for each keyword
let depthQueryResults = []; // Response body for each keyword - depth > 0
let xlsxResult = []; // Value we want to write in our xlsx

// Reading main.xlsx
let workbook = xlsx.readFile("main.xlsx");
let inputWorksheet = workbook.Sheets[workbook.SheetNames[0]]; // Get Input worksheet
let outputWorksheet = workbook.Sheets[workbook.SheetNames[1]]; // Get Output wordsheet

// Loop over column A of input sheet
for (let i = 2;;i++) {
  let desired_cell = inputWorksheet[`A${i}`]; // Get keywords from values of column A & Find desired cell
  let disired_value = desired_cell ? desired_cell.v : undefined; // Get the value
  if (disired_value !== undefined) {
    keywords.push(disired_value);
  } else break;
}

// Reading Depth cell
let depth_cell = inputWorksheet["C2"];
let depth_value = depth_cell ? depth_cell.v : undefined;

// Trim spaces and replace space with '+' between multi words keywords
function keywordCleaner(inputKeyword) {
  let kw = inputKeyword.trim();
  if (kw.charCodeAt(0) >= 0x0600 && kw.charCodeAt(0) <= 0x06FF) {
    kw = kw.match(/[^\s][\u0600-\u06FF]*/g); // Persian keywrods regex
  } else { kw = kw.match(/\b[^\s][a-z0-9]*\b/gi); } // English keywrods regex

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
      let response = await got(URL);
      queryResults[i] = await convert.xml2js(response.body, {compact: true, spaces: 2});

      for (let num = 0; num < await queryResults[i].toplevel.CompleteSuggestion.length; num++) {
        xlsxResult.push([queryResults[i].toplevel.CompleteSuggestion[num].suggestion._attributes.data]);
      }

      if (depth_value !== 0) {
        let xRL = [1, xlsxResult.length];

        for (let rounds = 1; rounds <= depth_value; rounds++){
          for (let t = xRL[rounds - 1]; t < xRL[rounds]; t++) {
            URL = `https://google.com/complete/search?output=toolbar&hl=en&q=${keywordCleaner(xlsxResult[t][0])}`;
            response = await got(URL);
            depthQueryResults[t - 1] = await convert.xml2js(response.body, {compact: true, spaces: 2});
            
            for (let numbers = 1; numbers < await depthQueryResults[t - 1].toplevel.CompleteSuggestion.length; numbers++) {
                xlsxResult.push([depthQueryResults[t - 1].toplevel.CompleteSuggestion[numbers].suggestion._attributes.data]);
            }
      
            xRL[rounds + 1] = xlsxResult.length;
          }
        }
      }

      // Writing data of Output sheet
      let ws_data = await xlsxResult;
      xlsx.utils.sheet_add_aoa(outputWorksheet, ws_data, {origin: {r: 0, c: i}}); // Append data to the sheet
      xlsx.writeFile(workbook, "main.xlsx"); // Write data to xlsx file
      xlsxResult = [];
    }

    console.log("Let's enjoy! 🥳")
  } catch (error) {
    console.log(error.response.body);
  }
})();