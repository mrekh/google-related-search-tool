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
  if (desired_cell == undefined && i === 2) throw new Error("There is no input keyword. Type your keywords in main.xslx");
  else if (desired_cell == undefined && i > 2) break;
  else keywords.push(keywordCleaner(desired_cell.v.trim())); // desired_cell.v: The value of desired cell
}

// Reading Depth cell
let depth_cell = inputWorksheet["C2"];
let depth_value = Number();
if (depth_cell == undefined) throw new Error("There is no value in main.xlsx depth cell");
else if (depth_cell.v < 0) throw new Error("You depth value in main.xslx at least must be 0");
else depth_value = depth_cell.v;

// Writing data of the Output sheet
function writeXlsx(data, column) {
  let ws_data = data;
  xlsx.utils.sheet_add_aoa(outputWorksheet, ws_data, {origin: {r: 0, c: column}}); // Append data to the sheet
  xlsx.writeFile(workbook, "main.xlsx"); // Write data to xlsx file
  xlsxResult = [];
}

// API URL
function URL(keyword) {
  return `https://google.com/complete/search?output=toolbar&hl=en&q=${keyword}`;
}

// Trim spaces and replace space with '+' between multi words keywords
function keywordCleaner(inputKw) { // inputKw: Input keyword
  if (inputKw.charCodeAt(0) >= 0x0600 && inputKw.charCodeAt(0) <= 0x06FF) inputKw = inputKw.match(/[^\s][\u0600-\u06FF]*/g); // Persian keywrods regex 
  else inputKw = inputKw.match(/\b[^\s][a-z0-9]*\b/gi); // English keywrods regex

  let str = "";
  for (let i = 0; i < inputKw.length; i++) {
    str = str + inputKw[i] + "+";
  }
  return str.slice(0, str.length - 1);
}

// Creating a get request
(async () => {
  try {
    for (let i = 0; i < keywords.length; i++) {
      let response = await got(URL(keywords[i]));
      queryResults[i] = await convert.xml2js(response.body, {compact: true, spaces: 2});
      // Update xlsxResult array which finally write in the Output sheet
      for (let num = 0; num < await queryResults[i].toplevel.CompleteSuggestion.length; num++) {
        xlsxResult.push([queryResults[i].toplevel.CompleteSuggestion[num].suggestion._attributes.data]);
      }

      // Loop for when depth is > 0
      let xRL = [1, xlsxResult.length]; // xlsx result length
      for (let rounds = 1; rounds <= depth_value; rounds++) {
        for (let t = xRL[rounds - 1]; t < xRL[rounds]; t++) {
          response = await got(URL(xlsxResult[t][0]));
          depthQueryResults[t - 1] = await convert.xml2js(response.body, {compact: true, spaces: 2});
          // Update xlsxResult array which finally write in the Output sheet
          for (let numbers = 1; numbers < await depthQueryResults[t - 1].toplevel.CompleteSuggestion.length; numbers++) {
            xlsxResult.push([depthQueryResults[t - 1].toplevel.CompleteSuggestion[numbers].suggestion._attributes.data]);
          }

          xRL[rounds + 1] = xlsxResult.length;
        }
      }

      writeXlsx(xlsxResult, i);
    }

    console.log("Let's enjoy! 🥳") // Final step :)
  } catch (err) {
    if (err == "RequestError: getaddrinfo ENOTFOUND google.com") console.log("There is a network conection issue");
    else console.log(err);
  }
})();