import * as xlsx from "xlsx";
import * as convert from "xml-js";

interface Suggestion {
  suggestion: {
    _attributes: {
      data: string;
    };
  };
}

interface XML2JSResult {
  toplevel?: {
    CompleteSuggestion?: Suggestion | Suggestion[];
  };
}

// --- Configuration ---
const EXCEL_FILE_PATH = "main.xlsx";
const INPUT_SHEET_INDEX = 0; // First sheet for input
const OUTPUT_SHEET_INDEX = 1; // Second sheet for output
const KEYWORD_COLUMN = "A";
const DEPTH_CELL = "C2";

// --- Helper Functions ---

/**
 * Builds the Google Autocomplete API URL for a given keyword.
 * @param keyword The search keyword.
 * @returns The formatted API URL.
 */
function buildApiUrl(keyword: string): string {
  // Encode the keyword to handle special characters properly in the URL
  const encodedKeyword = encodeURIComponent(keyword.replace(/\+/g, " ")); // Decode '+' first if it was used for spaces
  return `https://google.com/complete/search?output=toolbar&hl=en&q=${encodedKeyword}`;
}

/**
 * Cleans and formats a keyword for the API query.
 * Replaces spaces with '+' and handles basic Persian/English word extraction.
 * @param inputKw Raw keyword string.
 * @returns Cleaned keyword string suitable for the API.
 */
function keywordCleaner(inputKw: string): string {
  let cleanedKw = inputKw.trim();
  // Basic regex; might need refinement for edge cases
  const words = cleanedKw.match(/[^\s]+/g) || []; // Match non-space sequences
  return words.join("+"); // Join words with '+'
}

/**
 * Fetches suggestions from the Google Autocomplete API.
 * @param keyword The keyword to search for.
 * @returns A promise resolving to the parsed XML response, or null on error.
 */
async function fetchSuggestions(keyword: string): Promise<XML2JSResult | null> {
  const url = buildApiUrl(keyword);
  try {
    const response = await fetch(url);
    if (!response.ok) {
      console.error(
        `HTTP error! status: ${response.status} for keyword: ${keyword}`
      );
      return null;
    }
    const xmlText = await response.text();
    // Type assertion needed as xml-js types might not be perfectly aligned
    return convert.xml2js(xmlText, { compact: true }) as XML2JSResult;
  } catch (error) {
    console.error(`Failed to fetch suggestions for "${keyword}":`, error);
    if (error instanceof Error && error.message.includes("ENOTFOUND")) {
      console.error("Network connection issue or invalid domain.");
    }
    return null;
  }
}

/**
 * Writes results to the specified column in the output sheet.
 * @param workbook The XLSX workbook object.
 * @param outputSheet The target worksheet object.
 * @param data Data to write (array of arrays).
 * @param columnIndex The zero-based column index to start writing at.
 */
function writeResultsToSheet(
  workbook: xlsx.WorkBook,
  outputSheet: xlsx.WorkSheet,
  data: string[][],
  columnIndex: number
): void {
  if (data.length > 0) {
    xlsx.utils.sheet_add_aoa(outputSheet, data, {
      origin: { r: 0, c: columnIndex },
    });
    xlsx.writeFile(workbook, EXCEL_FILE_PATH);
  } else {
    console.warn(`No data to write for column index ${columnIndex}`);
  }
}

/**
 * Extracts suggestion strings from the parsed API response.
 * @param parsedResponse The parsed XML response object.
 * @returns An array of suggestion strings, or an empty array if none found or error.
 */
function extractSuggestions(parsedResponse: XML2JSResult | null): string[] {
  if (!parsedResponse?.toplevel?.CompleteSuggestion) {
    return [];
  }
  // Handle cases where there might be single or multiple suggestions
  const suggestions = Array.isArray(parsedResponse.toplevel.CompleteSuggestion)
    ? parsedResponse.toplevel.CompleteSuggestion
    : [parsedResponse.toplevel.CompleteSuggestion];

  return suggestions
    .map((suggestion) => suggestion?.suggestion?._attributes?.data)
    .filter(Boolean) as string[];
}

// --- Main Logic ---

async function main() {
  console.log("Starting Google Autocomplete suggestion fetch...");

  // 1. Read Workbook and Sheets
  let workbook: xlsx.WorkBook;
  try {
    workbook = xlsx.readFile(EXCEL_FILE_PATH);
  } catch (error) {
    console.error(`Error reading Excel file "${EXCEL_FILE_PATH}":`, error);
    return; // Exit if file can't be read
  }

  const inputSheet = workbook.Sheets[workbook.SheetNames[INPUT_SHEET_INDEX]];
  let outputSheet = workbook.Sheets[workbook.SheetNames[OUTPUT_SHEET_INDEX]];

  if (!inputSheet) {
    console.error(`Input sheet at index ${INPUT_SHEET_INDEX} not found.`);
    return;
  }
  // Create output sheet if it doesn't exist
  if (!outputSheet) {
    console.log(
      `Output sheet at index ${OUTPUT_SHEET_INDEX} not found. Creating it.`
    );
    outputSheet = xlsx.utils.aoa_to_sheet([]); // Create an empty sheet
    xlsx.utils.book_append_sheet(
      workbook,
      outputSheet,
      workbook.SheetNames[OUTPUT_SHEET_INDEX] ||
        `Output ${OUTPUT_SHEET_INDEX + 1}`
    );
  }

  // 2. Read Keywords from Input Sheet
  const initialKeywords: string[] = [];
  for (let i = 2; ; i++) {
    // Start from row 2 (assuming headers in row 1)
    const cellAddress = `${KEYWORD_COLUMN}${i}`;
    const cell = inputSheet[cellAddress];
    if (
      !cell ||
      cell.v === undefined ||
      cell.v === null ||
      String(cell.v).trim() === ""
    ) {
      // Stop if cell is empty or doesn't exist
      if (i === 2 && initialKeywords.length === 0) {
        console.error(
          `No input keywords found in column ${KEYWORD_COLUMN} starting from row 2.`
        );
        return;
      }
      break; // End of keywords
    }
    initialKeywords.push(String(cell.v));
  }

  if (initialKeywords.length === 0) {
    console.error("No valid keywords were read from the input sheet.");
    return;
  }

  // 3. Read Depth Value
  const depthCell = inputSheet[DEPTH_CELL];
  let depthValue = 0; // Default depth
  if (!depthCell || depthCell.v === undefined || depthCell.v === null) {
    console.warn(
      `Depth cell ${DEPTH_CELL} not found or empty. Using default depth 0.`
    );
  } else {
    const parsedDepth = Number(depthCell.v);
    if (isNaN(parsedDepth) || parsedDepth < 0) {
      console.warn(
        `Invalid depth value "${depthCell.v}" in ${DEPTH_CELL}. It must be a non-negative number. Using default depth 0.`
      );
    } else {
      depthValue = Math.floor(parsedDepth); // Ensure integer depth
    }
  }
  console.log(`Using depth: ${depthValue}`);

  // 4. Process Each Initial Keyword
  for (let colIndex = 0; colIndex < initialKeywords.length; colIndex++) {
    const baseKeyword = initialKeywords[colIndex];
    console.log(
      `\nProcessing keyword: "${baseKeyword}" (Column ${colIndex + 1})`
    );

    let currentKeywordsToProcess = [keywordCleaner(baseKeyword)]; // Start with the cleaned initial keyword
    const allResultsForColumn: Set<string> = new Set(); // Use Set to avoid duplicates

    // Depth 0 fetch
    const initialResponse = await fetchSuggestions(currentKeywordsToProcess[0]);
    const initialSuggestions = extractSuggestions(initialResponse);
    initialSuggestions.forEach((suggestion) =>
      allResultsForColumn.add(suggestion)
    );
    let nextLevelKeywords = initialSuggestions; // Keywords for the next depth level

    console.log(`  Depth 0: Found ${initialSuggestions.length} suggestions.`);

    // Process deeper levels
    for (let depth = 1; depth <= depthValue; depth++) {
      console.log(`  Processing Depth ${depth}...`);
      const keywordsForThisLevel = [...nextLevelKeywords]; // Copy keywords for this level
      nextLevelKeywords = []; // Reset for the *next* level

      if (keywordsForThisLevel.length === 0) {
        console.log(`    No further keywords to process at depth ${depth}.`);
        break; // Stop if no keywords from the previous level
      }

      const promises = keywordsForThisLevel.map((kw) =>
        fetchSuggestions(keywordCleaner(kw))
      );
      const responses = await Promise.all(promises);

      let suggestionsThisLevel = 0;
      responses.forEach((response) => {
        const suggestions = extractSuggestions(response);
        suggestions.forEach((suggestion) => {
          if (!allResultsForColumn.has(suggestion)) {
            allResultsForColumn.add(suggestion);
            nextLevelKeywords.push(suggestion); // Add new, unique suggestions for the next depth
            suggestionsThisLevel++;
          }
        });
      });
      console.log(`    Found ${suggestionsThisLevel} new unique suggestions.`);
    }

    // 5. Prepare and Write Results for the Column
    const finalResultsArray: string[][] = [
      [baseKeyword],
      ...Array.from(allResultsForColumn).map((s) => [s]),
    ]; // Add original keyword as header
    console.log(
      `Writing ${
        allResultsForColumn.size
      } unique suggestions for "${baseKeyword}" to column ${colIndex + 1}.`
    );
    writeResultsToSheet(workbook, outputSheet, finalResultsArray, colIndex);
  }

  console.log("\nProcessing complete. Let's enjoy! 🥳");
}

// --- Run Main Function ---
main().catch((err) => {
  console.error("\nAn unexpected error occurred:", err);
  // Check for specific network error types if needed
  if (err instanceof Error && err.message.includes("fetch")) {
    console.error(
      "This might be a network issue or a problem reaching the Google API."
    );
  }
});
