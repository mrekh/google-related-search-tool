# Google Related Search Tool

A Node.js tool to fetch Google Autocomplete suggestions based on keywords from an Excel file (`main.xlsx`) and write the results back to the file.

## BIG UPDATE 🥳
Now you can run the code within Google Sheets!
You can find more information about this Google Sheets in my LinkedIn post.
### [Get Google Related Searches in Google Sheets](https://www.linkedin.com/posts/alireza-esmikhani_seo-googleabrads-searchengineoptimization-activity-6837228087973834752-NqxD)

## Features

- Reads keywords from column A of the first sheet in `main.xlsx`.
- Reads the desired suggestion depth from cell C2 in the first sheet (0 means only direct suggestions for the keyword, 1 means suggestions for the keyword AND suggestions for those suggestions, etc.).
- Fetches Google Autocomplete suggestions using the specified keywords and depth.
- Writes the unique suggestions found for each keyword into separate columns in the second sheet of `main.xlsx`.
- Uses native `fetch` for HTTP requests.
- Uses `pnpm` for package management.

## Prerequisites

- [Node.js](https://nodejs.org/) (version 18 or later recommended, as it includes native `fetch`)
- [pnpm](https://pnpm.io/installation)

## Setup

1.  **Clone the repository:**
    ```bash
    git clone <repository-url>
    cd google-related-search-tool
    ```

2.  **Install dependencies:**
    ```bash
    pnpm install
    ```

3.  **Prepare the Excel file (`main.xlsx`):**
    - Make sure you have an Excel file named `main.xlsx` in the project root.
    - **Sheet 1 (Input):**
        - **Column A:** List your target keywords starting from cell `A2`.
        - **Cell C2:** Enter the desired depth (e.g., `0`, `1`, `2`). If empty or invalid, it defaults to `0`.
    - **Sheet 2 (Output):**
        - This sheet will be created or overwritten with the results.
        - Each column will contain the suggestions for the corresponding keyword from Sheet 1 (Column A). The keyword itself will be in the first row of its respective column.

## Usage

1.  **Build the TypeScript code:**
    ```bash
    pnpm build
    ```
    This compiles `app.ts` into JavaScript in the `dist` directory.

2.  **Run the tool:**
    ```bash
    pnpm start
    ```
    Alternatively, for development, you can run the TypeScript file directly:
    ```bash
    pnpm dev
    ```

The tool will read `main.xlsx`, fetch the suggestions, and update the second sheet in `main.xlsx` with the results.

## License

ISC
