# Pinnacle Real Estate Excel Add-in

This project contains a lightweight Office Add‑in that demonstrates the workflow described in `real_estate.xlsx`.

## Features

- Adds a **Pinnacle Real Estate** tab with a **Run** button.
- Recalculates only the operating expenses on the `Outputs` sheet even when Excel is in Manual Calculation mode.
- Copies those calculated values to the `Software Engineer Cash Flow` sheet while preserving formatting.
- Task pane UI allows you to manually override the total operating expenses if needed.
- Bonus logic generates a column chart of the expenses and optionally inserts a building image.

## Installation

1. Ensure you have [Node.js](https://nodejs.org/) installed.
2. In this folder run `npm install` if additional packages are added in the future (none are required for the current proof of concept).
3. Launch the local web server placeholder with `npm start`.
4. Side‑load the add-in in Excel using `manifest.xml` (see Microsoft documentation for side-loading Office Add-ins).

## Usage

1. Open `real_estate.xlsx` in Excel.
2. After side‑loading the add-in, a **Pinnacle Real Estate** tab appears in the ribbon.
3. Press **Run** to recalculate the operating expenses and update the cash flow tab.
4. To override the total operating expenses, open the task pane, enter a value, and press **Apply Override**.

Ranges and worksheet names referenced in `taskpane.js` may require adjustment depending on your workbook. Review that file if your model layout differs from the example workbook.

## Disclaimer

The Excel model and this add‑in are provided solely for case‑study purposes. Pinnacle Real Estate and its affiliates make no representation or warranty regarding the accuracy or completeness of the model. Use it at your own risk and do not rely on it for any real‑world financial decisions.

## Development Notes

- The `npm start` script currently prints a placeholder message. In a production scenario you would host the task pane files on a web server.
- No automated tests are included.

