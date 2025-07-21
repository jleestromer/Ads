# Pinnacle Real Estate Excel Add-in

This is a proof-of-concept Office Add-in implementing the tasks described in `real_estate.xlsx`.

The add-in adds a **Pinnacle Real Estate** tab with a **Run** button that recalculates operating expenses on the `Outputs` sheet and copies them to the `Software Engineer Cash Flow` sheet. A simple UI in the task pane allows overriding total operating expenses.

Due to environment constraints, adjust cell ranges or workbook names as needed in `taskpane.js`.

## Development

1. Run `npm start` to launch a placeholder command (actual hosting not included).
2. Side-load the add-in in Excel using the `manifest.xml` file.

No automated tests are present.
