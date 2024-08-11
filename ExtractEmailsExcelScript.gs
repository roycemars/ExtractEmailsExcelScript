// Snippet which extracts emails from hyperlinks in Excel spreadsheet. Built in ExcelScript
// Demo video: https://youtu.be/TfYknQzNf6c

function main(workbook: ExcelScript.Workbook) {
    // Get the active cell and worksheet.
    let cell = workbook.getActiveCell();
    let selectedSheet = workbook.getActiveWorksheet();

    do {
        cell = workbook.getActiveCell();

        if (cell.getHyperlink()) {

            // Extract email from hyperlink
            let link = cell.getHyperlink().address
            let text = link.replace("mailto:", "")

            // Put email to the 2nd column
            cell.getOffsetRange(0, 1).setValue(text)
            
            // Mark processed cell light green
            cell.getFormat().getFill().setColor("#99ff99");
        }

        cell.getOffsetRange(1, 0).select()
    } while (cell.getText() != "")
}
