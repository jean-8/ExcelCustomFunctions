async function isMergedCustomFunction(cellAddress) {
    return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const cell = sheet.getRange(cellAddress);
        const mergedArea = cell.getMergedAreas();
        mergedArea.load(["rowCount", "columnCount"]);
        await context.sync();

        return mergedArea.rowCount > 1 || mergedArea.columnCount > 1;
    });
}

async function mergedValueCustomFunction(cellAddress) {
    return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const cell = sheet.getRange(cellAddress);
        const mergedArea = cell.getMergedAreas();
        mergedArea.load("values");
        await context.sync();

        return mergedArea.values[0][0];
    });
}

export { isMergedCustomFunction, mergedValueCustomFunction };
