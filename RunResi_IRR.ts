
function main(workbook: ExcelScript.Workbook) {
    
    const table = workbook.getTable("IRR_Resi_Table");
    const assumptions = workbook.getWorksheet("Assumptions");
    assumptions.getRange("F7:F12").clear();
    assumptions.getRange("F21:F27").clear();
    assumptions.getRange("F7").setValue(6);

    // Resize table to include all used rows on the table's worksheet
    const tableRange = table.getRange();
    const sheet = tableRange.getWorksheet();
    const usedRange = sheet.getUsedRange();

    const tableStartRow = tableRange.getRowIndex();
    const tableStartCol = tableRange.getColumnIndex();
    const tableColCount = tableRange.getColumnCount();

    const usedLastRow = usedRange.getRowIndex() + usedRange.getRowCount() - 1;
    const newRowCount = usedLastRow - tableStartRow + 1;

    const resizedRange = sheet.getRangeByIndexes(
        tableStartRow,
        tableStartCol,
        newRowCount,
        tableColCount
    );

    table.resize(resizedRange);

    const data = table.getRangeBetweenHeaderAndTotal().getValues();
    const irrCol = table.getColumnByName("IRR (Current LEM)").getRangeBetweenHeaderAndTotal();
    const out: (number | string)[][] = [];

    const toNum = (v: unknown): number | null => {
        if (v === null || v === undefined || v === "") return null;
        if (typeof v === "number") return v;
        const n = Number(String(v).replace(/[$,\s]/g, ""));
        return Number.isFinite(n) ? n : null;
    };

    for (let i = 0; i < data.length; i++) {
        
        const estimate = toNum(data[i][15]);  // Column P
        const passings = toNum(data[i][8]);  // Column I

        if (estimate === null || passings === null || passings <= 0) {
            out.push(["Unavailable"]);
            continue;
        }

        assumptions.getRange("F8").setValue(passings);
        assumptions.getRange("F10").setValue(estimate);

        workbook.getApplication().calculate(ExcelScript.CalculationType.recalculate);

        const irr = assumptions.getRange("F15").getValue();

        out.push(typeof irr === "number" && Number.isFinite(irr) ? [irr] : ["Unavailable"]);
    }

    irrCol.setValues(out);
    irrCol.setNumberFormat("0.0%");
    workbook.getApplication().calculate(ExcelScript.CalculationType.recalculate);
}
