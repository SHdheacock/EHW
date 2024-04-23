'use strict';
$("#setup").on("click", () => tryCatch(setup));
$("#move").on("click", () => tryCatch(move));

async function move() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("Sample");
        const sheet2 = context.workbook.worksheets.getItem("Sheet1");
        // Place a label in front of the moved data.
        // sheet2.getRange("A1").values = [["Moved Range:"]];

        // Move the range from A1:E1 to G12:K12.
        // sheet.getRange("A1:E1").moveTo("G12");
        // sheet.getRange("A1:E1").moveTo(sheet2.getRange("A1"));
        // sheet.getRange("A1").values = [[7]];
        sheet.getRange("A1").moveTo(sheet2.getRange("A1"));
        await context.sync();
    });
}

async function setup() {
    await Excel.run(async (context) => {
        context.workbook.worksheets.getItemOrNullObject("Sample").delete();

        const sheet = context.workbook.worksheets.add("Sample");
        sheet.getRange("A1:D1").values = [["3", "5", "7", ""]];
        sheet.getRange("A1:D1").format.font.italic = true;
        sheet.getRange("A1:D1").format.font.color = "DarkMagenta";
        sheet.getRange("E1").formulas = [["=SUM(A1:D1)"]];
        sheet.getRange("E1").format.font.bold = true;
        sheet.getRange("E1").format.fill.color = "LightGreen";
        sheet.getRange("F1").format.columnWidth = 120;

        sheet.activate();
        await context.sync();
    });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
} ();
