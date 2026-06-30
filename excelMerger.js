const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

async function mergeToMultipleSheets() {
    try {
        const folderPath = __dirname;
        const today = new Date().toISOString().split("T")[0];

        // ✅ Get files
        const files = fs.readdirSync(folderPath)
            .filter(file => file.startsWith(`All_IGA_FAST_${today}`) && file.endsWith(".xlsx"))
            .sort();

        if (files.length === 0) {
            console.log("❌ No files found");
            return;
        }

        console.log("📂 Files:", files);

        const finalFile = `All_IGA_FAST_${today}_COMBINED.xlsx`;

        // ✅ STREAM WRITER
        const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
            filename: finalFile,
            useStyles: false,
            useSharedStrings: false
        });

        let sheetCount = 1;

        // ✅ Process each file → separate sheet
        for (const file of files) {

            console.log(`🔄 Processing: ${file}`);

            const sheetName = `Sheet${sheetCount}`;
            const worksheet = workbook.addWorksheet(sheetName);

            const filePath = path.join(folderPath, file);
            const stream = fs.createReadStream(filePath);

            const reader = new ExcelJS.stream.xlsx.WorkbookReader(stream);

            for await (const worksheetReader of reader) {

                if (worksheetReader.name !== "All Data") continue;

                for await (const row of worksheetReader) {
                    worksheet.addRow(row.values).commit();
                }
            }

            worksheet.commit();
            sheetCount++;
        }

        await workbook.commit();

        console.log(`✅ FINAL FILE CREATED: ${finalFile}`);

    } catch (error) {
        console.error("❌ Merge failed:", error);
    }
}

mergeToMultipleSheets();