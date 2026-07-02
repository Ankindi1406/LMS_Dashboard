// const ExcelJS = require("exceljs");
// const fs = require("fs");
// const path = require("path");

// // ✅ SharePoint synced folder
// const OUTPUT_DIR = "C:\\Users\\Ankush.Kumar\\OneDrive - InterGlobe Aviation Limited\\test - Testing\\lms_uploads"

// async function mergeToMultipleSheets() {
//     try {

//         const folderPath = process.cwd();  // ✅ FIXED
//         const today = new Date().toISOString().split("T")[0];

//         // ✅ Get batch files
//         const files = fs.readdirSync(folderPath)
//             .filter(file => file.startsWith(`All_IGA_FAST_${today}`) && file.endsWith(".xlsx"))
//             .sort();

//         if (files.length === 0) {
//             console.log("❌ No files found");
//             return;
//         }

//         console.log("📂 Files:", files);

//         // ✅ Save final file directly to SharePoint folder
//         const finalFile = path.join(
//             OUTPUT_DIR,
//             `All_IGA_FINAL_${today}.xlsx`
//         );

//         const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
//             filename: finalFile,
//             useStyles: false,
//             useSharedStrings: false
//         });

//         let sheetCount = 1;

//         for (const file of files) {

//             console.log(`🔄 Processing: ${file}`);

//             const worksheet = workbook.addWorksheet(`Batch_${sheetCount}`);

//             const filePath = path.join(folderPath, file);
//             const stream = fs.createReadStream(filePath);

//             const reader = new ExcelJS.stream.xlsx.WorkbookReader(stream);

//             for await (const worksheetReader of reader) {

//                 if (worksheetReader.name !== "All Data") continue;

//                 for await (const row of worksheetReader) {
//                     worksheet.addRow(row.values).commit();
//                 }
//             }

//             worksheet.commit();
//             sheetCount++;
//         }

//         await workbook.commit();

//         console.log(`✅ FINAL FILE CREATED: ${finalFile}`);

//         // ✅ CLEANUP batch files
//         console.log("🧹 Cleaning batch files...");

//         files.forEach(file => {
//             const fullPath = path.join(folderPath, file);
//             if (fs.existsSync(fullPath)) {
//                 fs.unlinkSync(fullPath);
//             }
//         });

//         console.log("✅ Batch files deleted");

//     } catch (error) {
//         console.error("❌ Merge failed:", error);
//     }
// }

// mergeToMultipleSheets();




const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

// ✅ SharePoint synced folder
const OUTPUT_DIR =
    "C:\\Users\\Ankush.Kumar\\OneDrive - InterGlobe Aviation Limited\\test - Testing\\lms_uploads";

async function mergeToMultipleSheets() {
    try {

        const folderPath = process.cwd();
        const today = new Date().toISOString().split("T")[0];

        // ✅ Get batch files from runner folder
        const files = fs.readdirSync(folderPath)
            .filter(file =>
                file.startsWith(`All_IGA_FAST_${today}`) &&
                file.endsWith(".xlsx")
            )
            .sort();

        if (files.length === 0) {
            console.log("❌ No files found");
            return;
        }

        console.log("📂 Files found:", files.length);

        // ✅ Create merged file in runner folder FIRST
        const finalFile = path.join(
            folderPath,
            `All_IGA_FINAL_${today}.xlsx`
        );

        const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
            filename: finalFile,
            useStyles: false,
            useSharedStrings: false
        });

        let sheetCount = 1;

        for (const file of files) {

            console.log(`🔄 Processing: ${file}`);

            const worksheet = workbook.addWorksheet(
                `Batch_${sheetCount}`
            );

            const filePath = path.join(folderPath, file);

            const reader =
                new ExcelJS.stream.xlsx.WorkbookReader(
                    fs.createReadStream(filePath)
                );

            for await (const worksheetReader of reader) {

                if (worksheetReader.name !== "All Data") {
                    continue;
                }

                for await (const row of worksheetReader) {
                    worksheet.addRow(row.values).commit();
                }
            }

            worksheet.commit();
            sheetCount++;
        }

        // ✅ Finish writing file
        await workbook.commit();

        console.log(`✅ Final merged file created: ${finalFile}`);

        // ✅ Copy to SharePoint sync folder AFTER creation
        const sharepointFile = path.join(
            OUTPUT_DIR,
            `All_IGA_FINAL_${today}.xlsx`
        );

        console.log("📤 Copying to SharePoint folder...");

        if (fs.existsSync(sharepointFile)) {
            fs.unlinkSync(sharepointFile);
        }

        fs.copyFileSync(finalFile, sharepointFile);

        console.log(`✅ Copied to SharePoint folder`);
        console.log(`📁 ${sharepointFile}`);

        // ✅ Delete batch files
        console.log("🧹 Cleaning batch files...");

        files.forEach(file => {
            const fullPath = path.join(folderPath, file);

            if (fs.existsSync(fullPath)) {
                fs.unlinkSync(fullPath);
            }
        });

        console.log("✅ Batch files deleted");

        // ✅ Delete local merged file
        if (fs.existsSync(finalFile)) {
            fs.unlinkSync(finalFile);
            console.log("✅ Local merged file deleted");
        }

        console.log("🎉 Merge process completed successfully");

    } catch (error) {
        console.error("❌ Merge failed:", error);
    }
}

mergeToMultipleSheets();
