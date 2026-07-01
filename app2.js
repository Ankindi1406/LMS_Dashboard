const axios = require("axios");
const ExcelJS = require("exceljs");

const OUTPUT_DIR = "C:\\Users\\Ankush.Kumar\\OneDrive - InterGlobe Aviation Limited\\test - Testing\\lms_uploads"

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}


// ✅ CONFIG (move secrets to env later)
const TOKEN_URL = "https://successfactor-uat-3scale-apicast-production.apps.ocpnonprodcl01.goindigo.in/SuccessFactor-oAuth/oauth/token";
const API_URL = "https://cpi-test-preprod-3scale-apicast-production.apps.ocpnonprodcl01.goindigo.in/http/LMS_LearningHistory_PROD";
const USER_API = "https://cpi-test-preprod-3scale-apicast-production.apps.ocpnonprodcl01.goindigo.in/http/LMS_User_PROD";

const CLIENT_ID = "sb-55bcdade-3dd1-41cf-a3c6-e1fc329786bb!b37834|it-rt-indigo-cpi-test-in-wsgqfy65!b148";
const CLIENT_SECRET = "b46e9a9b-91d8-482d-8fa6-97dc04683f53$cIJcoF1il17cTey5FP9gL7wDkcx42aMVdvBJl0rmTuE= ";

// ✅ TOKEN CACHE
let cachedToken = null;
let tokenExpiryTime = 0;
let isFetchingToken = false;

// ✅ GET TOKEN
async function getAccessToken() {
    const now = Date.now();

    if (cachedToken && now < tokenExpiryTime) return cachedToken;

    if (isFetchingToken) {
        while (isFetchingToken) {
            await new Promise(res => setTimeout(res, 100));
        }
        return cachedToken;
    }

    try {
        isFetchingToken = true;

        const params = new URLSearchParams();
        params.append("grant_type", "client_credentials");
        params.append("client_id", CLIENT_ID.trim());
        params.append("client_secret", CLIENT_SECRET.trim());

        const response = await axios.post(TOKEN_URL, params, {
            headers: {
                "Content-Type": "application/x-www-form-urlencoded",
                user_key: "1f36e20be56ff2eee62d41fac9a6b63b"
            }
        });

        cachedToken = response.data.access_token;

        const expiresIn = response.data.expires_in || 3600;
        tokenExpiryTime = Date.now() + (expiresIn - 60) * 1000;

        return cachedToken;

    } finally {
        isFetchingToken = false;
    }
}

// ✅ FETCH DATA
async function getStatusSummary(igaCode, token) {
    try {
        const payload = { IGACode: String(igaCode) };

        const [response, response2] = await Promise.all([
            axios.get(API_URL, {
                headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": "application/json",
                    user_key: "ea3841f6a456660c8f069899a27c1803"
                },
                params: payload
            }),
            axios.get(USER_API, {
                headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": "application/json",
                    user_key: "ea3841f6a456660c8f069899a27c1803"
                },
                params: payload
            })
        ]);

        const dataList = response.data.value || [];
        const userList = response2.data.value || [];

        const name =
            userList?.[0]
                ? `${userList[0].firstName} ${userList[0].lastName}`
                : "Unknown";

        return dataList.map(item => ({
            IGACode: igaCode,
            Name: name,
            title: item.title || "",
            grade: item.grade || "",
            comments: item.comments || "",
            status: (item.status || "unknown").toLowerCase(),
            totalHours: item.totalHours || 0,
            creditHours: item.creditHours || 0,
            cpeHours: item.cpeHours || 0,
            Designation: userList[0]?.jobPosID || "",
            Department: userList[0]?.orgID || "",
            Base: userList[0]?.jobLocID || "",
            completionDate: item.completionDate
                ? new Date(
                    item.completionDate > 1e12
                        ? item.completionDate
                        : item.completionDate * 1000
                ).toLocaleString("sv-SE", { timeZone: "Asia/Kolkata" }).replace("T", " ")
                : "",
            instructorName: item.instructorName || "",
            componentId: item.componentID || "",
            revisionNumber: item.revisionNumber || ""
        }));

    } catch (error) {
        if (error.response?.status === 401) {
            const newToken = await getAccessToken();
            return getStatusSummary(igaCode, newToken);
        }

        console.log(`❌ Failed IGACode ${igaCode}`);
        return [];
    }
}

// ✅ MAIN FUNCTION
// async function generateExcel() {
//     console.log("🚀 Job started");

//     const token = await getAccessToken();

//     const start = 1;
//     const end = 100;  // change dynamically later
//     let current = start;

//     const today = new Date().toISOString().split("T")[0];
//     const fileName = `All_IGA_FAST_${today}_${start}_${end}.xlsx`;

//     // ✅ STREAM WRITER (optimized)
//     const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
//         filename: fileName,
//         useStyles: false,
//         useSharedStrings: false
//     });

//     const worksheet = workbook.addWorksheet("All Data");

//     worksheet.columns = [
//         { header: "IGACode", key: "IGACode" },
//         { header: "Name", key: "Name" },
//         { header: "Title", key: "title" },
//         { header: "Designation", key: "Designation" },
//         { header: "Department", key: "Department" },
//         { header: "Base", key: "Base" },
//         { header: "Component ID", key: "componentId" },
//         { header: "Grade", key: "grade" },
//         { header: "Comments", key: "comments" },
//         { header: "CompletionDate", key: "completionDate" },
//         { header: "Instructor", key: "instructorName" },
//         { header: "Status", key: "status" },
//         { header: "Total Hours", key: "totalHours" },
//         { header: "Credit Hours", key: "creditHours" },
//         { header: "CPE Hours", key: "cpeHours" },
//         { header: "Revision Number", key: "revisionNumber" }
//     ];

//     const concurrency = 20;

//     async function worker() {
//         while (true) {

//             if (current > end) return;

//             const iga = current++;

//             try {
//                 console.log(`🔄 Processing IGACode: ${iga}`);

//                 const records = await getStatusSummary(iga, token);

//                 if (records.length === 0) continue;

//                 records.forEach(r => {
//                     worksheet.addRow(r).commit();
//                 });

//             } catch (err) {
//                 console.log(`❌ Error IGACode ${iga}`);
//             }
//         }
//     }

//     const workers = Array.from({ length: concurrency }, worker);

//     await Promise.all(workers);

//     worksheet.commit();
//     await workbook.commit();

//     console.log(`✅ File created: ${fileName}`);
// }



async function generateBatch(start, end) {
    const startTime = new Date();

    console.log(`\n🚀 STARTING BATCH: ${start} → ${end}`);
    console.log(`⏰ Start Time: ${startTime.toLocaleString()}`);

    const token = await getAccessToken();

    const today = new Date().toISOString().split("T")[0];
     const fileName = `All_IGA_FAST_${today}_${start}_${end}.xlsx`;


    // const fileName = path.join(
    // OUTPUT_DIR,
    // `All_IGA_FAST_${today}_${start}_${end}.xlsx`
    // );

    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: fileName,
        useStyles: false,
        useSharedStrings: false
    });

    const worksheet = workbook.addWorksheet("All Data");

    worksheet.columns = [
        { header: "IGACode", key: "IGACode" },
        { header: "Name", key: "Name" },
        { header: "Title", key: "title" },
        { header: "Designation", key: "Designation" },
        { header: "Department", key: "Department" },
        { header: "Base", key: "Base" },
        { header: "Component ID", key: "componentId" },
        { header: "Grade", key: "grade" },
        { header: "Comments", key: "comments" },
        { header: "CompletionDate", key: "completionDate" },
        { header: "Instructor", key: "instructorName" },
        { header: "Status", key: "status" },
        { header: "Total Hours", key: "totalHours" },
        { header: "Credit Hours", key: "creditHours" },
        { header: "CPE Hours", key: "cpeHours" },
        { header: "Revision Number", key: "revisionNumber" }
    ];

    let current = start;
    const concurrency = 20;
    let processedCount = 0;

    async function worker() {
        while (true) {

            if (current > end) break;

            const iga = current++;

            try {
                const records = await getStatusSummary(iga, token);

                records.forEach(r => {
                    worksheet.addRow(r).commit();
                });

                processedCount++;

                // ✅ OPTIONAL progress log every 1000
                if (processedCount % 1000 === 0) {
                    console.log(`📊 Batch ${start}-${end}: Processed ${processedCount} IGAs`);
                }

            } catch (err) {
                console.log(`❌ Failed IGACode ${iga}`);
            }
        }
    }

    await Promise.all(Array.from({ length: concurrency }, worker));

    worksheet.commit();
    await workbook.commit();

    const endTime = new Date();

    console.log(`✅ BATCH COMPLETED: ${start} → ${end}`);
    console.log(`📊 Total IGAs Processed: ${processedCount}`);
    console.log(`⏰ End Time: ${endTime.toLocaleString()}`);
    console.log(`⏱ Duration: ${(endTime - startTime) / 1000}s`);
    console.log(`📁 File Created: ${fileName}\n`);

    return fileName;
}





// async function runAllBatches() {

//     const BATCH_SIZE = 100;
//     const MAX_IGA = 1000;
//     const DELAY = 0.5 * 60 * 1000; // 5 minutes

//     console.log("🚀 FULL JOB STARTED\n");

//     for (let start = 1; start <= MAX_IGA; start += BATCH_SIZE) {

//         const end = Math.min(start + BATCH_SIZE - 1, MAX_IGA);

//         const fileName = await generateBatch(start, end);
//         filescreated.push(fileName);

//         // ✅ Logging after each batch
//         console.log(`✅ Finished batch ${start}-${end}`);

//         if (end < MAX_IGA) {
//             console.log(`⏳ Waiting 5 minutes before next batch...\n`);
//             await sleep(DELAY);
//         }
//     }

//     console.log("🎯 ALL BATCHES COMPLETED");

//     // ✅ FINAL MERGE
//     console.log("🔄 Starting merge...");
//    // await mergeFiles();

//     console.log("✅ FINAL MERGE COMPLETED");
// }





async function runAllBatches() {

    const BATCH_SIZE = 100;
    const MAX_IGA = 1000;
    const DELAY = 0.5 * 60 * 1000;

    console.log("🚀 FULL JOB STARTED\n");

    for (let start = 1; start <= MAX_IGA; start += BATCH_SIZE) {

        const end = Math.min(start + BATCH_SIZE - 1, MAX_IGA);

        await generateBatch(start, end);

        console.log(`✅ Finished batch ${start}-${end}`);

        if (end < MAX_IGA) {
            console.log(`⏳ Waiting before next batch...\n`);
            await sleep(DELAY);
        }
    }

    console.log("🎯 ALL BATCHES COMPLETED");
}




















// ✅ AUTO RUN (KEY)
(async () => {
    try {
        await runAllBatches();
        console.log("✅ Job completed");
        process.exit(0);
    } catch (err) {
        console.error("❌ Job failed:", err);
        process.exit(1);
    }
})();