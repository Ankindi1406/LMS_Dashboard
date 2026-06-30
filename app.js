const express = require("express");
const axios = require("axios");
const ExcelJS = require("exceljs");

const app = express();
app.use(express.json());

// CONFIG
const TOKEN_URL = "https://successfactor-uat-3scale-apicast-production.apps.ocpnonprodcl01.goindigo.in/SuccessFactor-oAuth/oauth/token";
const API_URL = "https://cpi-test-preprod-3scale-apicast-production.apps.ocpnonprodcl01.goindigo.in/http/LMS_LearningHistory_PROD";
const USER_API = "https://cpi-test-preprod-3scale-apicast-production.apps.ocpnonprodcl01.goindigo.in/http/LMS_User_PROD";

const CLIENT_ID = "sb-55bcdade-3dd1-41cf-a3c6-e1fc329786bb!b37834|it-rt-indigo-cpi-test-in-wsgqfy65!b148";
const CLIENT_SECRET = "b46e9a9b-91d8-482d-8fa6-97dc04683f53$cIJcoF1il17cTey5FP9gL7wDkcx42aMVdvBJl0rmTuE= ";

// TOKEN CACHE
let cachedToken = null;
let tokenExpiryTime = 0;
let isFetchingToken = false;


// TOKEN FUNCTION 
async function getAccessToken() {
    const now = Date.now();

    if (cachedToken && now < tokenExpiryTime) {
        return cachedToken;
    }

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



// async function getStatusSummary(igaCode, token) {
//     try {
//         const payload = { IGACode: String(igaCode) };

//         const response = await axios({
//             method: "GET",
//             url: API_URL,
//             headers: {
//                 Authorization: `Bearer ${token}`,
//                 "Content-Type": "application/json",
//                 user_key: "ea3841f6a456660c8f069899a27c1803"
//             },
//             params: payload
//         });

//         const response2 = await axios({
//             method: "GET",
//             url: USER_API,
//             headers: {
//                 Authorization: `Bearer ${token}`,
//                 "Content-Type": "application/json",
//                 user_key: "ea3841f6a456660c8f069899a27c1803"
//             },
//             params: payload
//         });

//         const dataList = response.data.value || [];
//         const userList = response2.data.value || [];

//         const records = dataList.map(item => ({
//             IGACode: igaCode,
//             title: item.title || "",
//             grade: item.grade || "",
//             comments: item.comments || "",
//             status: (item.status || "unknown").toLowerCase(),
//             totalHours: item.totalHours || 0,
//             creditHours: item.creditHours || 0,
//             cpeHours: item.cpeHours || 0,
//             Designation:item.jobPosID||"",
//             Department:item.orgID||"",
//             Base:item.jobLocID||"",
//             completionDate: item.completionDate
//             ? new Date(
//             item.completionDate > 1e12
//             ? item.completionDate
//             : item.completionDate * 1000
//            )
//            .toLocaleString("sv-SE", { timeZone: "Asia/Kolkata" })
//            .replace("T", " ")
//            : "",

//             instructorName: item.instructorName || "",
//             componentId: item.componentID || "",

//             revisionNumber: item.revisionNumber || "",

//             Name:
//                 userList.length > 0
//                     ? userList[0].firstName + " " + userList[0].lastName
//                     : "Unknown"
//         }));

//         return records;

//     } catch (error) {
//         if (error.response && error.response.status === 401) {
//             const newToken = await getAccessToken();
//             return getStatusSummary(igaCode, newToken);
//         }

//         console.log(`❌ Failed IGACode ${igaCode}`);
//         return [];
//     }
// }



// async function generateExcelForAllIGA() {
//     const token = await getAccessToken();

//     const workbook = new ExcelJS.Workbook();
//     const worksheet = workbook.addWorksheet("All Data");

//        worksheet.columns = [
//         { header: "IGACode", key: "IGACode", width: 15 },
//         { header: "Name", key: "Name", width: 25 },
//         { header: "Title", key: "title", width: 30 },
//         { header: "Designation", key: "Designation", width: 20 },
//         { header: "Department", key: "Department", width: 20 },
//         { header: "Base", key: "Base", width: 20 },
//         { header: "Component ID", key: "componentId", width: 20 },
//         { header: "Grade", key: "grade", width: 10 },
//         { header: "Comments", key: "comments", width: 30 },
//         { header: "CompletionDate", key: "completionDate", width: 20 },
//         { header: "Instructor", key: "instructorName", width: 20 },
//         { header: "Status", key: "status", width: 15 },
//         { header: "Total Hours", key: "totalHours", width: 15 },
//         { header: "Credit Hours", key: "creditHours", width: 15 },
//         { header: "CPE Hours", key: "cpeHours", width: 15 },
//         { header: "Revision Number", key: "revisionNumber", width: 20 }
//     ];

//     console.log("🚀 Worker pool started...");

//     const total = 1000;
//     const concurrency = 25;  

//     let current = 1; 

//     // ✅ WORKER FUNCTION
//     async function worker() {
//         while (true) {
//             let iga;

            
//             if (current > total) break;
//             iga = current++;

//             try {
//                 console.log(`Processing IGACode: ${iga}`);

//                 const records = await getStatusSummary(iga, token);

//                 records.forEach(record => {
//                     worksheet.addRow(record);
//                 });

//             } catch (error) {
//                 console.log(`❌ Failed IGACode ${iga}`);
//             }
//         }
//     }

   
//     const workers = [];

//     for (let i = 0; i < concurrency; i++) {
//         workers.push(worker());
//     }

    
//     await Promise.all(workers);

//     const fileName = `All_IGA_FAST_${Date.now()}.xlsx`;
//     await workbook.xlsx.writeFile(fileName);

//     console.log(`✅ Excel saved: ${fileName}`);
// }


async function getStatusSummary(igaCode, token) {
    try {
        const payload = { IGACode: String(igaCode) };

        const [response, response2] = await Promise.all([
            axios({
                method: "GET",
                url: API_URL,
                headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": "application/json",
                    user_key: "ea3841f6a456660c8f069899a27c1803"
                },
                params: payload
            }),
            axios({
                method: "GET",
                url: USER_API,
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
                  )
                      .toLocaleString("sv-SE", { timeZone: "Asia/Kolkata" })
                      .replace("T", " ")
                : "",
            instructorName: item.instructorName || "",
            componentId: item.componentID || "",
            revisionNumber: item.revisionNumber || ""
        }));
    } catch (error) {
        if (error.response && error.response.status === 401) {
            const newToken = await getAccessToken();
            return getStatusSummary(igaCode, newToken);
        }

        console.log(`❌ Failed IGACode ${igaCode}`);
        return [];
    }
}

async function generateExcelForAllIGA() {
    const token = await getAccessToken();

    // const fileName = `All_IGA_FAST_${Date.now()}.xlsx`;


    const start = 1
    const end = 100;
    const total = end - start + 1;
    let current = start;
    const today = new Date().toISOString().split("T")[0];

    const fileName = `All_IGA_FAST_${today}_${start}_${end}.xlsx`;


   
    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: fileName,
        useStyles: true,
        useSharedStrings: true
    });

    const worksheet = workbook.addWorksheet("All Data");

    worksheet.columns = [
        { header: "IGACode", key: "IGACode", width: 15 },
        { header: "Name", key: "Name", width: 25 },
        { header: "Title", key: "title", width: 30 },
        { header: "Designation", key: "Designation", width: 20 },
        { header: "Department", key: "Department", width: 20 },
        { header: "Base", key: "Base", width: 20 },
        { header: "Component ID", key: "componentId", width: 20 },
        { header: "Grade", key: "grade", width: 10 },
        { header: "Comments", key: "comments", width: 30 },
        { header: "CompletionDate", key: "completionDate", width: 20 },
        { header: "Instructor", key: "instructorName", width: 20 },
        { header: "Status", key: "status", width: 15 },
        { header: "Total Hours", key: "totalHours", width: 15 },
        { header: "Credit Hours", key: "creditHours", width: 15 },
        { header: "CPE Hours", key: "cpeHours", width: 15 },
        { header: "Revision Number", key: "revisionNumber", width: 20 }
    ];

    console.log("🚀 Worker pool started...");

    //const total = 100;        // total IGAs
    const concurrency = 25;    // threads


    
    async function worker() {
    while (true) {

        let iga;
        
        
        if (current > end) return;
        iga = current;
        current++;

        try {
            console.log(`Processing IGACode: ${iga}`);

            const records = await getStatusSummary(iga, token);

            console.log(`Records for ${iga}:`, records.length)

            if (records.length === 0) continue; 

            records.forEach(record => {
                worksheet.addRow(record).commit();
            });

        } catch (error) {
            console.log(`❌ Failed IGACode ${iga}`);
        }

        //console.log(`Records for ${iga}:`, records.length);
    }
}


    // worker start
    const workers = [];
    for (let i = 0; i < concurrency; i++) {
        workers.push(worker());
    }

    await Promise.all(workers);

    worksheet.commit();

    // ✅ FINALIZE FILE
    await workbook.commit();

    console.log(`✅ Excel saved: ${fileName}`);
}





app.get("/generate-all", async (req, res) => {
    try {
        await generateExcelForAllIGA();
        res.json({ message: "Excel generated successfully" });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});



const PORT = 3000;
app.listen(PORT, () => {
    console.log(`✅ Server running on http://localhost:${PORT}`);
});