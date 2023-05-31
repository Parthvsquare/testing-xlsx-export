"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
// import * as fs from 'fs';
var XLSX = require("xlsx");
// Data to be written to the Excel file
var data = {
    reportId: "0a24fd75-ff39-4f74-a936-dc2c28edc293",
    userId: "user_2QI3AI2rTDt3bbCRUUTVGG5shrw",
    generatedFor: "2023-05-29T18:31:45.940Z",
    savedAlerts: [],
    savedGrades: [],
    savedLiters: [
        {
            liter: 155.252,
            literId: "3cf99ae0-36ec-41ed-94c8-997242302ea4",
            createdAt: "2023-05-29T21:50:11.694Z",
            userId: "user_2QI3AI2rTDt3bbCRUUTVGG5shrw",
            reportDatabseReportId: "0a24fd75-ff39-4f74-a936-dc2c28edc293",
        },
        {
            liter: 154.252,
            literId: "86dd0604-79fe-404f-824d-0c01b35a51a8",
            createdAt: "2023-05-29T21:50:07.057Z",
            userId: "user_2QI3AI2rTDt3bbCRUUTVGG5shrw",
            reportDatabseReportId: "0a24fd75-ff39-4f74-a936-dc2c28edc293",
        },
        {
            liter: 148.252,
            literId: "bcce5ef6-0d15-4c2d-a375-a9a0db673df1",
            createdAt: "2023-05-29T21:50:05.410Z",
            userId: "user_2QI3AI2rTDt3bbCRUUTVGG5shrw",
            reportDatabseReportId: "0a24fd75-ff39-4f74-a936-dc2c28edc293",
        },
        {
            liter: 149.252,
            literId: "fdfb5247-e6e0-44f1-82ad-5efd164047d5",
            createdAt: "2023-05-29T21:50:09.152Z",
            userId: "user_2QI3AI2rTDt3bbCRUUTVGG5shrw",
            reportDatabseReportId: "0a24fd75-ff39-4f74-a936-dc2c28edc293",
        },
    ],
    savedWeights: [
        {
            weight: 167.002,
            weightId: "29996b88-9cdf-42e1-81b3-ceb53f432d0f",
            createdAt: "2023-05-29T21:49:42.029Z",
            userId: "user_2QI3AI2rTDt3bbCRUUTVGG5shrw",
            reportDatabseReportId: "0a24fd75-ff39-4f74-a936-dc2c28edc293",
        },
        {
            weight: 176.002,
            weightId: "5756ba9d-f1fc-4594-8406-1c3bcc2078ca",
            createdAt: "2023-05-29T21:49:45.360Z",
            userId: "user_2QI3AI2rTDt3bbCRUUTVGG5shrw",
            reportDatabseReportId: "0a24fd75-ff39-4f74-a936-dc2c28edc293",
        },
        {
            weight: 171.002,
            weightId: "e81bfa82-b771-4d28-87be-fc4e50cce761",
            createdAt: "2023-05-29T21:49:49.958Z",
            userId: "user_2QI3AI2rTDt3bbCRUUTVGG5shrw",
            reportDatabseReportId: "0a24fd75-ff39-4f74-a936-dc2c28edc293",
        },
    ],
    savedPieceCounting: [],
};
// Create a new workbook
var workbook = XLSX.utils.book_new();
// Iterate over the data and create a sheet for each table name
for (var tableName in data) {
    if (Array.isArray(data[tableName])) {
        var sheetName = tableName;
        var sheetData = data[tableName];
        // Ensure sheetData is defined and not empty
        if (sheetData && sheetData.length > 0) {
            // Create a new sheet
            var worksheet = XLSX.utils.json_to_sheet(sheetData);
            // Add the field headers at the top of the sheet
            var headers = Object.keys(sheetData[0]);
            XLSX.utils.sheet_add_aoa(worksheet, [headers], { origin: "A1" });
            // Convert createdAt values to localized string format
            for (var i = 0; i < sheetData.length; i++) {
                var createdAt = sheetData[i].createdAt;
                if (createdAt) {
                    sheetData[i].createdAt = new Date(createdAt).toLocaleString();
                }
            }
            // Add the modified sheet data to the sheet
            XLSX.utils.sheet_add_json(worksheet, sheetData, { skipHeader: true, origin: "A2" });
            // Add the sheet to the workbook
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        }
    }
}
// Write the workbook to a file
var excelFilePath = "output.xlsx";
XLSX.writeFile(workbook, excelFilePath);
console.log("Excel file created at: ".concat(excelFilePath));
