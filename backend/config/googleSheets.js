import { google } from "googleapis";
import dotenv from "dotenv";

dotenv.config();

const auth = new google.auth.GoogleAuth({
    keyFile: "./config/credenciais-google.json",
    scopes: [
        "https://www.googleapis.com/auth/spreadsheets"
    ]
});

const sheets = google.sheets({
    version: "v4",
    auth: auth
});

const spreadsheetId = process.env.SPREADSHEET_ID;

export { sheets, spreadsheetId };