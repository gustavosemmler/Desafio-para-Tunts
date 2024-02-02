const { google } = require("googleapis");

async function authenticate() {
  try {
    const auth = await google.auth.getClient({
      keyFile: "credentials.json",
      scopes: "https://www.googleapis.com/auth/spreadsheets",
    });

    return auth;
  } catch (error) {
    console.error("Error during authentication:", error.message);
    throw error;
  }
}

async function readSpreadsheet(auth) {
  try {
    const sheets = google.sheets({ version: "v4", auth });
    const spreadsheetId = "12mjE9Sv1G9IrzJR3hNEMXzFXuJCpXbllVwmaNVPOGaU";
    const range = "engineering_software"; 

    const response = await sheets.spreadsheets.get({
      spreadsheetId,
    });

    const sheetNames = response.data.sheets.map(sheet => sheet.properties.title);
    console.log("Sheets found:", sheetNames);

    const sheet = response.data.sheets.find(sheet => sheet.properties.title === range);

    if (!sheet) {
      throw new Error(`Sheet "${range}" not found.`);
    }

    const values = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${range}!A1:H${sheet.properties.gridProperties.rowCount}`,
    });

    return values.data.values;
  } catch (error) {
    console.error("Error while reading spreadsheet:", error.message);
    throw error;
  }
}

async function calculateAndUpdateSheet(auth, data) {
  try {
    const sheets = google.sheets({ version: "v4", auth });
    const spreadsheetId = "12mjE9Sv1G9IrzJR3hNEMXzFXuJCpXbllVwmaNVPOGaU";
    const rangeToUpdate = "engineering_software!G4:K100"; 

    const valuesToWrite = [];

    // Iterate over the data and calculate the situation, the final approval grade, and the average
    for (const row of data.slice(3)) { // Start from row 4
      const [matricula, student, absences, p1, p2, p3] = row.map(Number);
      let average = ((p1 + p2 + p3) / 3).toFixed(2);
      let situation = "";
      let finalApprovalGrade = 0; 
      let result = "";

      if (absences > 0.25 * 60) {
        situation = "Failed due to Absences";
      } else if (average < 50) {
        situation = "Failed by Grade";
      } else if (average >= 50 && average < 70) {
        situation = "Final Exam";
        result = "Awaiting Exam Result";

        // Calculate the Final Approval Grade 
        finalApprovalGrade = Math.max(0, Math.ceil(100 - average));
      } else if (average >= 70) {
        situation = "Approved";
      }

      valuesToWrite.push([situation, finalApprovalGrade, result]);
    }

    // Results back to the spreadsheet
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: rangeToUpdate,
      valueInputOption: "USER_ENTERED",
      resource: { values: valuesToWrite },
    });
  } catch (error) {
    console.error("Error during calculation and update:", error.message);
    throw error;
  }
}

async function main() {
  try {
    const auth = await authenticate();
    const spreadsheetData = await readSpreadsheet(auth);

    await calculateAndUpdateSheet(auth, spreadsheetData);

    console.log("Processing completed.");
  } catch (error) {
    console.error("Error:", error.message);
  }
}

main();

/*
  How to Run This Node.js Application:

  1. Install Dependencies:
     Run the following command to install the required dependencies.
     ```
     npm install
     ```

  2. Configure the Environment:
     - If needed, configure environment variables or modify configuration files.

  3. Run the Application:
     Execute the following command to start the application.
     ```
     node index.js
     ```

     Alternatively, if your application uses a custom script defined in 'package.json':
     ```
     npm start
     ```

  Additional Notes:
  - Ensure you have Node.js and npm installed on your machine.
  - Check the 'package.json' file for any specific scripts or commands.

  If there are any specific dependencies or additional configurations required,
  please refer to the project documentation or contact the project maintainers.
*/