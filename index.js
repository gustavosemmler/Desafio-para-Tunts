const { google } = require("googleapis");

async function authenticate() {
  const auth = await google.auth.getClient({
    keyFile: "credentials.json",
    scopes: "https://www.googleapis.com/auth/spreadsheets",
  });

  return auth;
}

async function readSpreadsheet(auth) {
  const sheets = google.sheets({ version: "v4", auth });
  const spreadsheetId = "12mjE9Sv1G9IrzJR3hNEMXzFXuJCpXbllVwmaNVPOGaU";
  const range = "engenharia_de_software"; // Ajuste conforme necessário

  const response = await sheets.spreadsheets.get({
    spreadsheetId,
  });

  const sheetNames = response.data.sheets.map(sheet => sheet.properties.title);
  console.log("Planilhas encontradas:", sheetNames);

  const sheet = response.data.sheets.find(sheet => sheet.properties.title === range);

  if (!sheet) {
    throw new Error(`Planilha "${range}" não encontrada.`);
  }

  const rangeToUpdate = `${range}!G4:K${sheet.properties.gridProperties.rowCount + 3}`; // Ajuste conforme necessário

  const values = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${range}!A1:H${sheet.properties.gridProperties.rowCount}`, // Ajuste conforme necessário
  });

  return values.data.values;
}

async function calculateAndUpdateSheet(auth, data) {
  const sheets = google.sheets({ version: "v4", auth });
  const spreadsheetId = "12mjE9Sv1G9IrzJR3hNEMXzFXuJCpXbllVwmaNVPOGaU";
  const rangeToUpdate = "engenharia_de_software!G4:K100"; // Ajuste conforme necessário

  // Adicione o cabeçalho para os resultados
  const valuesToWrite = [];

  // Iterar sobre os dados e calcular a situação, a nota para aprovação final e a média
  for (const row of data.slice(3)) { // Começar da linha 4
    const [matricula, aluno, faltas, p1, p2, p3] = row.map(Number);
    const media = ((p1 + p2 + p3) / 3).toFixed(2);
    let situacao = "";
    let resultado = "";

    if (faltas > 0.25 * 60) {
      situacao = "Reprovado por Falta";
      resultado = "";
    } else if (media < 50) {
      situacao = "Reprovado por Nota";
      resultado = "";
    } else if (media >= 50 && media < 70) {
      situacao = "Exame Final";
      resultado = "Aguardando resultado do Exame";
    } else if (media >= 70) {
      situacao = "Aprovado";
      resultado = "";
    }

    valuesToWrite.push([situacao, media, resultado]);
  }

  // Escrever os resultados de volta na planilha
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: rangeToUpdate,
    valueInputOption: "USER_ENTERED",
    resource: { values: valuesToWrite },
  });
}

async function main() {
  try {
    const auth = await authenticate();
    const spreadsheetData = await readSpreadsheet(auth);

    await calculateAndUpdateSheet(auth, spreadsheetData);

    console.log("Processamento concluído.");
  } catch (error) {
    console.error("Erro:", error.message);
  }
}

main();
