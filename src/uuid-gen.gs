function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📦 Status")
    .addItem("📜 Gerar UUID", "atualizaUUID")
    .addToUi();
}

function atualizaUUID() {
  const ABA = "Teste UUID";
  const LINHA_INICIAL = 2;
  const COL_A = 1;    // Coluna A
  const COL_UUID = 2;  // Coluna B

  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA);
  if (!sheet) return;

  const ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < LINHA_INICIAL) return;

  // Lê dados
  const colA = sheet
    .getRange(LINHA_INICIAL, COL_A, ultimaLinha - LINHA_INICIAL + 1, 1)
    .getValues();

  const colUUID = sheet
    .getRange(LINHA_INICIAL, COL_UUID, ultimaLinha - LINHA_INICIAL + 1, 1)
    .getValues();

  // UUIDs existentes (para garantir unicidade)
  const UUIDsExistentes = colUUID.flat().filter(String);

  for (let i = 0; i < colA.length; i++) {
    if (colA[i][0] !== "" && colUUID[i][0] === "") {

      let UUID;
      do {
        UUID = Utilities.getUuid()
          .replace(/-/g, "")
          .substring(0, 8)
          .toUpperCase();
      } while (UUIDsExistentes.includes(UUID));

      colUUID[i][0] = UUID;
      UUIDsExistentes.push(UUID); // evita repetir no mesmo loop
    }
  }

  // Escreve tudo de uma vez (rápido)
  sheet
    .getRange(LINHA_INICIAL, COL_UUID, colUUID.length, 1)
    .setValues(colUUID);
}
