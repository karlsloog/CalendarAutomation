function saadaMeeldetuletused() {
  const spreadsheetId = "1E8sJmdER2nbuW6hNsgs2b_CKfphod5wyj99M9pYoNbQ";
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Sheet1");
  const data = sheet.getDataRange().getValues(); // ← Kasutame getValues()!

  const tana = new Date();
  tana.setHours(0, 0, 0, 0); // normalizeerime aja

  for (let i = 1; i < data.length; i++) {
    const rida = data[i];
    const kaamera = rida[0];
    const nimi = rida[1];
    const email = rida[12]; // M veerg
    const grupp = rida[2];
    const algus = new Date(rida[3]); // D veerg
    const tagastus = new Date(rida[5]); // F veerg
    const tagastatud = rida[7] === true;
    const meeldetuletatud = rida[11] === true; // L veerg (12. veerg)

    Logger.log(`➡️ Kontrollin rida ${i + 1}: email=${email}, tagastus=${tagastus}, algus=${algus}, tana=${tana}`);

    if (
      !email ||
      !kaamera ||
      kaamera.toString().toLowerCase().trim() === "muu" ||
      isNaN(tagastus.getTime()) ||
      tagastatud ||
      meeldetuletatud
    ) {
      Logger.log(`⏭️ Jäeti vahele rida ${i + 1}`);
      continue;
    }

    // Kontroll: kas tagastuskuupäev on täna
    const vaheMillis = tagastus.getTime() - tana.getTime();
    const onTagastusTana = vaheMillis === 0;

    // Kontroll: kas laenutus ja tagastus on samal päeval
    const samaPaev = algus.getTime() === tagastus.getTime();

    if (onTagastusTana && !samaPaev) {
      const subject = `📸 Meeldetuletus: kaamera ${kaamera} tagastus täna`;
      const message = `
        <html>
          <body>
            <p>Tere <strong>${nimi}</strong>,</p>
            <p>Tuletame meelde, et laenatud kaamera (${kaamera}) tuleb tagastada täna (${tagastus.toLocaleDateString()}).</p>
            <p>Palun veendu, et seade oleks õigeks ajaks tagastatud.</p>
            <p>Parimate soovidega,<br>Tartu Kunstikool</p>
          </body>
        </html>
      `;

      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: message
        });

        sheet.getRange(i + 1, 11).setValue(true); // märgi L veerg "true"
        Logger.log(`✅ Meeldetuletus saadetud: ${email}`);
      } catch (e) {
        Logger.log(`❌ Viga meili saatmisel real ${i + 1}: ${e}`);
      }
    }
  }
}