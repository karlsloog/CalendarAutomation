function fillEmailsFromNames() {
  const spreadsheetId = "1E8sJmdER2nbuW6hNsgs2b_CKfphod5wyj99M9pYoNbQ";
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Sheet1");
  const domain = 'tartukunstikool.ee'; // <<< Asenda oma domeeniga!
  
  const names = sheet.getRange("B2:B").getValues(); // B veerg, alates 2. reast
  
  for (let i = 0; i < names.length; i++) {
    const fullName = names[i][0];
    if (!fullName) continue;
    
    try {
      const response = AdminDirectory.Users.list({
        domain: domain,
        query: `name:"${fullName}"`,
        maxResults: 1
      });

      if (response.users && response.users.length > 0) {
        const email = response.users[0].primaryEmail;
        sheet.getRange(i + 2, 13).setValue(email); // veerg M (13), samale reale
      } else {
        sheet.getRange(i + 2, 13).setValue("Ei leitud");
      }
    } catch (error) {
      Logger.log(`Viga kasutaja "${fullName}" puhul: ${error}`);
      sheet.getRange(i + 2, 13).setValue("Viga");
    }
  }
}

// nime eemaldamisel emaili kustutamine
function onEdit(e) {
  const range = e.range;
  const sheet = e.source.getSheetByName("Sheet1");

  if (sheet.getName() !== "Sheet1") return; // Ainult Sheet1
  if (range.getColumn() !== 2) return; // Ainult veerg B

  const row = range.getRow();
  const newValue = range.getValue();

  if (!newValue) {
    // Kui nimi kustutati, kustuta samalt realt M veeru väärtus
    sheet.getRange(row, 13).clearContent(); // M = 13
  }
}