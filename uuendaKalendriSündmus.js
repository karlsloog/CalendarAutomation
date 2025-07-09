function uuendaKalendriSundmus() {
  const spreadsheetId = "spreadsheetid";
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Sheet1");
  const data = sheet.getDataRange().getValues();
  const calendar = CalendarApp.getCalendarById("clanedar@group.calendar.google.com");

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const eventId = row[10];
    const kaamera = row[0];
    const email = row[1];
    const grupp = row[2];
    const algus = new Date(row[3]);
    const tahtaeg = new Date(row[5]);
    const tagastatud = row[7] === true;
    const lisainfo = `Oppegrupp: ${grupp}\nLisaseadmed: ${row[8]}\nMarkused: ${row[9]}`;

    if ((kaamera + "").toLowerCase() === "muu") continue;

    try {
      const event = calendar.getEventById(eventId);
      if (tagastatud) {
        event.deleteEvent();
        sheet.getRange(i + 1, 11).setValue(""); // Eemaldame Event ID
      } else {
        event.setTime(algus, tahtaeg);
        event.setDescription(lisainfo);
        event.setTitle(`Kaamera ${kaamera} - ${email}`);
      }
    } catch (e) {
      Logger.log(`❌ Viga rea ${i + 1} sundmusega: ${e.message}`);
    }
  }
}
