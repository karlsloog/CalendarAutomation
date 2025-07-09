function kustutaTagastatudSündmused() {
  const spreadsheetId = "spreadsheetId";
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Sheet1");
  const data = sheet.getDataRange().getValues();
  const calendar = CalendarApp.getCalendarById("c_ea917da921bd1bb2b03c13511099dd10ebe9afd8fe4409ae4c8ee906918bb6a3@group.calendar.google.com");

  for (let i = 1; i < data.length; i++) {
    const tagastatud = data[i][7] === true;
    const eventId = data[i][10];
    if (tagastatud && eventId) {
      try {
        const event = calendar.getEventById(eventId);
        if (event) event.deleteEvent();
      } catch (e) {}
      sheet.getRange(i + 1, 11).clearContent();
    }
  }
}
