function lisaKalendrisseKontrolliga() {
  const spreadsheetId = "1E8sJmdER2nbuW6hNsgs2b_CKfphod5wyj99M9pYoNbQ";
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Sheet1");
  const data = sheet.getDataRange().getValues();
  const calendar = CalendarApp.getCalendarById("c_ea917da921bd1bb2b03c13511099dd10ebe9afd8fe4409ae4c8ee906918bb6a3@group.calendar.google.com"); // <- asenda oma päris kalendri ID

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const kaamera = row[0];
    const email = row[1];
    const grupp = row[2];
    const algus = new Date(row[3]);
    algus.setHours(0, 0, 0);
    const valjastaja = row[4];
    const tagastus = new Date(row[5]);
    tagastus.setHours(23, 59, 59);
    const vastuvotja = row[6];
    const tagastatud = row[7] === true;
    const lisaseadmed = row[8];
    const markused = row[9];
    const olemasolevEventID = row[10];
    const lisainfo = `Õppegrupp: ${grupp}\nLisaseadmed: ${lisaseadmed}\nMärkused: ${markused}\nVäljastaja: ${valjastaja}`;

    //  Välista tühjad, "muu", tagastatud või vigased read
    if (
      !kaamera || kaamera.toString().toLowerCase().trim() === "muu" ||
      !email || isNaN(algus.getTime()) || isNaN(tagastus.getTime()) ||
      tagastatud
    ) {
      continue;
    }

    if (!olemasolevEventID) {
      // Kui sündmus veel puudub — loo uus
      const event = calendar.createEvent(`Kaamera ${kaamera} - ${email}`, algus, tagastus, {
        description: lisainfo
      });
      sheet.getRange(i + 1, 11).setValue(event.getId());
    } else {
      // Kui sündmus olemas — uuenda kuupäevad ja kirjeldus
      try {
        const event = calendar.getEventById(olemasolevEventID);
        if (!event) {
          Logger.log(`Ei leidnud sündmust ID-ga: ${olemasolevEventID}`);
          continue;
        }

        const praeguneAlgus = event.getStartTime();
        const praeguneLopp = event.getEndTime();
        const praeguneKirjeldus = event.getDescription();

        const kuupaevMuutunud = praeguneAlgus.getTime() !== algus.getTime() || praeguneLopp.getTime() !== tagastus.getTime();
        const kirjeldusMuutunud = praeguneKirjeldus !== lisainfo;

        if (kuupaevMuutunud) {
          event.setTime(algus, tagastus);
        }

        if (kirjeldusMuutunud) {
          event.setDescription(lisainfo);
        }

        if (kuupaevMuutunud || kirjeldusMuutunud) {
          Logger.log(`🔄 Uuendatud sündmus: Kaamera ${kaamera}, rida ${i + 1}`);
        }

      } catch (e) {
        Logger.log(`⚠️ Viga sündmuse uuendamisel real ${i + 1}: ${e}`);
      }
    }
  }
}