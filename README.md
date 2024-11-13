function getEmailsFromSent() {
  var ui = SpreadsheetApp.getUi(); // Per mostrare i dialoghi all'utente

  // Richiedi all'utente di inserire la data di inizio
  var responseStart = ui.prompt('Inserisci la data di inizio (YYYY/MM/DD) tra 2023/10/01 e 2023/12/31');
  if (responseStart.getSelectedButton() != ui.Button.OK) {
    ui.alert('Operazione annullata.');
    return;
  }
  var startDate = responseStart.getResponseText();

  // Richiedi all'utente di inserire la data di fine
  var responseEnd = ui.prompt('Inserisci la data di fine (YYYY/MM/DD) tra 2023/10/01 e 2023/12/31');
  if (responseEnd.getSelectedButton() != ui.Button.OK) {
    ui.alert('Operazione annullata.');
    return;
  }
  var endDate = responseEnd.getResponseText();

  // Valida le date inserite
  var start = new Date(startDate);
  var end = new Date(endDate);
  var minDate = new Date('2023/10/01');
  var maxDate = new Date('2023/12/31');

  if (isNaN(start) || isNaN(end)) {
    ui.alert('Le date inserite non sono valide.');
    return;
  }
  if (start < minDate || start > maxDate || end < minDate || end > maxDate) {
    ui.alert('Le date devono essere comprese tra 2023/10/01 e 2023/12/31.');
    return;
  }
  if (start > end) {
    ui.alert('La data di inizio deve essere precedente alla data di fine.');
    return;
  }

  // Formatta le date per la query di ricerca
  var queryStartDate = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  var queryEndDate = Utilities.formatDate(end, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  // Crea la query di ricerca
  var query = 'in:sent after:' + queryStartDate + ' before:' + queryEndDate;

  var threads = GmailApp.search(query); // Cerca le email nell'intervallo di date specificato
  var addresses = []; // Array per memorizzare nomi, email e oggetti delle email

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();

    for (var j = 0; j < messages.length; j++) {
      var toEmails = messages[j].getTo().split(",");
      var subject = messages[j].getSubject();

      for (var k = 0; k < toEmails.length; k++) {
        var recipient = toEmails[k].trim();
        var emailRegex = /([^<]+)?<(.+)>/;
        var match = recipient.match(emailRegex);
        var name, email;

        if (match) {
          name = match[1] ? match[1].trim() : '';
          email = match[2] ? match[2].trim() : recipient.trim();
        } else {
          name = '';
          email = recipient;
        }

        // Salva i dati nell'array
        addresses.push({
          name: name,
          email: email,
          subject: subject
        });
      }
    }
  }

  // Rimuove i duplicati basati su email e oggetto
  var uniqueAddresses = [];
  var uniqueKeys = new Set();

  for (var i = 0; i < addresses.length; i++) {
    var key = addresses[i].email + addresses[i].subject;
    if (!uniqueKeys.has(key)) {
      uniqueAddresses.push(addresses[i]);
      uniqueKeys.add(key);
    }
  }

  // Scrive i risultati sul foglio di Google
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clearContents(); // Pulisce il foglio corrente

  // Scrive l'intestazione
  sheet.appendRow(['Nome', 'Email', 'Oggetto']);

  // Scrive i dati
  for (var i = 0; i < uniqueAddresses.length; i++) {
    sheet.appendRow([uniqueAddresses[i].name, uniqueAddresses[i].email, uniqueAddresses[i].subject]);
  }

  ui.alert('I dati sono stati scritti sul foglio di Google.');
}

