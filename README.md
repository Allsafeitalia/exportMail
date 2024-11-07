# Gmail Sent Email Extractor

Questo script di Google Apps Script consente di estrarre gli indirizzi email dai messaggi inviati nella tua casella di posta Gmail. Raccoglie automaticamente tutti gli indirizzi dei destinatari dalle email inviate, rimuove i duplicati e li visualizza in un elenco.

## Requisiti
- Un account Google
- Accesso a Google Drive

## Istruzioni per l'uso

### Passo 1: Creare lo Script
1. Apri **Google Drive**.
2. Crea un nuovo **Documento Google** o **Foglio Google**.
3. Vai su **Estensioni > Apps Script** per aprire l'editor di script.

### Passo 2: Inserire il Codice
Copia e incolla il seguente codice nell'editor di script:

```javascript
function getEmailsFromSent() {
  var threads = GmailApp.search("in:sent"); // Cerca tutte le email nella cartella "inviate"
  var addresses = []; // Array vuoto per memorizzare gli indirizzi email
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages(); // Ottiene tutti i messaggi in ogni thread (conversazione)
    
    for (var j = 0; j < messages.length; j++) {
      var toEmails = messages[j].getTo().split(","); // Prende gli indirizzi email dei destinatari
      addresses = addresses.concat(toEmails); // Aggiunge gli indirizzi all'array 'addresses'
    }
  }
  
  var uniqueAddresses = Array.from(new Set(addresses)); // Rimuove i duplicati creando un Set
  Logger.log(uniqueAddresses.join("\n")); // Stampa tutti gli indirizzi unici nel Logger
}
