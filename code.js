// Documentation Google Apps Script
//https://developers.google.com/apps-script/reference/document/body

// Counts placeholders {{xxx}}, £IF, £FIN, £OK, £KO in all tabs
function countPlaceholders(docId) {

  const docFile = DocumentApp.openById(docId);

  let tabs = docFile.getTabs();
  let matchesCount = 0
  let blockMatchesCount = 0

  for (const tab of tabs) {

    let bodyText = tab.asDocumentTab().getBody().getText();

    // Vérifier la présence de placeholders sous forme {{...}}
    const placeholderPattern = /\{\{[^{}]+\}\}/g;
    const matches = bodyText.match(placeholderPattern);
    
    if (matches && matches.length > 0) {
      matchesCount =+ matches.length;
    }

    // Vérifier la présence de placeholders sous forme £
    const blockPlaceholderPattern = /£(F|S|O|K)/g;
    const blockMatches = bodyText.match(blockPlaceholderPattern);
    
    if (blockMatches && blockMatches.length > 0) {
      blockMatchesCount =+ blockMatches.length;
    }
    
  }

  return matchesCount + blockMatchesCount;
}

// Tranforms bloc £SI COLUMN_NAME=VALUE£ to £OK or £KO in all tabs using current line values
// leave it without change if COLUMN_NAME is unknown
function evaluateEqualConditions(tabs, values) {
    
    const condition_regex = '(\W|^)£SI[\t\n\f\r ]([^=]+)=([^£]+)£(\W|$)'

    for (const tab of tabs) {

      let body = tab.asDocumentTab().getBody();

      let condition_matches = body.findText(condition_regex);

      while (condition_matches) {
        const matchText = condition_matches.getElement().asText();
        const matchContent = matchText.getText();
        Logger.log(`Bloc condition détecté: ${matchContent}`);
        
        const processedCondition = matchContent.replace(new RegExp(condition_regex, "m"), (fullMatch, start, column, value, end) => {
          if (values[column]) {
            if (values[column].toString() == value.toString()) {
                Logger.log(`Bloc (${column}) : ${values[column]} == ${value} remplacé par £OK`);
                return '£OK';
            }
            else {
                Logger.log(`Bloc (${column}) : ${values[column]} == ${value} remplacé par £KO`);
                return '£KO';
            }
          }
          Logger.log(`Bloc (${column}): ${column} introuvable, non remplacé`);
          return fullMatch;
        });
        
        matchText.setText(processedCondition);

        condition_matches = body.findText(condition_regex, condition_matches);
      }
    }
}

// Tranforms bloc £SI COLUMN_NAME<>VALUE£ to £OK or £KO in all tabs using current line values
// leave it without change if COLUMN_NAME is unknown
function evaluateNotEqualConditions(tabs, values) {
    
    const condition_regex = '(\W|^)£SI[\t\n\f\r ]([^<]+)<>([^£]+)£(\W|$)'

    for (const tab of tabs) {

      let body = tab.asDocumentTab().getBody();

      let condition_matches = body.findText(condition_regex);

      while (condition_matches) {
        const matchText = condition_matches.getElement().asText();
        const matchContent = matchText.getText();
        Logger.log(`Bloc condition détecté: ${matchContent}`);
        
        const processedCondition = matchContent.replace(new RegExp(condition_regex, "m"), (fullMatch, start, column, value, end) => {
          if (values[column]) {
            if (values[column].toString() != value.toString()) {
                Logger.log(`Bloc (${column}) : ${values[column]} <> ${value} remplacé par £OK`);
                return '£OK'; // Condition meet, we keep the content
            }
            else {
                Logger.log(`Bloc (${column}) : ${values[column]} <> ${value} remplacé par £KO`);
                return '£KO'; // Bloc removed
            }
          }
          Logger.log(`Bloc (${column}): ${column} introuvable, non remplacé`);
          return fullMatch; // Invalid condition, we keep the content
        });

        matchText.setText(processedCondition);

        condition_matches = body.findText(condition_regex, condition_matches);
      }
    }
}

// Transform blocs £OKxxx£FIN in xxx
function cleanOKBocks(body) {
  let startElement = body.findText('£OK');

  while (startElement) {
    const endElement = body.findText('£FIN', startElement);
    if (endElement) {
      const startIndex = startElement.getStartOffset();
      const endIndex = endElement.getStartOffset();
      startElement.getElement().asText().deleteText(startIndex, startIndex+2);
      endElement.getElement().asText().deleteText(endIndex, endIndex+3);
    }
    else {
       Logger.log(`Erreur = £OK sans bloc £FIN correspondant`);
    }
    startElement = body.findText('£OK', startElement);
  }
}

// Remove blocs £KOxxx£FIN
// Only way found to fix multiple lines is to use the batch API
function removeKOBocks(text, docId) {
  const r = new RegExp('£KO[^£]*£FIN', "g");
  const matches = text.match(r);
  if (!matches || matches.length == 0) {
    return;
  }
  else {
    const requests = matches.map(text => ({ replaceAllText: { containsText: { matchCase: false, text }, replaceText: "" } }));
    Docs.Documents.batchUpdate({ requests }, docId);
  }

}

function logState(sheet, data, message) {
  Logger.log(message);
  data[1][6] += `\n${message}`
  sheet.getRange('G2:G2').setValue(data[1][6])
}

function generateDocsAndPDFs() {

  // Open active sheep
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
    // Google Docs ID
  const templateId = data[1][3];

  // Output Folder ID
  const outputFolderId = data[1][4];
  Logger.log(`Open ${outputFolderId}`);
  const folder = DriveApp.getFolderById(outputFolderId);

  // Get columns headers (line 4)
  const headers = data[3];

  // Reset state
  data[1][6] = "";

  for (let i = data[1][1]-1; i < data[1][2] ; i++) {

    const row = data[i];
    const values = {};

    // Associate header to columns
    headers.forEach((header, index) => {
      values[header] = row[index];
    });

    Logger.log(`Traitement de la ligne ${i}: ${JSON.stringify(values)}`);

    // Create a new Google Docs
    const doc = DriveApp.getFileById(templateId).makeCopy(`Courrier-${values.nom}`, folder);
    const docFile = DocumentApp.openById(doc.getId());
    
    let tabs = docFile.getTabs();
    
    // Pre-evaluate conditions
    evaluateEqualConditions(tabs, values);
    evaluateNotEqualConditions(tabs, values);

    fulltext = ""
    for (const tab of tabs) {

      let body = tab.asDocumentTab().getBody()
    
      cleanOKBocks(body);

      // Replace {{COLUMN_NAME blocs}} if COLUMN_NAME exists
      for (const key in values) {
        if (values[key]) {
            const placeholder = `{{${key}}}`;
            body.replaceText(placeholder, values[key]);
        }
      }

      // Used next to clear KO blocs
      fulltext += body.asText().getText();
    }

    // Close generated doc
    docFile.saveAndClose();

    // Request KO cleanup
    removeKOBocks(fulltext, doc.getId());

    // Check all placeholders were replaced
    let remainingPlaceholders = countPlaceholders(doc.getId())

    if (remainingPlaceholders > 0) {
      logState(sheet, data, `${doc.getName()} : Erreur : ${remainingPlaceholders} placeholders non remplacés subsistent dans le document`);
    } 
    else {
      // PDF export
      const pdf = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
      folder.createFile(pdf).setName(`Courrier-${values.nom}`);
      logState(sheet, data, `${doc.getName()} : OK : Document ${pdf.getName()} généré avec succès`);
    } 
  }
}

