// Julien BAYLE, 2025
// https://github.com/julienbayle/apps_script_doc_generator/edit/main/code.js

// Documentation Google Apps Script
//https://developers.google.com/apps-script/reference/document/body

// Counts placeholders {{xxx}}, £IF, £FIN, £OK, £KO in all tabs
function countPlaceholders(tabs) {

  let matchesCount = 0
  let blockMatchesCount = 0

  for (const tab of tabs) {

    let bodyText = tab.asDocumentTab().getBody().getText();

    // Vérifier la présence de placeholders sous forme {{...}}
    const placeholderPattern = /\{\{[^{}]+\}\}/g;
    const matches = bodyText.match(placeholderPattern);
    
    if (matches && matches.length > 0) {
      matchesCount += matches.length;
    }

    // Vérifier la présence de placeholders sous forme £
    const blockPlaceholderPattern = /£(F|S|O|K)/g;
    const blockMatches = bodyText.match(blockPlaceholderPattern);
    
    if (blockMatches && blockMatches.length > 0) {
      blockMatchesCount += blockMatches.length;
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

function cleanEvaluatedBocksInTables(body) {
  const tables = body.getTables();

  for (let tableIndex = 0; tableIndex < tables.length; tableIndex++) {
    const table = tables[tableIndex];
    const rowCount = table.getNumRows();

    Logger.log(`1 table ${tableIndex + 1} with ${rowCount} rows`);

    for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
      const row = table.getRow(rowIndex);
      const cellCount = row.getNumCells();

      for (let cellIndex = 0; cellIndex < cellCount; cellIndex++) {
        const cell = row.getCell(cellIndex);
        let cellElementCount = cell.getNumChildren();

        let removeNext = false;
        for (let cellChildrenIndex = 0; cellChildrenIndex < cellElementCount; cellChildrenIndex++) {
          let paragraph = cell.getChild(cellChildrenIndex)

          if (paragraph.findText("£FIN")) {
            Logger.log(`remove ${paragraph.getText()}`);
            cell.removeChild(paragraph)
            cellChildrenIndex--
            cellElementCount--
            removeNext = false
          }

          if (paragraph.findText("£OK")) {
            Logger.log(`remove ${paragraph.getText()}`);
            cell.removeChild(paragraph)
            cellChildrenIndex--
            cellElementCount--
            removeNext = false
          }

          if (removeNext) {
            Logger.log(`remove ${paragraph.getText()}`);
            cell.removeChild(paragraph)
            cellChildrenIndex--
            cellElementCount--
          }
          
          if (paragraph.findText("£KO")) {
            Logger.log(`remove ${paragraph.getText()}`);
            cell.removeChild(paragraph)
            cellChildrenIndex--
            cellElementCount--
            removeNext = true
          }
        }
      }
    }
  }
}

// Transform blocs £OKxxx£FIN in "xxx"
// Transform blocs £K0xxx£FIN in ""
function cleanEvaluatedBocks(body) {

  let paragraphs = body.getParagraphs();
  Logger.log(`${paragraphs.length} paragraphes`);

  let removeNext = false;

  for (const paragraph of paragraphs) {
  
    if (paragraph.findText("£FIN")) {
      Logger.log(`remove ${paragraph.getText()}`);
      body.removeChild(paragraph)
      removeNext = false
    }

    if (paragraph.findText("£OK")) {
      Logger.log(`remove ${paragraph.getText()}`);
      body.removeChild(paragraph)
      removeNext = false
    }

    if (removeNext) {
      Logger.log(`remove ${paragraph.getText()}`);
      body.removeChild(paragraph)
    }
    
    if (paragraph.findText("£KO")) {
      Logger.log(`remove ${paragraph.getText()}`);
      body.removeChild(paragraph)
      removeNext = true
    }
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
    const doc = DriveApp.getFileById(templateId).makeCopy(`Courrier-${values.INITIALES}`, folder);
    const docFile = DocumentApp.openById(doc.getId());
    
    let tabs = docFile.getTabs();
    
    // Pre-evaluate conditions
    evaluateEqualConditions(tabs, values);
    evaluateNotEqualConditions(tabs, values);

    for (const tab of tabs) {

      let body = tab.asDocumentTab().getBody()
    
      cleanEvaluatedBocksInTables(body)
      cleanEvaluatedBocks(body)

      // Replace {{COLUMN_NAME blocs}}% if COLUMN_NAME exists
      for (const key in values) {
        if (values[key] || values[key] === 0) {
            const placeholder = `{{${key}}}%`;
            let val_percent = Math.round(100*parseFloat(values[key])).toString()
            body.replaceText(placeholder, `${val_percent}%`);
        }
      }

      // Replace {{COLUMN_NAME blocs}} if COLUMN_NAME exists
      for (const key in values) {
        if (values[key] || values[key] === 0) {
            const placeholder = `{{${key}}}`;
            
            if (key.includes("DATE")) {
                const options = { day: '2-digit', month: '2-digit', year: 'numeric' };
                let formattedDate = new Intl.DateTimeFormat('fr-FR', options).format(values[key]);
                body.replaceText(placeholder, formattedDate);
            }
            else {
                body.replaceText(placeholder, values[key]);
            }
        }
      }
    }

    // Check all placeholders were replaced
    let remainingPlaceholders = countPlaceholders(tabs)

    // Close generated doc
    docFile.saveAndClose();

    if (remainingPlaceholders > 0) {
      logState(sheet, data, `${doc.getName()} : Erreur : ${remainingPlaceholders} placeholders restants`);
    } 
    else {
      // PDF export
      const pdf = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
      folder.createFile(pdf).setName(`Courrier-${values.INITIALES}`);
      logState(sheet, data, `${doc.getName()} : OK : Document ${pdf.getName()} généré avec succès`);
    } 
  }
}


function test() {
    const docFile = DocumentApp.openById("1ixNZME6mAHlYFoXdJy9jrHxkUnRm4UpZFKVWaIxkwmQ");
    let tabs = docFile.getTabs();
    for (const tab of tabs) {
      let body = tab.asDocumentTab().getBody()
      cleanEvaluatedBocksInTables(body)
      cleanEvaluatedBocks(body);
    }
}
