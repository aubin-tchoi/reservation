// Project : DVP_GForms_generation
// Author : Aubin Tchoï
// This script is meant to be bound to a certain format of Google Sheets

// Creating the menu
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Générer un Google Forms')
  .addItem('Prêt de matériel', 'gen_Forms')
  .addItem('Annulation de prêt', 'cancellation')
  .addToUi();
}

// Generating a Google Forms with data taken from the Sheets
function gen_Forms() {
  const ui = SpreadsheetApp.getUi(),
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Feuille 1"),
      data = sheet.getRange(2, 2, (sheet.getLastRow() - 1), 3).getValues(),
      heads = data.shift(),
      obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {}))),
      forms = FormApp.create(`DVP - prêt de ${sheet.getRange(3, 9, 1, 1).getValues()[0][0].toLowerCase()}`)
      .setDescription(sheet.getRange(4, 9, 1, 1).getValues()[0][0])
      .setConfirmationMessage(sheet.getRange(5, 9, 1, 1).getValues()[0][0]);
  
  // Loading screen
  let htmlLoading = HtmlService
  .createHtmlOutput(`<img src="https://www.demilked.com/magazine/wp-content/uploads/2016/06/gif-animations-replace-loading-screen-14.gif" alt="Loading" width="531" height="299">`)
  .setWidth(540)
  .setHeight(350);
  ui.showModelessDialog(htmlLoading, "Chargement du Google Forms..");

  // First questions
  let first_name = forms.addTextItem()
  .setTitle('Quel est votre prénom ?')
  .setRequired(true),
      last_name = forms.addTextItem()
  .setTitle('Quel est votre nom ?')
  .setRequired(true),
      tel_num = forms.addTextItem()
  .setTitle('Quel est votre numéro de portable ?')
  .setRequired(true);

  // That is gonna be the "main" question (each answer is linked to a distinct section)
  var item = forms.addListItem()
  .setTitle('Que souhaitez vous emprunter ?')
  .setRequired(true),
      item_choices = [];

  Logger.log(obj); // Data format is similar to a JSON

  // Each answer to the question "Que souhaitez vous emprunter" generates a distinct section
  obj.forEach(function(row) {
    
    // new section here + 2 additional questions before forms submission
    let sec = forms.addPageBreakItem()
    .setTitle(row["Bien prêté"])
    .setGoToPage(FormApp.PageNavigationType.SUBMIT),
        
        hor = forms.addCheckboxItem()
    .setTitle('Quand en aurez vous besoin ?')
    .setRequired(true),
        
        rmq = forms.addTextItem()
    .setTitle('Une remarque en particulier ?'),
        
        // As seen in the following regexes, the first 2 numbers in column C will be recognized as the start and the end of the time period.
        start_hour = Math.floor(row["Plage horaire"].match(/([0-9]+)h/gm)[0].match(/[0-9]+/gm)[0]),
        end_hour = Math.floor(row["Plage horaire"].match(/([0-9]+)h/gm)[1].match(/[0-9]+/gm)[0]),
        time_slot = Math.floor(row["Durée d'un créneau"].match(/([0-9]+)/gm)[0]);
    
    Logger.log(`Bien prêté : row["Bien prêté"]`);
    Logger.log(`Heure début : ${start_hour}`);
    Logger.log(`Heure fin : ${end_hour}`);
    Logger.log(`Durée créneau : ${time_slot}`);
    
    // Hour slots are created by splitting the time period in slots each time_slot long
    let hor_choices = [];
    for (let h_n = start_hour; h_n < end_hour; h_n += time_slot) {
      hor_choices.push(hor.createChoice(`${h_n}h à ${Math.min(h_n + time_slot, end_hour)}h`));
    }
    hor.setChoices(hor_choices);

    // The item is added to a list that will become the list of possible choices to the question "Que souhaitez vous emprunter ?"
    item_choices.push(item.createChoice(row["Bien prêté"], sec));
  item.setChoices(item_choices);
  
  });

  // Setting up the active spreadsheet as a destination in order to be able to update the form.
  forms.setDestination(FormApp.DestinationType.SPREADSHEET, SpreadsheetApp.getActiveSpreadsheet().getId());
  sheet.getRange(2, 9, 1, 1).setValues([[forms.getId()]]).setFontColor('white'); // I'm not proud of this
  
  // This trigger updates the form after every submission, removing the slots he chose
  ScriptApp.newTrigger('update_form')
  .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
  .onFormSubmit()
  .create();

  // Dialog box to print the links (one editor link and a reader link)
  let htmlOutput = HtmlService
  .createHtmlOutput(`<span style="font-family: 'trebuchet ms', sans-serif;">Voci le lien éditeur : <a href = "${forms.getEditUrl()}">${forms.getEditUrl()}</a><br/><br/> Voici le lien lecteur : <a href = "${forms.getPublishedUrl()}">${forms.getPublishedUrl()}</a>.<br/><br/><br/>&nbsp; La bise.<br/>${"&nbsp; ".repeat(16)}</span></span><img src="http://developponts.enpc.org/images/logo_petit.png" alt="DVP" width="115", height="100"><span style='font-size: 12pt;'>`);
  ui.showModelessDialog(htmlOutput, "Un Google Forms a été créé.")
}

// Updating the Google Forms by removing occupied slots
function update_form() {
  const ui = SpreadsheetApp.getUi(),
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Feuille 1"),
      data = sheet.getRange(2, 2, (sheet.getLastRow() - 1), 3).getValues(),
      heads = data.shift(),
      obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));
  if (sheet.getRange(2, 9, 1, 1).getValues()[0][0] == "") {
    return;
  }
  const forms = FormApp.openById(sheet.getRange(2, 9, 1, 1).getValues()[0][0]),
      responses = forms.getResponses(),
      last_response = responses[responses.length - 1].getItemResponses();
  
  last_response
  .filter(response => response.getItem().getTitle() == 'Quand en aurez vous besoin ?')
  .forEach(function(resp) {
    let old_choices = forms.getItemById(resp.getItem().getId()).asCheckboxItem().getChoices(),
        new_choices = old_choices.filter(choice => !resp.getResponse().includes(choice.getValue()));
    forms.getItemById(resp.getItem().getId()).asCheckboxItem().setChoices(new_choices);
  }); 
}

// Cancelling a reservation
function cancellation() {
  const ui = SpreadsheetApp.getUi();
  if (sheet.getRange(2, 9, 1, 1).getValues()[0][0] == "") {
    ui.alert("Annulation d'un prêt", "Impossible de trouver le Google Forms associé.", ui.ButtonSet.OK);
    return;
  }
  const forms = FormApp.openById(sheet.getRange(2, 9, 1, 1).getValues()[0][0]),
      first_name = ui.prompt("Annulation d'un prêt", "Quel est le prénom de la personne ?", ui.ButtonSet.OK_CANCEL),
      last_name = ui.prompt("Annulation d'un prêt", "Quel est le nom de la personne ?", ui.ButtonSet.OK_CANCEL),
      formating = str => str.toLowerCase().replace(/[éèê]/gmi, "e"),
      responses = forms.getResponses().filter(resp => formating(resp.getItemResponses().filter(ir => ir.getItem().getTitle() == 'Quel est votre prénom ?')[0].getResponse()) == formating(first_name) && formating(resp.getItemResponses().filter(ir => ir.getItem().getTitle() == 'Quel est votre nom ?')[0].getResponse()) == formating(last_name));
  
  if (responses.length == 0) {
    ui.alert("Annulation d'un prêt", "Impossible de trouver une personne à ce nom.", ui.ButtonSet.OK);
    return;
  }
  
  if (responses.length > 1) {
    let mult = ui.alert("Annulation d'un prêt", "Plusieurs emprunts ont été effectués à ce nom.", ui.ButtonSet.OK_CANCEL);
    if (mult == ui.Button.CANCEL) {
      return;
    }
    // Choose a response here and filter responses
    let resp_choices = "";
    responses.forEach(function(rit) {rit.getItemResponses().filter(response => response.getItem().getTitle() == 'Quand en aurez vous besoin ?').getResponse().forEach(function(shift) {resp_choices.push(shift);})});
    let resp_choice = ui.alert("Annulation d'un prêt", `Quel créneau souhaitez-vous annuler ? (Entrer le numéro correspondant) ${resp_choices.map(it, idx => `\\n ${idx++} ${it}`)}`, ui.ButtonSet.OK_CANCEL);
    responses.filter(rit => rit.getItemResponses().filter(response => response.getItem().getTitle() == 'Quand en aurez vous besoin ?').getResponse() == resp_choices[Math.floor(resp_choices) - 1]);
  }
  
  if (responses.length == 1) {
    // cancellation is an ItemResponse
    let cancellation = responses[0].getItemResponses().filter(response => response.getItem().getTitle() == 'Quand en aurez vous besoin ?'),
      checkbox_item = forms.getItemById(cancellation.getItem().getId()).asCheckBoxItem(),
      shifts = checkbox_item.getChoices();
    
    cancellation.getResponse().forEach(function(hour) {shifts.push(checkbox_item.createChoice(hour));});
    checkbox_item.setChoices(shifts);
  }
  
  ui.alert("Annulation d'un prêt", `Le prêt de ${first_name} {last_name} a bien été annulé !`, ui.ButtonSet.OK);
}
