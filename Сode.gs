function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
  .addItem('�����������', 'Copy')
  .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function Copy()
{
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetName = sheet.getName();
  var destination = SpreadsheetApp.openById('1nLoAZ5sgDiJ_gfLORkcdKRT2W-dRJpeXBPxwzBCRQwg');
  sheet.copyTo(destination);
  var sourceFormulas = sheet.getDataRange().getFormulas();
  destination.getSheetByName(sheetName + ' (�����)').setName(sheetName);
  var destinationSheet = destination.getSheetByName(sheetName);
  for (i = 1; i <= sourceFormulas.length; i++) {
    for (j = 1; j <= sourceFormulas[0].length; j++) {
      if (sourceFormulas[i-1][j-1] != '')
      {
      destinationSheet.getRange(i, j).setValue(sourceFormulas[i-1][j-1]);
      }
    }
  }
  var destinationNamedRanges = destinationSheet.getNamedRanges();
  for (i = 0; i < destinationNamedRanges.length; i++)
  {
    var name = destinationNamedRanges[i].getName();
    var range = destinationNamedRanges[i].getRange();
    destinationNamedRanges[i].remove();
    destination.setNamedRange(name, range);
  }
}
