/**
 * createAnswerReport
 */
function createAnswerReport() {
  var docID = '';
  var spreadID = '';
  var defaultColor = '#000000';
  var answerColor = '#4a86e8';
  var footerColor = '#4a86e8';
  var title = '答案卷';
  
  var doc = DocumentApp.openById(docID);
  var body = doc.getBody();
  body.setMarginBottom(10);
  body.setMarginTop(10);
  body.setMarginLeft(10);
  body.setMarginRight(10);
  
  body.appendParagraph(title);
  body.setText('#執行時間: ' + new Date());
  
  var Header = body.appendParagraph(title);
  Header.setFontFamily(DocumentApp.FontFamily.ARIAL);
  //Header.setFontSize(16);
  Header.setForegroundColor(defaultColor);
  Header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  Header.setBold(true);

  body.appendPageBreak();
  
  var spread = SpreadsheetApp.openById(spreadID);
  var sheet = spread.getActiveSheet();
  var columns = sheet.getLastColumn();
  var rows = sheet.getLastRow();
  var title = sheet.getRange(1, 2, 1, columns - 1).getValues()[0];
  var data = sheet.getRange(2, 2, rows - 1, columns - 1).getValues();
  
  body.setBold(false);
  for(var j = 0; j < rows - 1; j++) {
    var listID = null;
    for(var i = 0; i < columns - 1; i++) {
      var List = body.appendListItem(title[i]);
      if(listID == null)
        listID = List;
      else
        List.setListId(listID);
      List.setFontSize(12);
      List.setForegroundColor(defaultColor);
      
      var Body = body.appendParagraph('<content>'+data[j][i]+'</content>');
      Body.setForegroundColor(answerColor);
      Body.setIndentStart(60);
      Body.setIndentFirstLine(60);
      Body.setSpacingBefore(6);
    }
    body.appendPageBreak();
  }
  
  var endContent = body.appendParagraph('#結束時間: ' + new Date());
  endContent.setForegroundColor(defaultColor);
  
  var footer = doc.getFooter() || doc.addFooter();
  doc.getFooter().removeFromParent();
  footer = doc.addFooter();
  var divider = footer.appendHorizontalRule();
  var footerText = footer.appendParagraph('');
  footerText.setFontSize(9);
  footerText.setForegroundColor(footerColor);
  footerText.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  return doc;
}