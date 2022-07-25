function collectInfo() {

  // Find title and author information from Spreadsheet

  var ss = SpreadsheetApp.openById("1yKN7kcae6Z4Dcv5j3RZfIVlX_sbxHGMpIw_gNu7cZ6s");
  var ws = ss.getSheetByName("Form Responses");
  var data = ws.getDataRange().getValues();
  
  var lastRow = ws.getLastRow();
  var titleColumn = 1
  var authorColumn = 2


  //  Replace blank spaces in data with "+" for URL 


  for (var i=1; i< lastRow; i++){
    var title = data[i][titleColumn];
    var author = data[i][authorColumn];

    var titleFormat = title.toString().replace(/\s/g, "+")
    var authorFormat = author.toString().replace(/\s/g, "+")
  

    //  Connect to API

   var url = "https://www.googleapis.com/books/v1/volumes?q=intitle:\"" + titleFormat + "\" inauthor:\"" + authorFormat + "\"" + "&country=ZA";

  var urlEncoded = encodeURI(url);
  // Logger.log(urlEncoded)
  var response = UrlFetchApp.fetch(urlEncoded);
  var parse = JSON.parse(response);
  

    // Find useful data from JSON 

    var items = parse.items;

      var isbn = items[0].volumeInfo.industryIdentifiers[0].identifier;
      var averageRating = items[0].volumeInfo.averageRating;
      var pageCount = items[0].volumeInfo.pageCount;
      var publisher = items[0].volumeInfo.publisher;
      var title = items[0].volumeInfo.title;
      var publishedDate = items[0].volumeInfo.publishedDate;
      var description = items[0].volumeInfo.description;
      var authors = items[0].volumeInfo.authors;

  
  // Copy and edit Template Document

  var template = DriveApp.getFileById("1D15nqqnyHSKcqc-yjs7IdGGZUyyIIh4-8bljTa6H8iQ");
  var destination = DriveApp.getFoldersByName("Summaries").next();


    // Prevent creation of duplicate files

    if(destination.getFilesByName(title).hasNext() == true){
      Logger.log("Prevented file duplication")

    } else {
      var copy = template.makeCopy(title , destination );
    
      var summary = DocumentApp.openById(copy.getId());
      var body = summary.getBody();
        body.replaceText("{{title}}",title);
        body.replaceText("{{authors}}",authors);
        body.replaceText("{{isbn}}",isbn);
        body.replaceText("{{averageRating}}",averageRating);
        body.replaceText("{{pageCount}}",pageCount);
        body.replaceText("{{publisher}}",publisher);
        body.replaceText("{{publishedDate}}",publishedDate);
        body.replaceText("{{description}}",description);
    }
  }
  
}







