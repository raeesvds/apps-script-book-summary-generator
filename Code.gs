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

    var item = parse.items[0].volumeInfo;

      var isbn = item.industryIdentifiers[0].identifier;
      var averageRating = item.averageRating;
      var pageCount = item.pageCount;
      var publisher = item.publisher;
      var title = item.title;
      var publishedDate = item.publishedDate;
      var description = item.description;
      var authors = item.authors;

  
  // Copy and edit Template Document

  var template = DriveApp.getFileById("1D15nqqnyHSKcqc-yjs7IdGGZUyyIIh4-8bljTa6H8iQ");
  var destination = DriveApp.getFoldersByName("Summaries").next();


    // Prevent creation of duplicate files 

    if(destination.getFilesByName(title).hasNext() == true){
      Logger.log("Prevented duplicate of " + title)

    } else {
      

      // Create summary and handle exceptions
      
      
      var copy = template.makeCopy(title , destination );
      var summary = DocumentApp.openById(copy.getId());
      var body = summary.getBody();


        if(title){
        body.replaceText("{{title}}",title);
        } else{
        body.replaceText("{{title}}","Unavailabe");
        Logger.log("No Title");
        }

        if(authors){
        body.replaceText("{{authors}}",authors);
        } else{
        body.replaceText("{{authors}}","Unavailable");
        Logger.log("No Authors");
        }

        if(isbn){
        body.replaceText("{{isbn}}",isbn);
        } else{
        body.replaceText("{{isbn}}","Unavailable");
        Logger.log("No ISBN");
        }

        if(averageRating){
        body.replaceText("{{averageRating}}",averageRating)
        } else{
        body.replaceText("{{averageRating}}","Unavailable");
        Logger.log("No Rating");
        }

        if(pageCount){
        body.replaceText("{{pageCount}}",pageCount);
        } else{
          body.replaceText("{{pageCount}}","Unavailable");
          Logger.log("No Page Count");
        }

        if(publisher){
        body.replaceText("{{publisher}}",publisher);
        } else{
        body.replaceText("{{publisher}}","Unavalaible"); 
        Logger.log("No Publisher");         
        }

        if(publishedDate){
        body.replaceText("{{publishedDate}}",publishedDate);
        } else{
        body.replaceText("{{publishedDate}}","Unavailable");
        Logger.log("No Published Date");
        }

        if(description){
        body.replaceText("{{description}}",description);
        } else{
        body.replaceText("{{description}}","Unavailable");
        Logger.log("No Description");
        }

    }
  }
}







