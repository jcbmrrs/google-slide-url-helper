// Initialize variables
let printedLinks = [];
let carboncopy = [];
let emailLinks = "";
let slideURL = "";
let pretext = "";

// Function triggered by an event
function myFunction(e) {
  // Get the primary email from the event values
  var primeEmail = e.values[1];
  
  // Loop through the event values starting from index 2
  for (let i = 2; i < 11; i++) {
    if (i == 3) {
      // Skip index 3
    } else {
      if (e.values[i] == "") {
        // If the value is empty, skip it
      } else {
        // Add non-empty values to the carboncopy array
        carboncopy.push(e.values[i]);
      }
    }
  }
  
  // If carboncopy array is empty, set it to an empty string
  if (carboncopy.length === 0) {
    carboncopy = "";
  }
  
  // Get slideURL from the event values
  slideURL = e.values[3];
  
  // Get formNote from the event values
  let formNote = e.values[11];
  
  // Set pretext based on the length of formNote
  pretext = formNote.length > 0 ? ("Their message to you: \"" + formNote + "\"<br><br>") : "";
  
  // Call getLinksFromSlides function
  getLinksFromSlides(slideURL, primeEmail, carboncopy);
}

// Function to process each link
function linkBreaker(link, indent) {
  let addition = "";
  
  if (indent) {
    addition += "&emsp;&emsp;";
  }
  
  if (link.asString() == link.getTextStyle().getLink().getUrl()) { 
    addition += ("Link: " + link.getTextStyle().getLink().getUrl());
  } else if (link.getTextStyle().getLink().getUrl().startsWith('mailto:')) {
    addition += ("Email: " + link.asString());
  } else { 
    addition += (link.asString() + ": " + link.getTextStyle().getLink().getUrl());
  }
  
  addition += "<br><br>";
  emailLinks += addition;
}

// Function to print slide title
function printTitle(slideId, slidesTitle) {
  emailLinks += "<h3>[ " + slideId + " ] :: <strong>" + slidesTitle + "</strong></h3>";
}

// Function to extract links from slides
function getLinksFromSlides(slideURL, primeEmail, carboncopy) {
  // Open the presentation using the slideURL
  var presentation = SlidesApp.openByUrl(slideURL);
  
  // Get the presentation title
  var prezTitle = presentation.getName();
  
  // Get all slides in the presentation
  var slides = presentation.getSlides();
  
  // Initialize slideId to 1
  let slideId = 1;
  
  // Iterate over each slide
  slides.forEach(function (slide) {
    // Skip skipped slides
    if (slide.isSkipped()) {
      // Do nothing for skipped slides
    } else {
      let slidesTitle = "";
      let titleShared = false;
      let links = [];
      
      // Get all shapes in the slide
      let shapes = slide.getShapes();
      
      // Check if there are shapes in the slide
      if (shapes.length > 0) {
        slidesTitle = shapes[0].getText().asString();
      }
      
      // Iterate over each shape
      shapes.forEach(function (shape) {
        // Get the text content of the shape
        var textRange = shape.getText();
        
        // Get the links in the text content
        links = textRange.getLinks();
        
        // Print all links found
        if (links.length != 0 && titleShared == false) {
          printTitle(slideId, slidesTitle);
          titleShared = true;
        }
        
        links.forEach((link) => linkBreaker(link, false));
      });

      // Get all tables in the slide
      var tables = slide.getTables();
      
      // Iterate over each table
      tables.forEach(function (table) {
        var rows = table.getNumRows();
        var cols = table.getNumColumns();
        
        // Iterate over each cell in the table
        for (let i = 0; i < rows; i++) { 
          for (let j = 0; j < cols; j++) {
            var tableTextRange = table.getCell(i,j).getText();
            var tablelinks = tableTextRange.getLinks();
            
            if (tablelinks.length != 0 && titleShared == false) {
              printTitle(slideId, slidesTitle);
              titleShared = true;
            }
            
            tablelinks.forEach((link) => linkBreaker(link, false));
          }
        }
      });

      // Get the speaker notes for the slide
      var speakerNotes = slide.getNotesPage(); 
      var notesShape = speakerNotes.getSpeakerNotesShape();
      var notesText = notesShape.getText();
      links = notesText.getLinks();
      
      if (links != "") {
        if (links.length != 0 && titleShared == false) {
          printTitle(slideId, slidesTitle);
          titleShared = true;
        }
        
        emailLinks += "<h4>&emsp;&emsp;<em>Speaker Note Links:</em></h4>";
        links.forEach((link) => linkBreaker(link, true));
      }
    }
    
    // Increment the slideId
    slideId++;
  });
  
  // Convert carboncopy array to a string
  let goodCC = "";
  
  if (carboncopy == []) { 
    // If carboncopy is empty, do nothing
  } else {
    goodCC  = carboncopy.join(',');
  }
  
  // Send an email with the extracted links
  MailApp.sendEmail({
    to: primeEmail,
    cc: goodCC,
    subject: "Links from: " + prezTitle,
    htmlBody: "<h3>This message was sent to you by " + primeEmail + " from Jacob's <a href=\"https://forms.gle/5NoXZgSNEhEz85iU6\">Google Slide URL Extraction Tool</a></h3>" + pretext + "Links from <a href=" + slideURL + "\">" + prezTitle + "</a><br><br>Use these links to share in the chat window:<br><br>" + emailLinks
  });
}