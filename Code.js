/**
 * Configuration constants
 * Update these values to customize the email template and form URL
 */
const CONFIG = {
  AUTHOR: "Your Name or Organization",
  FORM_URL: "https://forms.google.com/your-form-url-here",
  EMAIL_SUBJECT_PREFIX: "Links from: "
};

/**
 * Form field indices mapping
 * These correspond to the Google Form response columns
 */
const FORM_FIELDS = {
  PRIMARY_EMAIL: 1,
  CC_START: 2,
  CC_END: 11,
  SLIDE_URL: 3,
  MESSAGE_TO_RECIPIENTS: 11
};

/**
 * Main function triggered by Google Form submission
 * @param {Object} e - The event object from form submission
 */
function myFunction(e) {
  try {
    // Validate event object
    if (!e || !e.values) {
      Logger.log('Error: Invalid event object');
      return;
    }

    // Extract form data
    const primeEmail = e.values[FORM_FIELDS.PRIMARY_EMAIL];
    const slideURL = e.values[FORM_FIELDS.SLIDE_URL];
    const formNote = e.values[FORM_FIELDS.MESSAGE_TO_RECIPIENTS] || "";

    // Validate required fields
    if (!primeEmail || !slideURL) {
      Logger.log('Error: Missing required fields (email or slide URL)');
      return;
    }

    // Collect CC emails (skip the slide URL field at index 3)
    const ccEmails = [];
    for (let i = FORM_FIELDS.CC_START; i < FORM_FIELDS.CC_END; i++) {
      if (i !== FORM_FIELDS.SLIDE_URL && e.values[i] && e.values[i].trim() !== "") {
        ccEmails.push(e.values[i].trim());
      }
    }

    // Build pretext message if provided
    const pretext = formNote.length > 0
      ? `Their message to you: "${formNote}"<br><br>`
      : "";

    // Extract links and send email
    getLinksFromSlides(slideURL, primeEmail, ccEmails, pretext);

  } catch (error) {
    Logger.log('Error in myFunction: ' + error.toString());
    // Optionally send error notification email
  }
}

/**
 * Formats a link for HTML email output
 * @param {TextRange} link - The link object from Google Slides
 * @param {boolean} indent - Whether to indent the link (for speaker notes)
 * @returns {string} Formatted HTML string for the link
 */
function formatLink(link, indent) {
  const linkText = link.asString();
  const linkUrl = link.getTextStyle().getLink().getUrl();
  const indentation = indent ? "&emsp;&emsp;" : "";

  let formattedLink;
  if (linkText === linkUrl) {
    formattedLink = `Link: ${linkUrl}`;
  } else if (linkUrl.startsWith('mailto:')) {
    formattedLink = `Email: ${linkText}`;
  } else {
    formattedLink = `${linkText}: ${linkUrl}`;
  }

  return `${indentation}${formattedLink}<br><br>`;
}

/**
 * Formats slide title for HTML email output
 * @param {number} slideId - The slide number
 * @param {string} slidesTitle - The title of the slide
 * @returns {string} Formatted HTML string for the slide title
 */
function formatSlideTitle(slideId, slidesTitle) {
  return `<h3>[ ${slideId} ] :: <strong>${slidesTitle}</strong></h3>`;
}

/**
 * Extracts all links from a Google Slides presentation and sends them via email
 * @param {string} slideURL - URL of the Google Slides presentation
 * @param {string} primeEmail - Primary recipient email address
 * @param {Array<string>} ccEmails - Array of CC email addresses
 * @param {string} pretext - Optional message to include at the top of the email
 */
function getLinksFromSlides(slideURL, primeEmail, ccEmails, pretext) {
  try {
    // Open the presentation
    const presentation = SlidesApp.openByUrl(slideURL);
    const prezTitle = presentation.getName();
    const slides = presentation.getSlides();

    // Build HTML content with all links
    const linksHtml = extractLinksFromAllSlides(slides);

    // Send email with extracted links
    sendLinksEmail(primeEmail, ccEmails, prezTitle, slideURL, pretext, linksHtml);

  } catch (error) {
    Logger.log('Error in getLinksFromSlides: ' + error.toString());
    // Attempt to notify user of failure
    try {
      MailApp.sendEmail({
        to: primeEmail,
        subject: "Error: Unable to extract links from slides",
        body: `There was an error processing your request: ${error.toString()}\n\nPlease verify that the slide URL is correct and that it's shared with the appropriate Google account.`
      });
    } catch (emailError) {
      Logger.log('Failed to send error notification email: ' + emailError.toString());
    }
  }
}

/**
 * Extracts links from all slides in a presentation
 * @param {Array} slides - Array of slide objects
 * @returns {string} HTML string containing all formatted links
 */
function extractLinksFromAllSlides(slides) {
  const linksParts = [];
  let slideNumber = 1;

  slides.forEach(function (slide) {
    // Skip hidden/skipped slides
    if (slide.isSkipped()) {
      slideNumber++;
      return;
    }

    const slideLinks = extractLinksFromSlide(slide, slideNumber);
    if (slideLinks) {
      linksParts.push(slideLinks);
    }

    slideNumber++;
  });

  return linksParts.join('');
}

/**
 * Extracts all links from a single slide
 * @param {Slide} slide - The slide object
 * @param {number} slideNumber - The slide number
 * @returns {string|null} HTML string with links or null if no links found
 */
function extractLinksFromSlide(slide, slideNumber) {
  const linksParts = [];
  let slideTitle = "";
  let titleAdded = false;

  // Helper function to add title if needed
  const addTitleIfNeeded = () => {
    if (!titleAdded) {
      linksParts.push(formatSlideTitle(slideNumber, slideTitle));
      titleAdded = true;
    }
  };

  // Get slide title from first shape if available
  const shapes = slide.getShapes();
  if (shapes.length > 0) {
    slideTitle = shapes[0].getText().asString();
  }

  // Extract links from shapes
  shapes.forEach(function (shape) {
    const textRange = shape.getText();
    const links = textRange.getLinks();

    if (links.length > 0) {
      addTitleIfNeeded();
      links.forEach((link) => linksParts.push(formatLink(link, false)));
    }
  });

  // Extract links from tables
  const tables = slide.getTables();
  tables.forEach(function (table) {
    const rows = table.getNumRows();
    const cols = table.getNumColumns();

    for (let i = 0; i < rows; i++) {
      for (let j = 0; j < cols; j++) {
        const cellText = table.getCell(i, j).getText();
        const cellLinks = cellText.getLinks();

        if (cellLinks.length > 0) {
          addTitleIfNeeded();
          cellLinks.forEach((link) => linksParts.push(formatLink(link, false)));
        }
      }
    }
  });

  // Extract links from speaker notes
  const speakerNotes = slide.getNotesPage();
  const notesShape = speakerNotes.getSpeakerNotesShape();
  const notesText = notesShape.getText();
  const notesLinks = notesText.getLinks();

  if (notesLinks.length > 0) {
    addTitleIfNeeded();
    linksParts.push("<h4>&emsp;&emsp;<em>Speaker Note Links:</em></h4>");
    notesLinks.forEach((link) => linksParts.push(formatLink(link, true)));
  }

  return titleAdded ? linksParts.join('') : null;
}

/**
 * Sends an email with the extracted links
 * @param {string} toEmail - Primary recipient email
 * @param {Array<string>} ccEmails - Array of CC emails
 * @param {string} prezTitle - Presentation title
 * @param {string} slideURL - Presentation URL
 * @param {string} pretext - Optional pretext message
 * @param {string} linksHtml - HTML content with links
 */
function sendLinksEmail(toEmail, ccEmails, prezTitle, slideURL, pretext, linksHtml) {
  // Build CC string
  const ccString = Array.isArray(ccEmails) && ccEmails.length > 0
    ? ccEmails.join(',')
    : "";

  // Build email HTML body
  const htmlBody = `
    <h3>This message was sent to you by ${toEmail} from ${CONFIG.AUTHOR}'s
    <a href="${CONFIG.FORM_URL}">Google Slide URL Extraction Tool</a></h3>
    ${pretext}
    Links from <a href="${slideURL}">${prezTitle}</a><br><br>
    Use these links to share in the chat window:<br><br>
    ${linksHtml}
  `;

  // Send email
  MailApp.sendEmail({
    to: toEmail,
    cc: ccString,
    subject: CONFIG.EMAIL_SUBJECT_PREFIX + prezTitle,
    htmlBody: htmlBody
  });

  Logger.log(`Email sent successfully to ${toEmail}`);
}