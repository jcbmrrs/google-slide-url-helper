# Google Slide URL Helper

A Google Apps Script tool that automatically extracts all hyperlinks from Google Slides presentations and emails them to you. Perfect for sharing links during Zoom meetings, webinars, or any presentation where you want to provide easy access to referenced URLs.

## Features

- 🔗 **Comprehensive Link Extraction**: Extracts links from:
  - Text boxes and shapes
  - Tables and cells
  - Speaker notes
- 📧 **Email Delivery**: Automatically sends formatted links via email
- 👥 **CC Support**: Send to multiple recipients (up to 8 CC addresses)
- 📝 **Custom Messages**: Include optional notes with your link email
- 🎯 **Smart Formatting**: Links are organized by slide number with clear labeling
- ⚡ **Google Form Integration**: Simple form-based submission
- 🛡️ **Error Handling**: Robust error handling with user notifications

## How It Works

1. User submits a Google Form with:
   - Google Slides URL
   - Primary email address (auto-captured)
   - Optional CC recipients
   - Optional message to recipients

2. The script automatically:
   - Accesses the Google Slides presentation
   - Extracts all hyperlinks from slides, tables, and speaker notes
   - Formats them in an organized HTML email
   - Sends the email to all specified recipients

3. Recipients receive an email with:
   - All links organized by slide number
   - Slide titles for context
   - Separate section for speaker note links
   - Direct link back to the presentation

## Prerequisites

- Google account
- Google Apps Script access
- Google Forms (for submission interface)
- Google Slides to extract links from

## Setup Instructions

### 1. Create a Google Form

Create a new Google Form with the following fields:

1. **Email** (automatically collected)
   - Settings → General → "Collect email addresses"

2. **Slide URL** (*required)
   - Type: Short Answer
   - Validation: Response validation → Regular expression → URL
   - Description: "Google Slides URL (must be accessible by the script's Google account)"

3. **Message to recipients** (optional)
   - Type: Paragraph or Short Answer
   - Description: "Optional message to include in the email"

4. **Email to CC** (optional, create 8 of these)
   - Type: Short Answer
   - Validation: Response validation → Text → Email
   - Description: "CC recipient email address"

> **Note**: See [googleForm.md](googleForm.md) for detailed form setup instructions.

### 2. Deploy the Script

1. Open your Google Form
2. Click the three dots (⋮) → "Script editor"
3. Delete any default code
4. Copy the contents of `Code.js` from this repository
5. Paste into the script editor

### 3. Configure the Script

Update the `CONFIG` object at the top of `Code.js`:

```javascript
const CONFIG = {
  AUTHOR: "Your Name or Organization",
  FORM_URL: "https://forms.google.com/your-form-url-here",
  EMAIL_SUBJECT_PREFIX: "Links from: "
};
```

### 4. Adjust Form Field Mapping (if needed)

If your form field order differs, update the `FORM_FIELDS` constants:

```javascript
const FORM_FIELDS = {
  PRIMARY_EMAIL: 1,      // Column index for primary email
  CC_START: 2,           // First CC field column
  CC_END: 11,            // Last CC field column + 1
  SLIDE_URL: 3,          // Slide URL column
  MESSAGE_TO_RECIPIENTS: 11  // Message field column
};
```

> **Tip**: Form responses are stored in a spreadsheet. Check the column order to verify indices.

### 5. Set Up Form Trigger

1. In Apps Script editor, click on the clock icon (Triggers)
2. Click "+ Add Trigger"
3. Configure:
   - Function: `myFunction`
   - Event source: "From form"
   - Event type: "On form submit"
4. Save and authorize permissions

### 6. Test the Setup

1. Submit a test form with a Google Slides URL you have access to
2. Verify the email is received with extracted links
3. Check the script logs (View → Logs) if issues occur

## Usage

### For End Users

1. Open the Google Form
2. Enter the Google Slides presentation URL
3. (Optional) Add CC recipient emails
4. (Optional) Add a message to recipients
5. Submit the form
6. Check your email for the extracted links

### Important Notes

- **Permissions**: The Google Slides presentation must be accessible by the Google account running the script
- **Share Settings**: If links aren't being extracted, verify the presentation is shared with the script's account
- **Processing Time**: Emails typically arrive within 1-2 minutes of form submission

## Email Output Format

The email you receive will look like this:

```
Subject: Links from: [Your Presentation Title]

This message was sent to you by user@example.com from
[Your Organization]'s Google Slide URL Extraction Tool

[Optional custom message if provided]

Links from [Presentation Title]

Use these links to share in the chat window:

[ 1 ] :: Slide Title Here
Link: https://example.com
Resource Name: https://resource.com

[ 2 ] :: Another Slide
  Speaker Note Links:
    Documentation: https://docs.example.com
```

## Project Structure

```
google-slide-url-helper/
├── Code.js              # Main Google Apps Script
├── appsscript.json      # Apps Script configuration
├── .clasp.json          # Clasp CLI configuration
├── README.md            # This file
└── googleForm.md        # Detailed form setup guide
```

## Technical Details

### Link Detection

The script detects links in:
- **Shapes**: All text boxes and shapes on slides
- **Tables**: Every cell in every table
- **Speaker Notes**: Hidden presenter notes

### Link Format Types

The script handles three link format types:

1. **Plain URLs**: When link text equals the URL
   - Output: `Link: https://example.com`

2. **Email Links**: mailto: protocol links
   - Output: `Email: contact@example.com`

3. **Named Links**: When link has custom display text
   - Output: `Click Here: https://example.com`

### Error Handling

The script includes error handling for:
- Invalid or inaccessible slide URLs
- Missing required fields
- Email sending failures
- Presentation access errors

When errors occur, the user receives an error notification email with troubleshooting guidance.

## Development

### Using Clasp

This project uses [clasp](https://github.com/google/clasp) for version control and deployment.

Install clasp:
```bash
npm install -g @google/clasp
```

Login and setup:
```bash
clasp login
clasp clone <scriptId>
```

Push changes:
```bash
clasp push
```

### Code Optimizations

Recent optimizations include:
- Eliminated global variable pollution
- Added comprehensive JSDoc documentation
- Implemented proper error handling
- Modularized code into focused functions
- Fixed array comparison bugs
- Added configuration constants for easy customization
- Improved string concatenation efficiency
- Enhanced code readability and maintainability

## Troubleshooting

### Links Not Being Extracted

**Problem**: Email arrives but contains no links

**Solutions**:
- Verify the Google Slides URL is correct
- Check that the presentation is shared with the script's Google account
- Ensure links exist in the slides (not just text that looks like URLs)
- Check Apps Script logs for errors

### Email Not Received

**Problem**: Form submitted but no email arrives

**Solutions**:
- Check spam/junk folder
- Verify the form trigger is properly configured
- Check Apps Script execution logs (View → Executions)
- Ensure email addresses are valid
- Verify script has email sending permissions

### Permission Errors

**Problem**: Script fails with authorization errors

**Solutions**:
- Re-authorize the script (Triggers → Re-authorize)
- Ensure you've granted all requested permissions
- Check that the Google account has access to the slides

### Form Trigger Not Firing

**Problem**: Script doesn't run when form is submitted

**Solutions**:
- Verify trigger is set to "On form submit"
- Delete and recreate the trigger
- Check that the trigger function name is `myFunction`
- Review trigger execution logs

## Limitations

- Maximum 8 CC recipients (can be adjusted in form and constants)
- Requires Google Slides presentations (doesn't work with PowerPoint, Keynote, etc.)
- Script execution time limit: 6 minutes (Google Apps Script limit)
- Daily email quota: 100 emails/day for free accounts (Google's limit)

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## License

This project is open source and available for use and modification.

## Credits

Built with Google Apps Script and captured to Git using [clasp](https://github.com/google/clasp).

---

**Questions or Issues?** Check the logs in the Apps Script editor (View → Logs) or review the troubleshooting section above.
