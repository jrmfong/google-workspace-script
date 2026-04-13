# Gift Aid Declaration Google Form Automation

This project contains a Google Apps Script that runs when a Google Form is submitted.

It automates the processing of Gift Aid declarations by generating a PDF copy of the declaration when the respondent accepts the required terms, storing the PDF in Google Drive, emailing a copy to the respondent, and notifying the Finance Team.

If the respondent does not accept the required terms, the script does not generate the Gift Aid Declaration PDF. Instead, it sends an email notification to both the respondent and the Finance Team.

## Features

- Triggered automatically on Google Form submission
- Checks whether the respondent has accepted the required Gift Aid terms
- Generates a PDF copy of the Gift Aid Declaration when the submission is valid
- Stores the generated PDF in a designated Google Drive folder
- Emails the PDF to the respondent for their records
- Sends an email notification to the Finance Team
- Sends a rejection notification if the required terms are not accepted

## How It Works

When a respondent submits the Google Form:

1. The script reads the form responses.
2. It checks whether the respondent agreed to the required Gift Aid declarations.
3. If both required terms are accepted:
   - a PDF copy of the Gift Aid Declaration is generated
   - the PDF is stored in the Finance Team shared drive
   - the respondent receives an email with the PDF attached
   - the Finance Team receives an email notification
4. If the required terms are not accepted:
   - no PDF is generated
   - the respondent receives an email notification
   - the Finance Team also receives an email notification

## Repository Contents

This repository contains:

- the Google Apps Script source code
- a template PDF file for creating Google Doc template for generating the Gift Aid Declaration
- this `README.md`

## Requirements

- Google Form
- Google Apps Script
- Google Drive access
- Gmail access for sending email notifications
- A Google Docs template file for the Gift Aid Declaration

## Google Docs Template

Google Docs template file is not stored in this repository.

To use this project:

1. Review the sample PDF in `GiftAidDeclarationTemplate.pdf`
2. Create your own Google Docs template file
3. Copy the wording and layout based on the sample
4. Update the script with the Google Docs template file ID

## Configuration

Update the constants in the script to match your environment, including:

- template Google Docs file ID
- target Google Drive folder ID
- organisation name
- Finance Team email address

## Deployment

1. It is recommended to use a dedicated account to create the Google Form and deploy this script.
2. Create the Google Form with reference to the sample Gift Aid declaration for multiple donations available here:  
   https://www.gov.uk/claim-gift-aid/gift-aid-declarations
3. Create the template Google Docs file with reference to the template pdf provided in this repository. Make sure the deployment account has read access to the Google Doc file.
4. Make sure the target storage folder exists in Google Drive. Create it if necessary. Make sure the deployment account has read and write access to it.
5. Update the script constants as required.
6. Attach the Apps Script project to the Google Form and configure the form submission trigger.

## Notes

- Make sure the Google Form question titles exactly match the field names expected in the script.
- Make sure the template placeholders in the Google Docs file match the placeholders used in the script.
- Test the script with sample submissions before using it in production.
- Review access permissions carefully before deployment.

## Reference

Gift Aid declaration guidance and sample wording:  
https://www.gov.uk/claim-gift-aid/gift-aid-declarations
