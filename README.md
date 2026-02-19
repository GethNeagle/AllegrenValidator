# AllegrenValidator
An app that validates allergen data for recipes.

Overview

The Allergen Validator App is a desktop-based validation tool designed to check allergen declaration data within structured Excel recipe exports.

It automates structural validation of allergen columns, ensuring:

Mandatory allergens are declared

Values are restricted to Y / M / N

Spelt and Kamut declarations are handled correctly

Output is clearly flagged for review

This tool reduces manual checking time but does not replace compliance review.

How the Application Works

The user selects:

An Input Excel file

An Output file location

The app:

Reads allergen columns

Applies validation rules from allergen_config.json

Flags missing or incorrectly formatted data

Generates a validated output file

An email template is available for structured reporting of issues.

First-Time Setup

Before running validation for the first time:

Launch the application.

Click “Save Allergens and Email Template”.

This will generate:

allergen_config.json

Email template file

Both files are saved in the same directory as the executable.

If these files are missing, the app may not validate correctly.

File Requirements

Input file must:

Be .xlsx

Contain no filters

Contain no formulas

Use static values only

Have correct allergen column headers

Files containing formulas or filters may cause errors or incomplete validation.

Configuration File

The allergen_config.json file controls:

Allergen names

Mandatory status

Accepted values

Example structure:

{
  "Lupin": {
    "mandatory": true,
    "column_name": "Lupin",
    "accepted_values": ["Y", "M", "N"]
  }
}


Editing Notes:

Maintain valid JSON structure.

Keep commas between entries.

Ensure column names match Excel headers exactly.

Restart the app after making changes.

Improper formatting will prevent the app from loading correctly.

Email Template

The email template file:

Is generated via the “Save Allergens and Email Template” button

Can be edited to customise messaging

Must retain placeholders used by the app

Should not be renamed

Limitations

The app validates declared data only.

It does not verify ingredient-level allergen accuracy.

It cannot assess contextual compliance risks.

It relies entirely on the integrity of the source Excel file.

Disclaimer

This tool is designed to support and streamline allergen validation.

It does not replace professional review.

All final responsibility for allergen accuracy remains with the relevant compliance or operational team.

Output quality depends entirely on the quality and completeness of the input data.
