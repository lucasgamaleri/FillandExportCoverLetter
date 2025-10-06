# FillandExportCoverLetter
A simple Macro to fill and export cover letters

Macro Usage Guide: Automated Cover Letter Generation
====================================================

This macro streamlines the creation of personalized cover letters when the core body remains consistent across applications, but company-specific details (e.g., name, address, job title) vary.

Step 1: Setup
-------------
1. Enable Macros in your text editor (e.g., Microsoft Word, LibreOffice Writer).
2. Import the Macro into your system following your editor’s macro installation procedure.

Step 2: Prepare Your Cover Letter Template
------------------------------------------
1. Insert placeholder tags for dynamic fields. These will be automatically replaced via a form prompt when the macro runs.
2. Use the following placeholder format exactly as shown:

   [COMPANY_NAME]   : Full legal name of the company
   [CITY_ADDRESS]   : Company’s main address (street, postal code, city)
   [COUNTRY]        : Country where the company is located
   [POSITION_NAME]  : Job title you are applying for

3. Ensure the rest of the letter body is generic and reusable across applications.

Step 3: Save and Reuse
----------------------
- Save the document as a template file or set it to read-only mode.
- This preserves the original formatting and allows the macro to overwrite placeholders without altering the base structure.
- You can reuse the same template for multiple applications by simply rerunning the macro and inputting new values.
