# MailMerge

A configurable C# tool for generating personalized Word documents from Excel data and a DOCX template.

## Features

* Generate multiple documents from a single Word template
* Import data from Excel spreadsheets
* Configurable through `Configuration.json`
* Support for relative and absolute paths
* Export imported data to `Table.json`
* Separated components for importing, processing, and document generation

## Usage

1. Prepare a `.docx` template with your merge fields.
2. Prepare an Excel spreadsheet containing the corresponding data.
3. Configure the template, data, and output paths in `Configuration.json`.
4. Run the application.

The application processes the spreadsheet and generates personalized documents from the Word template.

## Project Structure

The application separates the main responsibilities into components such as:

* `ExcelImporter` — imports data from Excel
* `TableWriter` — writes the imported data to JSON
* `WordDocumentGenerator` — generates documents from the template
* `MailMergeApplication` — coordinates the workflow

## Requirements

* Windows
* .NET
* A `.docx` Word template
* An Excel spreadsheet

## License

This project is licensed under the [MIT License](LICENSE).
