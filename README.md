# MailMerge

A simple and configurable C# tool for generating personalized documents from a Microsoft Word template and structured data.

The project is designed to automate repetitive document generation while keeping the template, input data, and output configuration separate. Configuration is handled through a JSON file, making the tool easy to adapt without modifying the source code.

## Features

* Uses a `.docx` template as the document source
* Generates personalized documents automatically
* Configurable through `Configuration.json`
* Supports relative and absolute file paths
* Separates configuration, processing logic, and output handling
* Designed to be simple, lightweight, and easy to extend

## Use Cases

Useful for generating certificates, letters, invoices, notices, reports, forms, and other documents where the same template needs to be populated with different data.

## Configuration

The application uses a JSON configuration file to define the template and output locations, keeping environment-specific settings outside the application code.

## Requirements

* Windows
* .NET
* Microsoft Word-compatible `.docx` templates

## Status

This project is actively developed and may evolve as additional mail-merge features are added.
