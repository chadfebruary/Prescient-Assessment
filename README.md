# Prescient-Assessment
# DailyMTM Importer

This console application downloads XLS files from the JSE Client Portal, extracts relevant data, and inserts it into a SQL Server database. The application uses Microsoft Excel Interop to read the XLS files and SqlClient for database operations.

## Features

- Downloads XLS files with "2024" in their names from the specified URL.
- Skips downloading files if they already exist in the Downloads folder.
- Extracts data from the downloaded XLS files starting from the 6th row.
- Parses the file name to extract the date.
- Handles potential null values in the columns.
- Inserts the extracted data into a SQL Server table `DailyMTM`.

## Prerequisites

- .NET Framework 4.8
- Microsoft Excel Interop Library
- SQL Server

## Setup

1. Clone the repository:

2. Open the project in Visual Studio.

3. Ensure that you have the Microsoft Excel Interop Library installed. You can install it via NuGet Package Manager:

4. Update the connection string in the `Program.cs` file with your SQL Server details:

## Usage

1. Build the project in Visual Studio.

2. Run the console application. It will:

    - Download XLS files with "2024" in their names from the JSE Client Portal.
    - Skip downloading files if they already exist in the Downloads folder.
    - Extract data from each XLS file starting from the 6th row.
    - Parse the date from the file name.
    - Insert the data into the `DailyMTM` table in your SQL Server database.
