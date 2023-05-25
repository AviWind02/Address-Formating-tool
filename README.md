# Address Format

Address Format is a C# console application that reads data from a CSV file containing company address information and generates a Word document with a formatted table of the addresses.

## Prerequisites

- .NET Framework (version 4.7.2)
- Microsoft Office (version 2013 or later)

## Getting Started

1. Clone the repository or download the source code.
2. Open the solution in your preferred C# IDE.
3. Build the solution to ensure all dependencies are resolved.

## Usage

1. Prepare a CSV file containing the company address information.
2. Update the `filePath` variable in the `Main` method of the `Program` class to point to your CSV file.
3. Run the application.
4. The application will read the CSV file, store the address information in memory, and generate a Word document.
5. The generated Word document will be saved in the specified file path.

## Customization

You can customize the following aspects of the generated Word document:

- Page orientation (portrait or landscape): Modify the `Orientation` property in the `CreateWordDocument` method of the `creatingwordfile` class.
- Table layout: Adjust the table dimensions, cell widths, and cell heights in the `CreateWordDocument` method of the `creatingwordfile` class.
- Cell content: Modify the cell content concatenation in the `CreateWordDocument` method of the `creatingwordfile` class to meet your specific requirements.

## License

This project is licensed under the [MIT License](LICENSE).

## Acknowledgments

- [Microsoft Office Interop API](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word)
- [CSV file format](https://en.wikipedia.org/wiki/Comma-separated_values)

## Files used in testing

- [Test.CSV](https://github.com/AviWind02/Address-Formating-tool/files/11569480/Test.CSV)
- [CompanyInfo.docx](https://github.com/AviWind02/Address-Formating-tool/files/11569458/CompanyInfo.docx)

