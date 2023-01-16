# Spire.XLS for C++ - A C++ Library for Processing Excel Documents

[![Foo](https://i.imgur.com/VwKGaGh.png)](https://www.e-iceblue.com/Introduce/xls-for-CPP.html)

[Product Page](https://www.e-iceblue.com/Introduce/xls-for-CPP.html)  |  [Forum](https://www.e-iceblue.com/forum/spire-xls-f4.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)

[Spire.XLS for C++](https://www.e-iceblue.com/Introduce/xls-for-CPP.html) is a professional **Excel C++ API** that can be used to create, read, write, convert and print Excel files in any type of C++ application without installing Microsoft Office.

This API supports both the old Excel 97-2003 format (.xls) and the new Excel 2007, Excel 2010, Excel 2013, Excel 2016 and Excel 2019 (.xlsx, .xlsb, .xlsm), along with Open Office(.ods) format. It features fast and reliable compared with developing your own spreadsheet manipulation solution or using Microsoft Automation.

### 100% Standalone C++ API

Spire.XLS for C++ is a 100% standalone Excel C++ library without requiring Microsoft Excel or Microsoft Office to be installed on the system.

### Freely Operate Excel Files

- Create/Save/Merge/Split/Get Excel files.
- Encrypt/Decrypt Excel files, add/delete digital signature, tracking changes, lock/unlock worksheets.
- Create/Add/Rename/Edit/Delete/Move worksheets.
- Insert/Modify/Remove hyperlinks.
- Add/Remove/Change/Hide/Show comments in Excel.
- Merge/Unmerge Excel cells, freeze/unfreeze Excel panes, insert/delete Excel rows and columns.
- Add/Read/Calculate/Remove Excel formulas.
- Create/Refresh pivot table.
- Apply/Remove conditional format in Excel.
- Add/Set/Change Excel header and footer.

### Powerful & High Quality Excel File Conversion

- Convert Excel to PDF/Excel to HTML/Excel to XML/Excel to CSV/Excel to Image/Excel to XPS/Excel to SVG
- Convert CSV to Excel/CSV to PDF/Datatable
- Convert selected range of cells to PDF
- Convert XLS to XLSM and maintain macro
- Convert Excel to OpenDocument Spreadsheet(.ods) format
- Save Excel chart sheet to SVG/Image
- Convert HTML to Excel

### Examples
### Create an Excel File in C++

```
#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring output_path = OUTPUTPATH;
	std::wstring outputFile = output_path + L"HelloWorld.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();
	//Get the first sheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	//Set text for cell range
	sheet->GetRange(L"A1")->SetText(L"Hello World");
	//Set autofit column width 
	sheet->GetRange(L"A1")->AutoFitColumns();

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);
	workbook->Dispose();
}
```

### Convert Excel to PDF in C++

```
#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToPDF.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ToPDF.pdf";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	workbook->GetConverterSetting()->SetSheetFitToPage(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	workbook->Dispose();
}
```

### Convert Excel to CSV in C++

```
#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToCSV.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ToCSV.csv";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//convert to CSV file
	sheet->SaveToFile(outputFile.c_str(), L",", Encoding::GetUTF8());
	workbook->Dispose();
}
```

### Convert Excel Worksheet to Image in C++

```
#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SheetToImage.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"SheetToImage.png";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Save to file.
	sheet->ToImage(sheet->GetFirstRow(), sheet->GetFirstColumn(), sheet->GetLastRow(), sheet->GetLastColumn())->Save(outputFile.c_str());
	workbook->Dispose();
}
```

[Product Page](https://www.e-iceblue.com/Introduce/xls-for-CPP.html)  |  [Forum](https://www.e-iceblue.com/forum/spire-xls-f4.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)
