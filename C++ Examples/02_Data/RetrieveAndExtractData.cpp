#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_3.xlsx";
	wstring outputFile = output_path + L"RetrieveAndExtractData_result.xlsx";

	// Create a new workbook instance and get the first worksheet.
	Workbook* newBook = new Workbook();
	Worksheet* newSheet = newBook->GetWorksheets()->Get(0);

	//Create a new workbook instance and load the sample Excel file.
	Workbook* workbook = new Workbook();
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet.
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Retrieve data and extract it to the first worksheet of the new excel workbook.
	int i = 1;
	int columnCount = sheet->GetColumns()->GetCount();

	for (int i = 0; i < sheet->GetColumns()->GetItem(0)->GetCells()->GetCount(); i++)
	{
		XlsRange* range = sheet->GetColumns()->GetItem(0)->GetCells()->GetItem(i);
		if (wcscmp(range->GetText(), L"teacher") == 0)
		{
			int x = range->GetRow();
			CellRange* sourceRange = sheet->GetRange(range->GetRow(), 1, range->GetRow(), columnCount);
			CellRange* destRange = newSheet->GetRange(i, 1, i, columnCount);
			sheet->Copy(sourceRange, destRange, true);
			i++;
		}
	}

	//Save to file.
	newBook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	newBook->Dispose();
}

