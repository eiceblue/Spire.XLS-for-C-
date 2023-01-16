#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"Template_Xls_4.xlsx";
	std::wstring outputFile = output_path + L"CopySheetWithinWorkbook.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first and the second worksheets.
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	Worksheet* sheet1 = workbook->GetWorksheets()->Add(L"MySheet");
	CellRange* sourceRange = sheet->GetAllocatedRange();

	//Copy the first worksheet to the second one.
	sheet->Copy(sourceRange, sheet1, sheet->GetFirstRow(), sheet->GetFirstColumn(), true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

