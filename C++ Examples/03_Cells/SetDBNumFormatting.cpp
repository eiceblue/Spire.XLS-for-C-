#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring outputFolder = OUTPUTPATH;
	wstring outputFile = outputFolder + L"SetDBNumFormatting_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();
	workbook->CreateEmptySheets(1);

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set value for cells
	sheet->GetRange(L"A1")->SetNumberValue(123);
	sheet->GetRange(L"A2")->SetNumberValue(456);
	sheet->GetRange(L"A3")->SetNumberValue(789);

	//Get the cell range
	CellRange* range = sheet->GetRange(L"A1:A3");

	//Set the DB num format
	range->SetNumberFormat(L"[DBNum2][$-804]General");

	//Auto fit columns
	range->AutoFitColumns();

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}