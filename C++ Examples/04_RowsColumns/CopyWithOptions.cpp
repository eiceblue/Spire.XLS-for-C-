#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"Sample.xlsx";
	wstring outputFile = outputFolder + L"CopyWithOptions_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add a new worksheet as destination sheet
	Worksheet* destinationSheet = workbook->GetWorksheets()->Add(L"DestSheet");

	//Specify a copy range of original sheet
	CellRange* cellRange = sheet->GetRange(L"B2:D4");

	//Copy the specified range to added worksheet and keep original styles and update reference
	workbook->GetWorksheets()->Get(0)->Copy(cellRange, workbook->GetWorksheets()->Get(1), 2, 1, true, true);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}