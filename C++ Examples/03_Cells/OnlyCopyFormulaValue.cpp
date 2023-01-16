#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"CopyOnlyFormulaValue1.xlsx";
	wstring outputFile = outputFolder + L"CopyOnlyFormulaValue_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set the copy option--OnlyCopyFormulaValue
	CopyRangeOptions copyOptions = CopyRangeOptions::OnlyCopyFormulaValue;

	//Copy ranges
	CellRange* sourceRange = sheet->GetRange(L"A6:E6");
	sheet->Copy(sourceRange, sheet->GetRange(L"A8:E8"), copyOptions);

	sourceRange->Copy(sheet->GetRange(L"A10:E10"), copyOptions);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}