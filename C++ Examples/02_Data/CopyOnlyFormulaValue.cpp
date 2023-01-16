#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"CopyOnlyFormulaValue.xlsx";
	wstring outputFile = output_path + L"CopyOnlyFormulaValue_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set the copy option
	CopyRangeOptions copyOptions = CopyRangeOptions::OnlyCopyFormulaValue;

	//Copy range
	sheet->Copy(sheet->GetRange(L"A2:C2"), sheet->GetRange(L"A5:C5"), copyOptions);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

