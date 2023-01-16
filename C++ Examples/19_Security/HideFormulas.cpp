#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"FormulasSample.xlsx";
	wstring outputFile = output + L"HideFormulas.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Hide the formulas in the used range
	sheet->GetAllocatedRange()->SetIsFormulaHidden(true);

	//Protect the worksheet with password
	sheet->XlsWorksheetBase::Protect(L"e-iceblue");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}