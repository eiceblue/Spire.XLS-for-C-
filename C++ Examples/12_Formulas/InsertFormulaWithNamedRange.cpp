#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertFormulaWithNamedRange.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set value
	sheet->GetRange(L"A1")->SetValue(L"1");
	sheet->GetRange(L"A2")->SetValue(L"1");

	//Create a named range
	INamedRange* NamedRange = workbook->GetNameRanges()->Add(L"NewNamedRange");

	NamedRange->SetNameLocal(L"=SUM(A1+A2)");

	//Set the formula
	sheet->GetRange(L"C1")->SetFormula(L"NewNamedRange");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}