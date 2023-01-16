#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ExcelSample_N1.xlsx";
	wstring outputFile = output_path + L"SetFormulaWithNamedRange.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Create a named range
	INamedRange* NamedRange = workbook->GetNameRanges()->Add(L"MyNamedRange");
	//Refers to range
	NamedRange->SetRefersToRange(sheet->GetRange(L"B10:B12"));

	//Set the formula of range to named range
	sheet->GetRange(L"B13")->SetFormula(L"=SUM(MyNamedRange)");

	//Set value of ranges
	sheet->GetRange(L"B10")->SetNumberValue(10);
	sheet->GetRange(L"B11")->SetNumberValue(20);
	sheet->GetRange(L"B12")->SetNumberValue(30);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();

	ifstream f(outputFile.c_str());
}