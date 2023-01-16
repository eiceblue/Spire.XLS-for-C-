#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"UseArrayFormulas.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	sheet->GetRange(L"A1")->SetNumberValue(1);
	sheet->GetRange(L"A2")->SetNumberValue(2);
	sheet->GetRange(L"A3")->SetNumberValue(3);
	sheet->GetRange(L"B1")->SetNumberValue(4);
	sheet->GetRange(L"B2")->SetNumberValue(5);
	sheet->GetRange(L"B3")->SetNumberValue(6);
	sheet->GetRange(L"C1")->SetNumberValue(7);
	sheet->GetRange(L"C2")->SetNumberValue(8);
	sheet->GetRange(L"C3")->SetNumberValue(9);

	//Write array formula
	sheet->GetRange(L"A5:C6")->SetFormulaArray(L"=LINEST(A1:A3,B1:C3,TRUE,TRUE)");

	//Calculate Formulas
	workbook->CalculateAllValue();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}