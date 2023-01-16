#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"SetDataValidationOnSeparateSheet.xlsx";
	wstring outputFile = output_path + L"SetDataValidationOnSeparateSheet_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//This is the first sheet
	Worksheet* sheet1 = workbook->GetWorksheets()->Get(0);

	sheet1->GetRange(L"B10")->SetText(L"Here is a data validation example.");
	//This is the second sheet
	Worksheet* sheet2 = workbook->GetWorksheets()->Get(1);
	//The property is to enable the data can be from different sheet.
	sheet2->GetParentWorkbook()->SetAllow3DRangesInDataValidation(true);
	sheet1->GetRange(L"B11")->GetDataValidation()->SetDataRange(sheet2->GetRange(L"A1:A7"));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

