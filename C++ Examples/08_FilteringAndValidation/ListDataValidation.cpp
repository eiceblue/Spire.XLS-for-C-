#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DataValidation.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ListDataValidation_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set text for cells 
	sheet->GetRange(L"A7")->SetText(L"Beijing");
	sheet->GetRange(L"A8")->SetText(L"New York");
	sheet->GetRange(L"A9")->SetText(L"Denver");
	sheet->GetRange(L"A10")->SetText(L"Paris");

	//Set data validation for cell
	CellRange* range = sheet->GetRange(L"D10");
	range->GetDataValidation()->SetShowError(true);
	range->GetDataValidation()->SetAlertStyle(AlertStyleType::Stop);
	range->GetDataValidation()->SetErrorTitle(L"Error");
	range->GetDataValidation()->SetErrorMessage(L"Please select a city from the list");
	range->GetDataValidation()->SetDataRange(sheet->GetRange(L"A7:A10"));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
