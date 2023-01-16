#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"WorksheetSample2.xlsx";
	wstring outputFile = output + L"ActivateWorksheet.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the second worksheet from the workbook
	Worksheet* sheet = workbook->GetWorksheets()->Get(1);

	//Activate the sheet
	sheet->Activate();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
