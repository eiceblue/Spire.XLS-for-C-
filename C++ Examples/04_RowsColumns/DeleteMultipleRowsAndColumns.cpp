#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"CommonTemplate1.xlsx";
	wstring outputFile = outputFolder + L"DeleteMultipleRowsAndColumns_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Delete 4 rows starting with the fifth row
	sheet->DeleteRow(5, 4);

	//Delete 2 columns starting with the second column
	sheet->DeleteColumn(2, 2);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}