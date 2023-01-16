#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"WorksheetSample2.xlsx";
	wstring outputFile = outputFolder + L"SetSummaryRowDirection_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Group Columns
	sheet->GroupByColumns(1, 4, true);

	//Set summary columns to right of details
	sheet->GetPageSetup()->SetIsSummaryColumnRight(true);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}