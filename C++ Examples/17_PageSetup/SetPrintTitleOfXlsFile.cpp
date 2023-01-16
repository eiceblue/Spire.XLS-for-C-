#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"Template_Xls_4.xlsx";
	wstring outputFile = output + L"SetPrintTitleOfXlsFile.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	PageSetup* pageSetup = sheet->GetPageSetup();

	//Define column numbers A & B as title columns.
	pageSetup->SetPrintTitleColumns(L"$A:$B");

	//Defining row numbers 1 & 2 as title rows.
	pageSetup->SetPrintTitleRows(L"$1:$2");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}