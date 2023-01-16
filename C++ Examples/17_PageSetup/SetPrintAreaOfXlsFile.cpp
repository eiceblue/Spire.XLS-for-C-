#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"Template_Xls_4.xlsx";
	wstring outputFile = output + L"SetPrintAreaOfXlsFile.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the reference of the PageSetup of the worksheet.
	PageSetup* pageSetup = sheet->GetPageSetup();

	//Specify the cells range of the print area.
	pageSetup->SetPrintArea(L"A1:E5");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}