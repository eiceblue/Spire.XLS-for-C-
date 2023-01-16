#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"WorkbookToHTML.xlsx";
	wstring outputFile = outputFolder + L"RemoveRowBasedOnKeyword_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Find the string
	CellRange* cr = sheet->FindString(L"Address", false, false);

	//Delete the row which includes the string
	sheet->DeleteRow(cr->GetRow());

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}