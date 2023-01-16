#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"AddWorksheet.xlsx";
	wstring outputFile = output + L"AddWorksheet.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Add a new worksheet named AddedSheet
	Worksheet* sheet = workbook->GetWorksheets()->Add(L"AddedSheet");
	sheet->GetRange(L"C5")->SetText(L"This is a new sheet.");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}