#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring output_path = OUTPUTPATH;
	std::wstring outputFile = output_path + L"HelloWorld.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();
	//Get the first sheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	//Set text for cell range
	sheet->GetRange(L"A1")->SetText(L"Hello World");
	//Set autofit column width 
	sheet->GetRange(L"A1")->AutoFitColumns();

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);
	workbook->Dispose();
}

