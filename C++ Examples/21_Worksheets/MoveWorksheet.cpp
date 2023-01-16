#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"WorksheetSample2.xlsx";
	std::wstring outputFile = output_path + L"MoveWorksheet.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Move worksheet
	sheet->MoveWorksheet(2);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}