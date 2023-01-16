#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Subtotal.xlsx";
	wstring outputFile = output_path + L"Subtotal_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Select data range
	CellRange* range = sheet->GetRange(L"A1:B18");
	//Subtotal selected data
	sheet->Subtotal(range, 0, { 1 }, SubtotalTypes::Sum, true, false, true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

