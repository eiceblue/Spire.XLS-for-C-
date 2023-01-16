#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"CreateTable.xlsx";
	wstring outputFile = output_path + L"ScopedNamedRange.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add range name
	INamedRange* namedRange = sheet->GetNames()->Add(L"Range1");

	//Define the range
	namedRange->SetRefersToRange(sheet->GetRange(L"A1:D19"));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();

	ifstream f(outputFile.c_str());
}