#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"NamedRanges.xlsx";
	wstring outputFile = output_path + L"NamedRanges_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Creating a named range
	INamedRange* NamedRange = workbook->GetNameRanges()->Add(L"NewNamedRange");
	//Setting the range of the named range
	NamedRange->SetRefersToRange(sheet->GetRange(L"A8:E12"));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

