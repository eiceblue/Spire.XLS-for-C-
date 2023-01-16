#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"AutofilterBlank.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"AutofilterBlank_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Match the blank data
	(dynamic_cast<AutoFiltersCollection*>(sheet->GetAutoFilters()))->MatchBlanks(0);

	//Filter
	(dynamic_cast<AutoFiltersCollection*>(sheet->GetAutoFilters()))->Filter();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
