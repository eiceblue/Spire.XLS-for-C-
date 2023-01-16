#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"DataSorting.xls";
	wstring outputFile = output_path + L"DataSorting_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	workbook->GetDataSorter()->GetSortColumns()->Add(2, OrderBy::Ascending);
	workbook->GetDataSorter()->GetSortColumns()->Add(3, OrderBy::Ascending);
	workbook->GetDataSorter()->Sort(sheet->GetRange(L"A1:E19"));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

