#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ImportDataFromArrayList_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Create an empty worksheet
	workbook->CreateEmptySheets(1);

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Create an ArrayList object
	vector<LPCWSTR> list;

	//Add strings in list
	list.push_back(L"Spire.Doc");
	list.push_back(L"Spire.XLS");
	list.push_back(L"Spire.PDF");
	list.push_back(L"Spire.Presentation");

	//Insert arrary list in worksheet 
	sheet->InsertArray(list, 1, 1, true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

