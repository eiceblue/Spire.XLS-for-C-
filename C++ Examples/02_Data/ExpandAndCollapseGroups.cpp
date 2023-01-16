#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_3.xlsx";
	wstring outputFile = output_path + L"ExpandAndCollapseGroups_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Expand the grouped rows with ExpandCollapseFlags set to expand parent
	sheet->GetRange(L"A16:G19")->ExpandGroup(GroupByType::ByRows, ExpandCollapseFlags::ExpandParent);

	//Collapse the grouped rows
	sheet->GetRange(L"A10:G12")->CollapseGroup(GroupByType::ByRows);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

