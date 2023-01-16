#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_1.xlsx";
	wstring outputFile = output_path + L"DetectMergedCells_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the merged cell ranges in the first worksheet and put them into a CellRange array.
	Spire::Common::IList<XlsRange>* range = sheet->GetMergedCells();

	//Traverse through the array and unmerge the merged cells.
	//for (auto cell : range)
	for (int i = 0; i < range->GetCount(); i++)
	{
		XlsRange* cell = range->GetItem(i);
		cell->UnMerge();
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

