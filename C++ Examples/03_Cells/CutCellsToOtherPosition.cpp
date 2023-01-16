#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"SampleB_2.xlsx";
	wstring outputFile = output_path + L"CutCellsToOtherPosition_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	CellRange* Ori = sheet->GetRange(L"A1:C5");
	CellRange* Dest = sheet->GetRange(L"A26:C30");

	//Copy the range to other position
	sheet->Copy(Ori, Dest, true, true, true);

	//Remove all content in original cells
	for (int i = 0; i < Ori->GetCells()->GetCount(); i++)
	{
		CellRange* cr = Ori->GetCells()->GetItem(i);
		cr->ClearAll();
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

