#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH; 
	wstring inputFile = fn + L"ShowSubTotals.xlsx";
	wstring outputFile = output + L"ShowSubTotals.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(L"Pivot Table");

	XlsPivotTable* pt = dynamic_cast<XlsPivotTable*>(sheet->GetPivotTables()->Get(0));

	//Show Subtotals
	pt->SetShowSubtotals(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}