#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"PivotTableExample.xlsx";
	wstring outputFile = output + L"DisablePivotTableRibbon.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(L"PivotTable");

	XlsPivotTable* pt = dynamic_cast<XlsPivotTable*>(sheet->GetPivotTables()->Get(0));
	//Disable ribbon for this pivot table
	pt->SetEnableWizard(false);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
