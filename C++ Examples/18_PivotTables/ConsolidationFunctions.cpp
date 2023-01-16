#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"PivotTableExample.xlsx";
	wstring outputFile = output + L"ConsolidationFunctions.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(L"PivotTable");

	XlsPivotTable* pt = dynamic_cast<XlsPivotTable*>(sheet->GetPivotTables()->Get(0));
	//Apply Average consolidation function to first data field
	pt->GetDataFields()->Get(0)->SetSubtotal(SubtotalTypes::Average);
	//Apply Max consolidation function to second data field
	pt->GetDataFields()->Get(1)->SetSubtotal(SubtotalTypes::Max);
	pt->CalculateData();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}