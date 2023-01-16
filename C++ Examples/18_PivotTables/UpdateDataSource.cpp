#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"PivotTableExample.xlsx";
	wstring outputFile = output + L"UpdateDataSource.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* data = workbook->GetWorksheets()->Get(L"Data");
	CellRange* a2 = data->Get(L"A2");

	a2->SetText(L"NewValue");
	data->Get(L"D2")->SetNumberValue(28000);

	//Get the sheet in which the pivot table is located
	Worksheet* sheet = workbook->GetWorksheets()->Get(L"PivotTable");

	XlsPivotTable* pt = dynamic_cast<XlsPivotTable*>(sheet->GetPivotTables()->Get(0));
	//Refresh and calculate
	pt->GetCache()->SetIsRefreshOnLoad(true);
	pt->CalculateData();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}