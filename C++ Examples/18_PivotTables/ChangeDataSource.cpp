#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"ChangeDataSource.xlsx";
	wstring outputFile = output + L"ChangeDataSource.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	CellRange* Range = sheet->GetRange(L"A1:C15");
	XlsPivotTable* table = dynamic_cast<XlsPivotTable*>(workbook->GetWorksheets()->Get(1)->GetPivotTables()->Get(0));

	//Change data source
	table->ChangeDataSource(Range);
	table->GetCache()->SetIsRefreshOnLoad(false);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

