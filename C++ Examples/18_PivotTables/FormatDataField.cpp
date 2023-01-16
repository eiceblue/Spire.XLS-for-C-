#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"FormatDataField.xlsx";
	wstring outputFile = output + L"FormatDataField.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	// Access the PivotTable
	XlsPivotTable* pt = dynamic_cast<XlsPivotTable*>(sheet->GetPivotTables()->Get(0));
	// Access the data field.
	PivotDataField* pivotDataField = pt->GetDataFields()->Get(0);
	// Set data display format
	pivotDataField->SetShowDataAs(PivotFieldFormatType::PercentageOfColumn);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}