#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"InsertRowsAndColumns.xls";
	wstring outputFile = outputFolder + L"InsertRowsAndColumns_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Insert a row into the worksheet 
	sheet->InsertRow(2);
	//Insert a column into the worksheet 
	sheet->InsertColumn(2);
	//Insert multiple rows into the worksheet
	sheet->InsertRow(5, 2);
	//Insert multiple columns into the worksheet
	sheet->InsertColumn(5, 2);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}