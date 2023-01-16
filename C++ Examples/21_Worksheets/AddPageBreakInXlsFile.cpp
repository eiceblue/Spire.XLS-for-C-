#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"Template_Xls_4.xlsx";
	wstring outputFile = output + L"AddPageBreakInXlsFile.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add page break in Excel file.
	(dynamic_cast<HPageBreaksCollection*>(sheet->GetHPageBreaks()))->Add(sheet->GetRange(L"E4"));
	(dynamic_cast<HPageBreaksCollection*>(sheet->GetHPageBreaks()))->Add(sheet->GetRange(L"C4"));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}