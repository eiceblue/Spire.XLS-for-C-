#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"ProtectCell.xlsx";
	wstring outputFile = output + L"ProtectCell.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Protect cell
	sheet->GetRange(L"B3")->GetStyle()->SetLocked(true);
	sheet->GetRange(L"C3")->GetStyle()->SetLocked(false);

	sheet->XlsWorksheetBase::Protect(L"TestPassword");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}