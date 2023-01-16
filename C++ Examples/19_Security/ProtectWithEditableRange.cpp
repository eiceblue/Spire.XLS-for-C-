#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"ProtectWithEditableRange.xlsx";
	wstring outputFile = output + L"ProtectWithEditableRange.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Define the specified ranges to allow users to edit while sheet is protected
	sheet->AddAllowEditRange(L"EditableRanges", sheet->GetRange(L"B4:E12"));

	//Protect worksheet with a password.
	sheet->XlsWorksheetBase::Protect(L"TestPassword");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}