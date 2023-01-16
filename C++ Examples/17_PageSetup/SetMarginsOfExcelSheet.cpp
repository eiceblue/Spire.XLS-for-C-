#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"Template_Xls_4.xlsx";
	wstring outputFile = output + L"SetMarginsOfExcelSheet.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the PageSetup object of the first worksheet.
	PageSetup* pageSetup = sheet->GetPageSetup();

	//Set bottom,left,right and top page margins.
	pageSetup->SetBottomMargin(2);
	pageSetup->SetLeftMargin(1);
	pageSetup->SetRightMargin(1);
	pageSetup->SetTopMargin(3);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}