#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output = OUTPUTPATH;
	wstring outputFile = output + L"SetExcelPaperSize.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set the paper size of the worksheet as A4 paper.
	sheet->GetPageSetup()->SetPaperSize(PaperSizeType::PaperA4);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}