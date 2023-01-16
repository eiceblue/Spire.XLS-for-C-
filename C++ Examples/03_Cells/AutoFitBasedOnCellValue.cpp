#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AutoFitBasedOnCellValue_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set value for B8
	CellRange* cell = sheet->GetRange(L"B8");
	cell->SetText(L"Welcome to Spire.XLS!");

	//Set the cell style
	CellStyle* style = cell->GetStyle();
	style->GetFont()->SetSize(16);
	style->GetFont()->SetIsBold(true);

	//Autofit column width and row height based on cell value
	cell->AutoFitColumns();
	cell->AutoFitRows();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

