#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"TextDirection.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Access the "B5" cell from the worksheet
	CellRange* cell = sheet->GetRange(L"B5");

	//Add some value to the "B5" cell
	cell->SetText(L"Hello Spire!");

	//Set the reading order from right to left of the text in the "B5" cell
	cell->GetStyle()->SetReadingOrder(ReadingOrderType::RightToLeft);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}