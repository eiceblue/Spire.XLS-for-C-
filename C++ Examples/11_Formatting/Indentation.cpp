#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Indentation.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Access the "B5" cell from the worksheet
	CellRange* cell = sheet->GetRange(L"B5");

	//Add some value to the "B5" cell
	cell->SetText(L"Hello Spire!");

	//Set the indentation level of the text (inside the cell) to 2
	cell->GetStyle()->SetIndentLevel(2);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}