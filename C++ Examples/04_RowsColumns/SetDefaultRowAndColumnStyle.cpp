#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring outputFolder = OUTPUTPATH;
	wstring outputFile = outputFolder + L"SetDefaultRowAndColumnStyle_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Create a cell style and set the color
	CellStyle* style = workbook->GetStyles()->Add(L"Mystyle");
	style->SetColor(Spire::Common::Color::GetYellow());

	//Set the default style for the first row and column 
	sheet->SetDefaultRowStyle(1, style);
	sheet->SetDefaultColumnStyle(1, style);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}