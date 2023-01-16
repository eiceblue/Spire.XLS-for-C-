#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ApplyMultipleFontsInSingleCell_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Create a font object in workbook, setting the font color, size and type.
	ExcelFont* font1 = workbook->CreateExcelFont();
	font1->SetKnownColor(ExcelColors::LightBlue);
	font1->SetIsBold(true);
	font1->SetSize(10);

	//Create another font object specifying its properties.
	ExcelFont* font2 = workbook->CreateExcelFont();
	font2->SetKnownColor(ExcelColors::Red);
	font2->SetIsBold(true);
	font2->SetIsItalic(true);
	font2->SetFontName(L"Times New Roman");
	font2->SetSize(11);

	//Write a RichText string to the cell 'A1', and set the font for it.
	RichText* richText = sheet->GetRange(L"A5")->GetRichText();
	richText->SetText(L"This document was created with Spire.XLS for C++.");
	richText->SetFont(0, 29, font1);
	richText->SetFont(31, 48, font2);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

