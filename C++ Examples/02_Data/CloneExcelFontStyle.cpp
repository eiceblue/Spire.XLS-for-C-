#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CloneExcelFontStyle_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add the text to the Excel sheet cell range A1.
	sheet->GetRange(L"A1")->SetText(L"Text1");

	//Set A1 cell range's CellStyle.
	CellStyle* style = workbook->GetStyles()->Add(L"style");
	style->GetFont()->SetFontName(L"Calibri");
	style->GetFont()->SetColor(Spire::Common::Color::GetRed());
	style->GetFont()->SetSize(12);
	style->GetFont()->SetIsBold(true);
	style->GetFont()->SetIsItalic(true);
	sheet->GetRange(L"A1")->SetCellStyleName(style->GetName());

	//Clone the same style for B2 cell GetRange.
	CellStyle* csOrieign = style->clone();
	sheet->GetRange(L"B2")->SetText(L"Text2");
	sheet->GetRange(L"B2")->SetCellStyleName(csOrieign->GetName());

	//Clone the same style for C3 cell GetRange and then reset the font color for the text.
	CellStyle* csGreen = style->clone();
	csGreen->GetFont()->SetColor(Spire::Common::Color::GetGreen());
	sheet->GetRange(L"C3")->SetText(L"Text3");
	sheet->GetRange(L"C3")->SetCellStyleName(csGreen->GetName());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

