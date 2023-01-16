#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"WriteRichText_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	ExcelFont* fontBold = workbook->CreateExcelFont();
	fontBold->SetIsBold(true);

	ExcelFont* fontUnderline = workbook->CreateExcelFont();
	fontUnderline->SetUnderline(FontUnderlineType::Single);

	ExcelFont* fontItalic = workbook->CreateExcelFont();
	fontItalic->SetIsItalic(true);

	ExcelFont* fontColor = workbook->CreateExcelFont();
	fontColor->SetKnownColor(ExcelColors::Green);

	RichText* richText = sheet->GetRange(L"B11")->GetRichText();
	richText->SetText(L"Bold and underlined and italic and colored text.");
	richText->SetFont(0, 3, fontBold);
	richText->SetFont(9, 18, fontUnderline);
	richText->SetFont(24, 29, fontItalic);
	richText->SetFont(35, 41, fontColor);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

