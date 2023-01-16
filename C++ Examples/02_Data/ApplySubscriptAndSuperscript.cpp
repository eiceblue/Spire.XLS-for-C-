#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ApplySubscriptAndSuperscript_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	sheet->GetRange(L"B2")->SetText(L"This is an example of Subscript:");
	sheet->GetRange(L"D2")->SetText(L"This is an example of Superscript:");

	//Set the rtf value of "B3" to "R100-0.06".
	CellRange* range = sheet->GetRange(L"B3");
	range->GetRichText()->SetText(L"R100-0.06");

	//Create a font. Set the IsSubscript property of the font to "true".
	ExcelFont* font = workbook->CreateExcelFont();
	font->SetIsSubscript(true);
	font->SetColor(Spire::Common::Color::GetGreen());

	//Set font for specified range of the text in "B3".
	range->GetRichText()->SetFont(4, 8, font);

	//Set the rtf value of "D3" to "a2 + b2 = c2".
	range = sheet->GetRange(L"D3");
	range->GetRichText()->SetText(L"a2 + b2 = c2");

	//Create a font. Set the IsSuperscript property of the font to "true".
	font = workbook->CreateExcelFont();
	font->SetIsSuperscript(true);

	//Set font for specified range of the text in "D3".
	range->GetRichText()->SetFont(1, 1, font);
	range->GetRichText()->SetFont(6, 6, font);
	range->GetRichText()->SetFont(11, 11, font);

	sheet->GetAllocatedRange()->AutoFitColumns();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

