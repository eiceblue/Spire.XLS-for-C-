#include "pch.h"
using namespace Spire::Xls;

int main() {
    wstring output_path = OUTPUTPATH;
    wstring outputFile = output_path + L"SetCommentFillColor.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Create Excel font
	ExcelFont* font = workbook->CreateExcelFont();
	font->SetFontName(L"Arial");
	font->SetSize(11);
	font->SetKnownColor(ExcelColors::Orange);

	//Add the comment
	CellRange* range = sheet->GetRange(L"A1");
	range->GetComment()->SetText(L"This is a comment");
	wstring text = range->GetComment()->GetText();
	range->GetComment()->GetRichText()->SetFont(0, (text.size() - 1), font);

	//Set comment Color
	range->GetComment()->GetFill()->SetFillType(ShapeFillType::SolidColor);
	range->GetComment()->GetFill()->SetForeColor(Spire::Common::Color::GetSkyBlue());

	//Set "visible"
	range->GetComment()->SetVisible(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
