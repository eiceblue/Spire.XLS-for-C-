#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"WriteComment.xlsx";
                wstring output_path = OUTPUTPATH;
                wstring outputFile = output_path + L"WriteComment_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Creates font
	ExcelFont* font = workbook->CreateExcelFont();
	font->SetFontName(L"Arial");
	font->SetSize(11);
	font->SetKnownColor(ExcelColors::Orange);
	ExcelFont* fontBlue = workbook->CreateExcelFont();
	fontBlue->SetKnownColor(ExcelColors::LightBlue);
	ExcelFont* fontGreen = workbook->CreateExcelFont();
	fontGreen->SetKnownColor(ExcelColors::LightGreen);
	CellRange* range = sheet->GetRange(L"B11");
	range->SetText(L"Regular comment");
	//Regular comment
	range->GetComment()->SetText(L"Regular comment");
	range->AutoFitColumns();
	range = sheet->GetRange(L"B12");
	range->SetText(L"Rich text comment");
	range->GetRichText()->SetFont(0, 16, font);
	range->AutoFitColumns();
	//Rich text comment
	range->GetComment()->GetRichText()->SetText(L"Rich text comment");
	range->GetComment()->GetRichText()->SetFont(0, 4, fontGreen);
	range->GetComment()->GetRichText()->SetFont(5, 9, fontBlue);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
