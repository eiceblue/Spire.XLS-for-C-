#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"UsePredefinedStyles.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Create a new style
	CellStyle* style = workbook->GetStyles()->Add(L"newStyle");
	style->GetFont()->SetFontName(L"Calibri");
	style->GetFont()->SetIsBold(true);
	style->GetFont()->SetSize(15);
	style->GetFont()->SetColor(Spire::Common::Color::GetCornflowerBlue());

	//Get "B5" cell
	CellRange* range = sheet->GetRange(L"B5");
	range->SetText(L"Welcome to use Spire.XLS");
	range->SetCellStyleName(style->GetName());
	range->AutoFitColumns();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}