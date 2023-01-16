#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"templateAz.xlsx";
	wstring outputFile = output_path + L"GetStyleSetStyle.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get "B4" cell
	CellRange* range = sheet->GetRange(L"B4");
	//Get the style of cell
	CellStyle* style = range->GetStyle();
	style->GetFont()->SetFontName(L"Calibri");
	style->GetFont()->SetIsBold(true);
	style->GetFont()->SetSize(15);
	style->GetFont()->SetColor(Spire::Common::Color::GetCornflowerBlue());

	range->SetStyle(style);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
