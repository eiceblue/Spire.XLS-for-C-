#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ColorsAndPalette.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Adding Orchid color to the palette at 60th index
	workbook->ChangePaletteColor(Spire::Common::Color::GetOrchid(), 60);

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	CellRange* cell = sheet->GetRange(L"B2");
	cell->SetText(L"Welcome to use Spire.XLS");

	//Set the Orchid (custom) color to the font
	cell->GetStyle()->GetFont()->SetColor(Spire::Common::Color::GetOrchid());
	cell->GetStyle()->GetFont()->SetSize(20);
	cell->AutoFitColumns();
	cell->AutoFitRows();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}