#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"FontStyles.xlsx";
	wstring outputFile = output_path + L"FontStyles.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set font style
	sheet->GetRange(L"B1")->GetStyle()->GetFont()->SetFontName(L"Comic Sans MS");
	sheet->GetRange(L"B2:D2")->GetStyle()->GetFont()->SetFontName(L"Corbel");
	sheet->GetRange(L"B3:D7")->GetStyle()->GetFont()->SetFontName(L"Aleo");

	//Set font size
	sheet->GetRange(L"B1")->GetStyle()->GetFont()->SetSize(45);
	sheet->GetRange(L"B2:D3")->GetStyle()->GetFont()->SetSize(25);
	sheet->GetRange(L"B3:D7")->GetStyle()->GetFont()->SetSize(12);

	//Set excel cell data to be bold
	sheet->GetRange(L"B2:D2")->GetStyle()->GetFont()->SetIsBold(true);

	//Set excel cell data to be underline
	sheet->GetRange(L"B3:B7")->GetStyle()->GetFont()->SetUnderline(FontUnderlineType::Single);

	//set excel cell data color
	sheet->GetRange(L"B1")->GetStyle()->GetFont()->SetColor(Spire::Common::Color::GetCornflowerBlue());
	sheet->GetRange(L"B2:D2")->GetStyle()->GetFont()->SetColor(Spire::Common::Color::GetCadetBlue());
	sheet->GetRange(L"B3:D7")->GetStyle()->GetFont()->SetColor(Spire::Common::Color::GetFirebrick());

	//set excel cell data to be italic
	sheet->GetRange(L"B3:D7")->GetStyle()->GetFont()->SetIsItalic(true);

	//Add strikethrough
	sheet->GetRange(L"D3")->GetStyle()->GetFont()->SetIsStrikethrough(true);
	sheet->GetRange(L"D7")->GetStyle()->GetFont()->SetIsStrikethrough(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}