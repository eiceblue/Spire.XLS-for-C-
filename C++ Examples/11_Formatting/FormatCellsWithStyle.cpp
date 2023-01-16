#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SampleB_2.xlsx";
	wstring outputFile = output_path + L"FormatCellsWithStyle.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Create a style
	CellStyle* style = workbook->GetStyles()->Add(L"newStyle");
	//Set the shading color
	style->SetColor(Spire::Common::Color::GetDarkGray());
	//Set the font color
	style->GetFont()->SetColor(Spire::Common::Color::GetWhite());
	//Set font name
	style->GetFont()->SetFontName(L"Times New Roman");
	//Set font size
	style->GetFont()->SetSize(12);
	//Set bold for the font
	style->GetFont()->SetIsBold(true);
	//Set text rotation
	style->SetRotation(45);
	//Set alignment
	style->SetHorizontalAlignment(HorizontalAlignType::Center);
	style->SetVerticalAlignment(VerticalAlignType::Center);

	//Set the style for the specific range
	workbook->GetWorksheets()->Get(0)->GetRange(L"A1:J1")->SetCellStyleName(style->GetName());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
