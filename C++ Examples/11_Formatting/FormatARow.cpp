#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"FormatARow.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Create a new style
	CellStyle* style = workbook->GetStyles()->Add(L"newStyle");

	//Set the vertical alignment of the text
	style->SetVerticalAlignment(VerticalAlignType::Center);

	//Set the horizontal alignment of the text
	style->SetHorizontalAlignment(HorizontalAlignType::Center);

	//Set the font color of the text
	style->GetFont()->SetColor(Spire::Common::Color::GetBlue());

	//Shrink the text to fit in the cell
	style->SetShrinkToFit(true);

	//Set the bottom border color of the cell to OrangeRed
	style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetColor(Spire::Common::Color::GetOrangeRed());

	//Set the bottom border type of the cell to Dotted
	style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetLineStyle(LineStyleType::Dotted);

	//Apply the style to the second row
	sheet->GetRows()->GetItem(1)->SetCellStyleName(style->GetName());

	sheet->GetRows()->GetItem(1)->SetText(L"Test");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
