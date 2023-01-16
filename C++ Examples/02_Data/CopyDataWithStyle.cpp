#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CopyDataWithStyle_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set the values for some cells.
	CellRange* cells = sheet->GetRange(L"A1:J50");
	for (int i = 1; i <= 10; i++)
	{
		for (int j = 1; j <= 8; j++)
		{
			cells->Get(i, j)->SetText((to_wstring(i - 1) + L"," + to_wstring(j - 1)).c_str());
		}
	}
	//Get a source range (A1:D3).
	CellRange* srcRange = sheet->GetRange(L"A1:D3");

	//Create a style object.
	CellStyle* style = workbook->GetStyles()->Add(L"style");

	//Specify the font attribute.
	style->GetFont()->SetFontName(L"Calibri");

	//Specify the shading color.
	style->GetFont()->SetColor(Spire::Common::Color::GetRed());

	//Specify the border attributes.
	style->GetBorders()->Get(BordersLineType::EdgeTop)->SetLineStyle(LineStyleType::Thin);
	style->GetBorders()->Get(BordersLineType::EdgeTop)->SetColor(Spire::Common::Color::GetBlue());
	style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetLineStyle(LineStyleType::Thin);
	style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetColor(Spire::Common::Color::GetBlue());
	style->GetBorders()->Get(BordersLineType::EdgeTop)->SetLineStyle(LineStyleType::Thin);
	style->GetBorders()->Get(BordersLineType::EdgeTop)->SetColor(Spire::Common::Color::GetBlue());
	style->GetBorders()->Get(BordersLineType::EdgeRight)->SetLineStyle(LineStyleType::Thin);
	style->GetBorders()->Get(BordersLineType::EdgeRight)->SetColor(Spire::Common::Color::GetBlue());
	srcRange->SetCellStyleName(style->GetName());

	//Set the destination range
	CellRange* destRange = sheet->GetRange(L"A12:D14");

	//Copy the range data with style
	srcRange->Copy(destRange, true, true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

