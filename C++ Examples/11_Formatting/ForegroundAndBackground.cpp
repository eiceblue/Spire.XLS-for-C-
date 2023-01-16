#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ForegroundAndBackground.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();
	workbook->SetVersion(ExcelVersion::Version2010);

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Create a new style
	CellStyle* style = workbook->GetStyles()->Add(L"newStyle1");

	//Set filling pattern type
	style->GetInterior()->SetFillPattern(ExcelPatternType::Gradient);

	//Set filling Background color
	style->GetInterior()->GetGradient()->SetBackKnownColor(ExcelColors::Green);

	//Set filling Foreground color
	style->GetInterior()->GetGradient()->SetForeKnownColor(ExcelColors::Yellow);

	style->GetInterior()->GetGradient()->SetGradientStyle(GradientStyleType::From_Center);

	//Apply the style to  "B2" cell
	sheet->GetRange(L"B2")->SetCellStyleName(style->GetName());
	sheet->GetRange(L"B2")->SetText(L"Test");
	sheet->GetRange(L"B2")->SetRowHeight(30);
	sheet->GetRange(L"B2")->SetColumnWidth(50);


	//Create a new style
	style = workbook->GetStyles()->Add(L"newStyle2");

	//Set filling pattern type
	style->GetInterior()->SetFillPattern(ExcelPatternType::Gradient);

	//Set filling Foreground color
	style->GetInterior()->GetGradient()->SetForeKnownColor(ExcelColors::Red);

	//Apply the style to  "B4" cell
	sheet->GetRange(L"B4")->SetCellStyleName(style->GetName());
	sheet->GetRange(L"B4")->SetRowHeight(30);
	sheet->GetRange(L"B4")->SetColumnWidth(60);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}