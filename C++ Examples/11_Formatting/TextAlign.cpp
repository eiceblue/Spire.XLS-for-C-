#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"TextAlign.xlsx";
	wstring outputFile = output_path + L"TextAlign.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set the vertical alignment to Top
	sheet->GetRange(L"B1:C1")->GetStyle()->SetVerticalAlignment(VerticalAlignType::Top);

	//Set the vertical alignment to Center
	sheet->GetRange(L"B2:C2")->GetStyle()->SetVerticalAlignment(VerticalAlignType::Center);

	//Set the vertical alignment of to Bottom
	sheet->GetRange(L"B3:C3")->GetStyle()->SetVerticalAlignment(VerticalAlignType::Bottom);

	//Set the horizontal alignment to General
	sheet->GetRange(L"B4:C4")->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::General);

	//Set the horizontal alignment of to Left
	sheet->GetRange(L"B5:C5")->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Left);

	//Set the horizontal alignment of to Center
	sheet->GetRange(L"B6:C6")->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Center);

	//Set the horizontal alignment of to Right
	sheet->GetRange(L"B7:C7")->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Right);

	//Set the rotation degree
	sheet->GetRange(L"B8:C8")->GetStyle()->SetRotation(45);

	sheet->GetRange(L"B9:C9")->GetStyle()->SetRotation(90);

	//Set the row height of cell
	sheet->GetRange(L"B8:C9")->SetRowHeight(60);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}