#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ConditionalFormatting.xlsx";
	wstring outputFile = output_path + L"SimpleConditionalFormatting.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	sheet->GetAllocatedRange()->SetRowHeight(15);
	sheet->GetAllocatedRange()->SetColumnWidth(16);

	//Create conditional formatting rule
	XlsConditionalFormats* xcfs1 = sheet->GetConditionalFormats()->Add();
	xcfs1->AddRange(sheet->GetRange(L"A1:D1"));
	IConditionalFormat* cf1 = xcfs1->AddCondition();
	cf1->SetFormatType(ConditionalFormatType::CellValue);
	cf1->SetFirstFormula(L"150");
	cf1->SetOperator(ComparisonOperatorType::Greater);
	cf1->SetFontColor(Spire::Common::Color::GetRed());
	cf1->SetBackColor(Spire::Common::Color::GetLightBlue());

	XlsConditionalFormats* xcfs2 = sheet->GetConditionalFormats()->Add();
	xcfs2->AddRange(sheet->GetRange(L"A2:D2"));
	IConditionalFormat* cf2 = xcfs2->AddCondition();
	cf2->SetFormatType(ConditionalFormatType::CellValue);
	cf2->SetFirstFormula(L"300");
	cf2->SetOperator(ComparisonOperatorType::Less);
	//Set border color
	cf2->SetLeftBorderColor(Spire::Common::Color::GetPink());
	cf2->SetRightBorderColor(Spire::Common::Color::GetPink());
	cf2->SetTopBorderColor(Spire::Common::Color::GetDeepSkyBlue());
	cf2->SetBottomBorderColor(Spire::Common::Color::GetDeepSkyBlue());
	cf2->SetLeftBorderStyle(LineStyleType::Medium);
	cf2->SetRightBorderStyle(LineStyleType::Thick);
	cf2->SetTopBorderStyle(LineStyleType::Double);
	cf2->SetBottomBorderStyle(LineStyleType::Double);

	//Add data bars
	XlsConditionalFormats* xcfs3 = sheet->GetConditionalFormats()->Add();
	xcfs3->AddRange(sheet->GetRange(L"A3:D3"));
	IConditionalFormat* cf3 = xcfs3->AddCondition();
	cf3->SetFormatType(ConditionalFormatType::DataBar);
	cf3->GetDataBar()->SetBarColor(Spire::Common::Color::GetCadetBlue());

	//Add icon sets
	XlsConditionalFormats* xcfs4 = sheet->GetConditionalFormats()->Add();
	xcfs4->AddRange(sheet->GetRange(L"A4:D4"));
	IConditionalFormat* cf4 = xcfs4->AddCondition();
	cf4->SetFormatType(ConditionalFormatType::IconSet);
	cf4->GetIconSet()->SetIconSetType(IconSetType::ThreeTrafficLights1);

	//Add color scales
	XlsConditionalFormats* xcfs5 = sheet->GetConditionalFormats()->Add();
	xcfs5->AddRange(sheet->GetRange(L"A5:D5"));
	IConditionalFormat* cf5 = xcfs5->AddCondition();
	cf5->SetFormatType(ConditionalFormatType::ColorScale);

	//Highlight duplicate values in range "A6:D6" with BurlyWood color
	XlsConditionalFormats* xcfs6 = sheet->GetConditionalFormats()->Add();
	xcfs6->AddRange(sheet->GetRange(L"A6:D6"));
	IConditionalFormat* cf6 = xcfs6->AddCondition();
	cf6->SetFormatType(ConditionalFormatType::DuplicateValues);
	cf6->SetBackColor(Spire::Common::Color::GetBurlyWood());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}