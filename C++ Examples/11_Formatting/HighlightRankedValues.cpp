#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"Template_Xls_6.xlsx";
	wstring outputFile = output_path + L"HighlightRankedValues.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Apply conditional formatting to range D2:D10 to highlight the top 2 values.
	XlsConditionalFormats* xcfs = sheet->GetConditionalFormats()->Add();
	xcfs->AddRange(sheet->GetRange(L"D2:D10"));
	IConditionalFormat* format1 = xcfs->AddTopBottomCondition(TopBottomType::Top, 2);
	format1->SetFormatType(ConditionalFormatType::TopBottom);
	format1->SetBackColor(Spire::Common::Color::GetRed());

	//Apply conditional formatting to range E2:E10 to highlight the bottom 2 values.
	XlsConditionalFormats* xcfs1 = sheet->GetConditionalFormats()->Add();
	xcfs1->AddRange(sheet->GetRange(L"E2:E10"));
	IConditionalFormat* format2 = xcfs1->AddTopBottomCondition(TopBottomType::Bottom, 2);
	format2->SetFormatType(ConditionalFormatType::TopBottom);
	format2->SetBackColor(Spire::Common::Color::GetForestGreen());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}