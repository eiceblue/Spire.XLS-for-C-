#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Xls_6.xlsx";
	wstring outputFile = output_path + L"HighlightDuplicateUniqueValues.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Use conditional formatting to highlight duplicate values in range "C2:C10" with IndianRed color.
	XlsConditionalFormats* xcfs = sheet->GetConditionalFormats()->Add();
	xcfs->AddRange(sheet->GetRange(L"C2:C10"));
	IConditionalFormat* format1 = xcfs->AddCondition();
	format1->SetFormatType(ConditionalFormatType::DuplicateValues);
	format1->SetBackColor(Spire::Common::Color::GetIndianRed());

	//Use conditional formatting to highlight unique values in range "C2:C10" with Yellow color.
	XlsConditionalFormats* xcfs1 = sheet->GetConditionalFormats()->Add();
	xcfs1->AddRange(sheet->GetRange(L"C2:C10"));
	IConditionalFormat* format2 = xcfs->AddCondition();
	format2->SetFormatType(ConditionalFormatType::UniqueValues);
	format2->SetBackColor(Spire::Common::Color::GetYellow());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}