#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Xls_6.xlsx";
	wstring outputFile = output_path +L"HighlightAverageValues.xlsx";

	//Create a workbook
	Workbook*  workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add conditional format.
	XlsConditionalFormats* format1 = sheet->GetConditionalFormats()->Add();
	//Set the cell range to apply the formatting.
	format1->AddRange(sheet->GetRange(L"E2:E10"));
	//Add below average condition.
	IConditionalFormat* cf1 = format1->AddAverageCondition(AverageType::Below);
	//Highlight cells below average values.
	cf1->SetBackColor(Spire::Common::Color::GetSkyBlue());

	//Add conditional format.
	XlsConditionalFormats* format2 = sheet->GetConditionalFormats()->Add();
	//Set the cell range to apply the formatting.
	format2->AddRange(sheet->GetRange(L"E2:E10"));
	//Add above average condition.
	IConditionalFormat* cf2 = format1->AddAverageCondition(AverageType::Above);
	//Highlight cells above average values.
	cf2->SetBackColor(Spire::Common::Color::GetOrange());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}