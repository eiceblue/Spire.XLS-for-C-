#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Xls_6.xlsx";
	wstring outputFile = output_path + L"ConditionallyFormatDate.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Highlight cells that contain a date occurring in the last 7 days.
	XlsConditionalFormats* xcfs = sheet->GetConditionalFormats()->Add();
	xcfs->AddRange(sheet->GetAllocatedRange());
	IConditionalFormat* conditionalFormat = xcfs->AddTimePeriodCondition(TimePeriodType::Last7Days);
	conditionalFormat->SetBackColor(Spire::Common::Color::GetOrange());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}