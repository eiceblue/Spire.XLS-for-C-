#include "pch.h"
using namespace Spire::Xls;

int main() {;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path  + L"ApplyIconSetsToCellRange.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Insert data to cell range from A1 to C4.
	sheet->GetRange(L"A1")->SetNumberValue(582);
	sheet->GetRange(L"A2")->SetNumberValue(234);
	sheet->GetRange(L"A3")->SetNumberValue(314);
	sheet->GetRange(L"A4")->SetNumberValue(50);
	sheet->GetRange(L"B1")->SetNumberValue(150);
	sheet->GetRange(L"B2")->SetNumberValue(894);
	sheet->GetRange(L"B3")->SetNumberValue(560);
	sheet->GetRange(L"B4")->SetNumberValue(900);
	sheet->GetRange(L"C1")->SetNumberValue(134);
	sheet->GetRange(L"C2")->SetNumberValue(700);
	sheet->GetRange(L"C3")->SetNumberValue(920);
	sheet->GetRange(L"C4")->SetNumberValue(450);
	sheet->GetAllocatedRange()->SetRowHeight(15);
	sheet->GetAllocatedRange()->SetColumnWidth(17);

	//Add icon sets.
	XlsConditionalFormats* xcfs = sheet->GetConditionalFormats()->Add();
	xcfs->AddRange(sheet->GetAllocatedRange());
	IConditionalFormat* format = xcfs->AddCondition();
	format->SetFormatType(ConditionalFormatType::IconSet);
	format->GetIconSet()->SetIconSetType(IconSetType::ThreeTrafficLights1);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}