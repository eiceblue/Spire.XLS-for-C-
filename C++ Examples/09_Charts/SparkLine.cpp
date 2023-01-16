#include "pch.h"
using namespace Spire::Xls;

int main() {
                wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SparkLine.xlsx";
	wstring outputFile = output_path + L"SparkLine.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add sparkline
	SparklineGroup* sparklineGroup = sheet->GetSparklineGroups()->AddGroup(SparklineType::Line);
	SparklineCollection* sparklines = sparklineGroup->Add();
	sparklines->Add(sheet->GetRange(L"A2:D2"), sheet->GetRange(L"E2"));
	sparklines->Add(sheet->GetRange(L"A3:D3"), sheet->GetRange(L"E3"));
	sparklines->Add(sheet->GetRange(L"A4:D4"), sheet->GetRange(L"E4"));
	sparklines->Add(sheet->GetRange(L"A5:D5"), sheet->GetRange(L"E5"));
	sparklines->Add(sheet->GetRange(L"A6:D6"), sheet->GetRange(L"E6"));
	sparklines->Add(sheet->GetRange(L"A7:D7"), sheet->GetRange(L"E7"));
	sparklines->Add(sheet->GetRange(L"A8:D8"), sheet->GetRange(L"E8"));
	sparklines->Add(sheet->GetRange(L"A9:D9"), sheet->GetRange(L"E9"));
	sparklines->Add(sheet->GetRange(L"A10:D10"), sheet->GetRange(L"E10"));
	sparklines->Add(sheet->GetRange(L"A11:D11"), sheet->GetRange(L"E11"));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
