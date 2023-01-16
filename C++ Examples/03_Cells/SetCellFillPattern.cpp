#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"CommonTemplate.xlsx";
	wstring outputFile = outputFolder + L"SetCellFillPattern_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set cell color
	sheet->GetRange(L"B7:F7")->GetStyle()->SetColor(Spire::Common::Color::GetYellow());
	//Set cell fill pattern
	sheet->GetRange(L"B8:F8")->GetStyle()->SetFillPattern(ExcelPatternType::Percent125Gray);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}