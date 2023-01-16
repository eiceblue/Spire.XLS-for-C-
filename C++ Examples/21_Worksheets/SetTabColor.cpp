#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"SetTabColor.xlsx";
	std::wstring outputFile = output_path + L"SetTabColor.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Set the tab color of first sheet to be red 
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	sheet->SetTabColor(Spire::Common::Color::GetRed());

	//Set the tab color of first sheet to be green 
	sheet = workbook->GetWorksheets()->Get(1);
	sheet->SetTabColor(Spire::Common::Color::GetGreen());

	//Set the tab color of first sheet to be blue 
	sheet = workbook->GetWorksheets()->Get(2);
	sheet->SetTabColor(Spire::Common::Color::GetLightBlue());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}