#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"CreateTable.xlsx";
	std::wstring outputFile = output_path + L"SetExcelCalculationMode.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Set excel calculation mode as Manual
	workbook->SetCalculationMode(ExcelCalculationMode::Manual);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}