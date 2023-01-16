#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring data_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = data_path + L"templateAz2.xlsx";
	std::wstring outputFile = output_path + L"OpenExistingFile.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();
	workbook->LoadFromFile(inputFile.c_str());

	//Add a new sheet, named MySheet
	Worksheet* sheet = workbook->GetWorksheets()->Add(L"MySheet");

	//Get the reference of "A1" cell from the cells collection of a worksheet
	sheet->GetRange(L"A1")->SetText(L"Hello World");

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);
	workbook->Dispose();
}

