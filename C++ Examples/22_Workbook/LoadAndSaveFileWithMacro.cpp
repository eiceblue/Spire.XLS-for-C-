#include "pch.h"
using namespace Spire::Xls;

int main(){
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"MacroSample.xls";
	std::wstring outputFile = output_path + L"LoadAndSaveFileWithMacro.xls";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	sheet->GetRange(L"A5")->SetText(L"This is a simple test!");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version97to2003);
	workbook->Dispose();
}