#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile_1 = input_path + L"ReadImages.xlsx";
	std::wstring inputFile_2 = input_path + L"sample.xlsx";
	std::wstring outputFile = output_path + L"CopyWorksheet.xlsx";

	//Create a workbook
	Workbook* sourceWorkbook = new Workbook();

	//Load the Excel document from disk
	sourceWorkbook->LoadFromFile(inputFile_1.c_str());

	//Get the first worksheet
	Worksheet* srcWorksheet = sourceWorkbook->GetWorksheets()->Get(0);

	//Create a workbook
	Workbook* targetWorkbook = new Workbook();

	//Load the target Excel document from disk
	targetWorkbook->LoadFromFile(inputFile_2.c_str());

	//Add a new worksheet
	Worksheet* targetWorksheet = targetWorkbook->GetWorksheets()->Add(L"added");

	//Copy the first worksheet of source Excel document to the new added worksheet of target Excel document
	targetWorksheet->CopyFrom(srcWorksheet);

	//Save the document
	targetWorkbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	sourceWorkbook->Dispose();
}