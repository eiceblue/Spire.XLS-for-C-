#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"AutofilterSample.xlsx";
   	wstring output_path = OUTPUTPATH;
   	wstring outputFile = output_path + L"ToCSVWithFilteredValue.csv";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Convert to CSV file with filtered value
	workbook->GetWorksheets()->Get(0)->SaveToFile(outputFile.c_str(), L";", false);
	workbook->Dispose();
}
