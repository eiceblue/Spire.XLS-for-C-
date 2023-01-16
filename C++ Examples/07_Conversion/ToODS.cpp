#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToODS.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ToODS.ods";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Convert to ODS file
	workbook->SaveToFile(outputFile.c_str(), FileFormat::ODS);
	workbook->Dispose();
}
