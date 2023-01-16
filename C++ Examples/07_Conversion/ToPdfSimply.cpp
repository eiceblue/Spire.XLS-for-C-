#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToPDF.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ToPdfSimply.pdf";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Convert excel to pdf
	workbook->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	workbook->Dispose();
}
