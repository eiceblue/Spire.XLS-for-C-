#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConversionSample1.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"Image";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Specify Cell Ranges and Save to certain Image formats
	sheet->ToImage(1, 1, 7, 5)->Save((outputFile + L"SpecificCellsToImage.png").c_str(), ImageFormat::GetPng());
	sheet->ToImage(8, 1, 15, 5)->Save((outputFile + L"SpecificCellsToImage.jpg").c_str(), ImageFormat::GetJpeg());
	sheet->ToImage(17, 1, 23, 5)->Save((outputFile + L"SpecificCellsToImage.bmp").c_str(), ImageFormat::GetBmp());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
