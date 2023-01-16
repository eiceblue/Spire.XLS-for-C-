#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartToImage.xlsx";
	wstring outputFile = output_path + L"ChartToImage.png";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	Image* image = workbook->SaveChartAsImage(workbook->GetWorksheets()->Get(0), 0);
	
	//Save to file.
	image->Save(outputFile.c_str(), ImageFormat::GetPng());
	workbook->Dispose();
}