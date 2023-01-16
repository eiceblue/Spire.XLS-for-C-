#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"ReadImages.xlsx";
	wstring outputFile = outputFolder + L"ReadImages_out.png";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the first image
	XlsBitmapShape* pic = sheet->GetPictures()->Get(0);
	pic->GetPicture()->Save(outputFile.c_str(), Spire::Common::ImageFormat::GetPng());
}