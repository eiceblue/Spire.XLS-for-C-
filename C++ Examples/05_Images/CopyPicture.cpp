#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"ReadImages.xlsx";
	wstring outputFile = outputFolder + L"CopyPicture_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add a new worksheet as destination sheet
	Worksheet* destinationSheet = workbook->GetWorksheets()->Add(L"DestSheet");

	//Get the first picture from the first worksheet
	XlsBitmapShape* sourcePicture = sheet->GetPictures()->Get(0);

	//Get the image
	Image* image = sourcePicture->GetPicture();

	//Add the image into the added worksheet 
	destinationSheet->GetPictures()->Add(2, 2, image);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}