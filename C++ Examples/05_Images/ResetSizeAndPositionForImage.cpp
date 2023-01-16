#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"intro_xls.png";
	wstring outputFile = outputFolder + L"ResetSizeAndPositionForImage_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add a picture to the first worksheet
	IPictureShape* picture = sheet->GetPictures()->XlsPicturesCollection::Add(1, 1, inputFile.c_str());

	//Set the size for the picture
	picture->SetWidth(200);
	picture->SetHeight(200);

	//Set the position for the picture
	picture->SetLeft(200);
	picture->SetTop(100);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}