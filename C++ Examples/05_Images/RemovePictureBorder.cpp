#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"PictureBorder.xlsx";
	wstring outputFile = outputFolder + L"RemovePictureBorder_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the first picture from the first worksheet
	XlsBitmapShape* picture = sheet->GetPictures()->Get(0);

	//Remove the picture border
	//Method-1:
	picture->GetLine()->SetVisible(false);

	//Method-2:
	//picture->GetLine()->SetWeight(0);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}