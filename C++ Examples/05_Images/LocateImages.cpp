#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"ReadImages.xlsx";
	wstring outputFile = outputFolder + L"LocateImages_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the first picture
	XlsBitmapShape* pic = sheet->GetPictures()->Get(0);

	//Set offset
	pic->SetLeftColumnOffset(300);
	pic->SetTopRowOffset(300);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}