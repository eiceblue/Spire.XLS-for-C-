#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"ReadImages.xlsx";
	wstring outputFile = outputFolder + L"DeleteAllImages_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Delete all images from the worksheet
	for (int i = sheet->GetPictures()->GetCount() - 1; i >= 0; i--)
	{
		sheet->GetPictures()->Get(i)->Remove(true);
	}

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}