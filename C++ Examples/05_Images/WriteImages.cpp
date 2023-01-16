#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputImage = inputFolder + L"intro_xls.png";
	wstring outputFile = outputFolder + L"WriteImages_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add an image to the specific cell
	sheet->GetPictures()->XlsPicturesCollection::Add(14, 5, inputImage.c_str());

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}