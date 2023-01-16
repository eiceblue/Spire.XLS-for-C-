#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"logo.png";
	wstring outputFile = outputFolder + L"PictureOffset_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Insert a picture
	ExcelPicture* pic = dynamic_cast<ExcelPicture*>(sheet->GetPictures()->Add(2, 2, inputFile.c_str()));

	//Set left offset and top offset from the current range
	pic->SetLeftColumnOffset(200);
	pic->SetTopRowOffset(100);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}