#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"Template_Xls_1.xlsx";
	wstring inputImage = inputFolder + L"Background.png";
	wstring outputFile = outputFolder + L"InsertExcelBackgroundImage_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Load an image
	Spire::Common::Image* im = Bitmap::FromFile(inputImage.c_str());
	Bitmap* bm = Object::Convert<Bitmap>(im);

	//Set the image as background image of the worksheet
	sheet->GetPageSetup()->SetBackgoundImage(bm);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}