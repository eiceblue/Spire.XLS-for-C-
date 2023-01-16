#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"ReadImages.xlsx";
	wstring outputFile = outputFolder + L"CroppedPositionOfPicture_out.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the image from the first sheet
	XlsBitmapShape* picture = sheet->GetPictures()->Get(0);

	//Get the cropped position
	int left = picture->GetLeft();
	int top = picture->GetTop();
	int width = picture->GetWidth();
	int height = picture->GetHeight();

	//Create string to append text 
	wstring* content = new wstring();

	//Set string format for displaying
	wstring displayString = L"Crop position: Left "+ to_wstring(left) +  L"\r\nCrop position: Top " + to_wstring(top) + L"\r\nCrop position: Width " + to_wstring(width) + L"\r\nCrop position: Height " + to_wstring(height);

	content->append(displayString);

	//Save to file
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}