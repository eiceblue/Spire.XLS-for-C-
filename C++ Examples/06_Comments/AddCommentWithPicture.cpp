#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"Logo.png";
	wstring outputFile = outputFolder + L"AddCommentWithPicture_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	sheet->GetRange(L"C6")->SetText(L"E-iceblue");
	//Add the comment
	ExcelComment* comment = sheet->GetRange(L"C6")->AddExcelComment();

	//Load the image file
	Image* image = Image::FromFile(inputFile.c_str());

	comment->GetFill()->CustomPicture(image, L"logo.png");
	//Set the height and width of comment
	comment->SetHeight(image->GetHeight());
	comment->SetWidth(image->GetWidth());
	comment->SetVisible(true);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}