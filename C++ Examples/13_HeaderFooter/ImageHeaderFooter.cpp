#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ImageHeaderFooter.xlsx";
	wstring outputFile = output_path + L"ImageHeaderFooter.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Load an image from disk
	Image* image = Image::FromFile((input_path + L"Logo.png").c_str());

	//Set the image header
	sheet->GetPageSetup()->SetLeftHeaderImage(image);
	sheet->GetPageSetup()->SetLeftHeader(L"&G");

	//Set the image footer
	sheet->GetPageSetup()->SetCenterFooterImage(image);
	sheet->GetPageSetup()->SetCenterFooter(L"&G");

	//Set the view mode of the sheet
	sheet->SetViewMode(ViewMode::Layout);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}