#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_2.xlsx";
   	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ToImageWithoutWhiteSpace.png";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set the margin as 0 to remove the white space around the image
	sheet->GetPageSetup()->SetLeftMargin(0);
	sheet->GetPageSetup()->SetBottomMargin(0);
	sheet->GetPageSetup()->SetTopMargin(0);
	sheet->GetPageSetup()->SetRightMargin(0);
	Image* image = sheet->ToImage(sheet->GetFirstRow(), sheet->GetFirstColumn(), sheet->GetLastRow(), sheet->GetLastColumn());

	//Save to file.
	image->Save(outputFile.c_str());
	workbook->Dispose();
}
