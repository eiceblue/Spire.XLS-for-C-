#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SheetToImage.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"SheetToImage.png";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Save to file.
	sheet->ToImage(sheet->GetFirstRow(), sheet->GetFirstColumn(), sheet->GetLastRow(), sheet->GetLastColumn())->Save(outputFile.c_str());
	workbook->Dispose();
}
