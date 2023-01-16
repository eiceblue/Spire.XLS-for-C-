#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"CommentSample.xlsx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"HideOrShowComment.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Hide the second comment
	sheet->GetComments()->Get(1)->SetIsVisible(false);

	//Show the third comment
	sheet->GetComments()->Get(2)->SetIsVisible(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
