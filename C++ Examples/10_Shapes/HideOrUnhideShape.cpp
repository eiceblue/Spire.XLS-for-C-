#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Xls_5.xlsx";
	wstring outputFile = output_path + L"HideOrUnhideShape.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Show the second shape in the worksheet
	//sheet.PrstGeomShapes[1].Visible = true;

	//Hide the second shape in the worksheet
	sheet->GetPrstGeomShapes()->Get(1)->SetVisible(false);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}