#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConversionSample1.xlsx";
   	 wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"SelectedRangeToPDF.pdf";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Add a new sheet to workbook
	workbook->GetWorksheets()->Add(L"newsheet");

	//Copy your area to new sheet.
	workbook->GetWorksheets()->Get(0)->GetRange(L"A9:E15")->Copy(workbook->GetWorksheets()->Get(1)->GetRange(L"A9:E15"), false, true);
	
	//Auto fit column width
	workbook->GetWorksheets()->Get(1)->GetRange(L"A9:E15")->AutoFitColumns();

	//Save to file.
	workbook->GetWorksheets()->Get(1)->SaveToPdf(outputFile.c_str());
	workbook->Dispose();
}
