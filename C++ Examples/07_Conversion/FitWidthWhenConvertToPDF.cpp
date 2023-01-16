#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_2.xlsx";
   	 wstring output_path = OUTPUTPATH;
   	 wstring outputFile = output_path + L"FitWidthWhenConvertToPDF.pdf";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	for (int i = 0; i < workbook->GetWorksheets()->GetCount(); i++)
	{
		Worksheet* sheet = workbook->GetWorksheets()->Get(i);
		//Auto fit page height
		sheet->GetPageSetup()->SetFitToPagesTall(0);
		//Fit one page width
		sheet->GetPageSetup()->SetFitToPagesWide(1);
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	workbook->Dispose();
}
