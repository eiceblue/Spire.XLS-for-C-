#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"EachWorksheetToDifferentPDFSample.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"EachWorksheetToPDF";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	for (int i = 0; i < workbook->GetWorksheets()->GetCount(); i++)
	{
		Worksheet* sheet = workbook->GetWorksheets()->Get(i);
		wstring FileName = outputFile + sheet->GetName() + L".pdf";
		//Save the sheet to PDF
		sheet->SaveToPdf(FileName.c_str());
	}
	workbook->Dispose();
}
