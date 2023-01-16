#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"templateAz.xlsx";
	wstring outputFile = output_path + L"MakeCellActive.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the 2nd sheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(1);

	//Set the 2nd sheet as an active sheet.
	sheet->Activate();

	//Set B2 cell as an active cell in the worksheet.
	sheet->SetActiveCell(sheet->GetRange(L"B2"));

	//Set the B column as the first visible column in the worksheet.
	sheet->SetFirstVisibleColumn(1);

	//Set the 2nd row as the first visible row in the worksheet.
	sheet->SetFirstVisibleRow(1);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}