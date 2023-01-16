#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_1.xlsx";
	wstring outputFile = output_path + L"CopyCellFormat_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Copy the cell format from column 2 and apply to cells of column 5.
	int count = sheet->GetRows()->GetCount();
	for (int i = 1; i < count + 1; i++)
	{
		sheet->GetRange((L"E" + to_wstring(i)).c_str())->SetStyle(sheet->GetRange((L"B" + to_wstring(i)).c_str())->GetStyle());
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

