#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_1.xlsx";
	wstring outputFile = output_path + L"MergeCells_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Merge the seventh column in Excel file.
	sheet->GetColumns()->GetItem(6)->Merge();

	//Merge the particular range in Excel file.
	sheet->GetRange(L"A14:D14")->XlsRange::Merge();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();


}

