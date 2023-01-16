#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_4.xlsx";
	wstring outputFile = output_path + L"GetCellValueByCellName_result.txt";

	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Specify a cell by its name.
	CellRange* cell = sheet->GetRange(L"A2");

	wstring* content = new wstring();

	//Get vaule of cell "A2".
	wstring cellValue = cell->GetValue();
	content->append(L"The vaule of cell A2 is: " + cellValue);

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}

