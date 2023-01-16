#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring inputFile = data_path + L"Template_Xls_4.xlsx";


	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	wstring* content = new wstring();

	//Get the number of cells.
	content->append(L"Number of Cells: " + to_wstring(sheet->GetCells()->GetCount()));

	workbook->Dispose();
	wcout << *content << endl;
	system("pause");
}

