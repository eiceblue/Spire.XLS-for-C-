#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"AccessCell.xlsx";
	wstring outputFile = output_path + L"AccessCell_result.txt";
	wfstream ofs;
	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());
	wstring* builder = new wstring();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Access cell by its name
	CellRange* range1 = sheet->GetRange(L"A1");
	wstring s1 = range1->GetText();
	builder->append(L"Value of range1: " + s1 + L"\n");

	//Access cell by index of row and column
	CellRange* range2 = sheet->GetRange(2, 1);
	wstring s2 = range2->GetText();
	builder->append(L"Value of range2: " + s2 + L"\n");

	//Access cell in cell collection
	XlsRange* range3 = sheet->GetCells()->GetItem(2);
	wstring s3 = range3->GetText();
	builder->append(L"Value of range3: " + s3 + L"\n");

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *builder << endl;
	ofs.close();
	workbook->Dispose();
}

