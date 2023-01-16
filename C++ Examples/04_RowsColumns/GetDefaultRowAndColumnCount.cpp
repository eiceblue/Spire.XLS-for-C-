#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring outputFolder = OUTPUTPATH;
	wstring outputFile = outputFolder + L"GetDefaultRowAndColumnCount_out.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Clear all worksheets
	workbook->GetWorksheets()->Clear();

	//Create a new worksheet
	Worksheet* sheet = workbook->CreateEmptySheet();

	wstring* builder = new wstring();
	//Get row and column count
	int rowCount = sheet->GetRows()->GetCount();
	int columnCount = sheet->GetColumns()->GetCount();

	//Append text in string
	builder->append(L"The default row count is :" + to_wstring(rowCount) + L"\n");
	builder->append(L"The default column count is :" + to_wstring(columnCount) + L"\n");

	//Save to file
	ofs.open(outputFile, ios::out);
	ofs << *builder << endl;
	ofs.close();
	workbook->Dispose();
}