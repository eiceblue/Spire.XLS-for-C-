#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"DataSorting.xls";
	wstring outputFile = output_path + L"GetCellAddress_result.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	wstring* builder = new wstring();

	//Get a cell range
	CellRange* range = sheet->GetRange(L"A1:B5");

	//Get address of range
	wstring address = range->GetRangeAddressLocal();
	builder->append(L"Address of range: " + address + L"\n");

	//Get the cell count of range
	int count = range->GetCellsCount();
	builder->append(L"Cell count of range: " + to_wstring(count) + L"\n");

	//Get the address of the entire column of range
	wstring entireColAddress = range->GetEntireColumn()->GetRangeAddressLocal();
	builder->append(L"Address of entire column of the range: " + entireColAddress+ L"\n");

	//Get the address of the entire row of range
	wstring entireRowAddress = range->GetEntireColumn()->GetRangeAddressLocal();
	builder->append(L"Address of entire row of the range " + entireRowAddress + L"\n");

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *builder << endl;
	ofs.close();
	workbook->Dispose();
}

