#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"FindCellsSample.xlsx";
	wstring outputFile = output_path + L"FindDataInSpecificRange_result.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Specify a range
	CellRange* range = sheet->GetRange(1, 1, 12, 8);

	//Create a string
	wstring* builder = new wstring();

	//Find string from this range
	auto textRanges = range->FindAllString(L"E-iceblue", false, false);

	//Append the address of found cells in builder
	if (textRanges->GetCount() != 0)
	{
		for (int i = 0; i < textRanges->GetCount(); i++)
		{
			CellRange* cr = textRanges->GetItem(i);
			wstring address = cr->GetRangeAddress();
			builder->append(L"The address of found text cell is: " + address + L"\n");
		}
	}
	else
	{
		builder->append(L"No cell contain the text.\n");
	}


	//Find number from this range
	auto ranges = range->FindAllNumber(100, true);

	//Append the address of found cells in builder
	if (ranges->GetCount() != 0)
	{
		for (int i = 0; i < ranges->GetCount(); i++)
		{
			CellRange* r = ranges->GetItem(i);
			wstring address = r->GetRangeAddress();
			builder->append(L"The address of found number cell is: " + address + L"\n");
		}
	}
	else
	{
		builder->append(L"No cell contain the number.\n");
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *builder << endl;
	ofs.close();
	workbook->Dispose();

}

