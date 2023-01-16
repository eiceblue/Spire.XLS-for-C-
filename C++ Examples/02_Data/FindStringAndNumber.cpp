#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"FindCellsSample.xlsx";
	wstring outputFile = output_path + L"FindStringAndNumber_result.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Find cells with the input string
	auto ranges = sheet->FindAllString(L"E-iceblue", false, false);

	//Create a string* builder = new string()
	wstring* builder = new wstring();

	//Append the address of found cells in builder
	if (ranges->GetCount() != 0)
	{
		for (int i = 0; i < ranges->GetCount(); i++)
		{
			CellRange* cr = ranges->GetItem(i);
			wstring address = cr->GetRangeAddress();
			builder->append(L"The address of found text cell is: " + address + L"\n");
		}
	}
	else
	{
		builder->append(L"No cells that contain the text\n");
	}

	//Find cells with the input integer or double
	auto numberRanges = sheet->FindAllNumber(100, true);

	//Append the address of found cells in builder
	if (numberRanges->GetCount() != 0)
	{
		for (int i = 0; i < numberRanges->GetCount(); i++)
		{
			CellRange* cr = numberRanges->GetItem(i);
			wstring address = cr->GetRangeAddress();
			builder->append(L"The address of found number cell is: " + address + L"\n");
		}
	}
	else
	{
		builder->append(L"No cells that contain the number\n");
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *builder << endl;
	ofs.close();
	workbook->Dispose();

}

