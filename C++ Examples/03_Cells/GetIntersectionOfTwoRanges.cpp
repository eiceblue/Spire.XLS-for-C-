#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_1.xlsx";
	wstring outputFile = output_path + L"GetIntersectionOfTwoRanges_result.txt";

	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the two ranges.
	CellRange* range = sheet->GetRange(L"A2:D7")->Intersect(sheet->GetRange(L"B2:E8"));

	wstring* content = new wstring();
	content->append(L"The intersection of the two ranges \"A2:D7\" and \"B2:E8\" is:\n");

	for (int i = 0; i < range->GetCells()->GetCount(); i++)
	{
		CellRange* cr = range->GetCells()->GetItem(i);
		content->append(cr->GetValue());
		content->append(L"\n");
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}

