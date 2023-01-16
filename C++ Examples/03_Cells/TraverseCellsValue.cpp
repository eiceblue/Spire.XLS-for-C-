#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"CellValues.xlsx";
	wstring outputFile = outputFolder + L"TraverseCellsValue_out.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the cell range collection 
	Spire::Common::IList<XlsRange>* cellRangeCollection = sheet->GetCells();

	//Create string to append text
	wstring* content = new wstring();
	content->append(L"Values of the first sheet:\n");

	//Traverse cells value
	for (int i = 0; i < cellRangeCollection->GetCount(); i++)
	{
		XlsRange* cr = cellRangeCollection->GetItem(i);
		//Set string format for displaying
		wstring cell = cr->GetRangeAddress();
		wstring result = L"Cell: " + cell + L"   Value: " + cr->GetValue();
		content->append(result + L"\n");
	}

	//Save to txt file
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;

	ofs.close();
	workbook->Dispose();
}