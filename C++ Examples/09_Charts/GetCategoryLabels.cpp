#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"SampeB_4.xlsx";
	wstring outputFile = output_path + L"GetCategoryLabels.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the chart
	Chart* chart = sheet->GetCharts()->Get(0);

	//Get the cell range of the category labels
	wstring* content = new wstring();
	IXLSRange* cr = chart->GetPrimaryCategoryAxis()->GetCategoryLabels();
	for (int i = 0; i < cr->GetCells()->GetCount(); i++)
	{
		auto cell = cr->GetCells()->GetItem(i);
		wstring value = cell->GetValue();
		content->append(value + L"\r\n");
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}