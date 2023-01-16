#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartSample1.xlsx";
	wstring outputFile = output_path + L"AjustBarSpace.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet from workbook and then get the first chart from the worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	Chart* chart = sheet->GetCharts()->Get(0);
	int y = workbook->GetWorksheets()->GetCount();
	int x = chart->GetSeries()->GetCount();
	//Ajust the space between bars
	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		ChartSerie* cs = chart->GetSeries()->Get(i);
		cs->GetFormat()->GetOptions()->SetGapWidth(200);
		cs->GetFormat()->GetOptions()->SetOverlap(0);
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}