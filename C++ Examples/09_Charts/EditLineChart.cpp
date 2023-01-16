#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"LineChart.xlsx";
	wstring outputFile = output_path + L"EditLineChart.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the line chart
	Chart* chart = sheet->GetCharts()->Get(0);

	//Add a new series
	ChartSerie* cs = chart->GetSeries()->Add(L"Added");

	//Set the values for the series
	cs->SetValues(sheet->GetRange(L"I1:L1"));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}