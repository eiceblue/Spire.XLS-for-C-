#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChangeDataLabel.xlsx";
	wstring outputFile = output_path + L"ChangeDataLabel_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the chart
	Spire::Xls::Chart* chart = sheet->GetCharts()->Get(0);

	//Change data label of the frist datapoint of the first series
	chart->GetSeries()->Get(0)->GetDataPoints()->Get(0)->GetDataLabels()->SetText(L"changed data label");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}