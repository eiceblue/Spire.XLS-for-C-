#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SetBorderWidthOfMarker.xlsx";
	wstring outputFile = output_path + L"SetBorderWidthOfMarker.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the chart from the first worksheet
	Chart* chart = workbook->GetWorksheets()->Get(0)->GetCharts()->Get(0);

	chart->GetSeries()->Get(0)->GetDataFormat()->SetMarkerBorderWidth(1.5); //unit is pt

	chart->GetSeries()->Get(1)->GetDataFormat()->SetMarkerBorderWidth(2.5); //unit is pt

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
