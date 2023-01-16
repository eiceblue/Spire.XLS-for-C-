#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartSample1.xlsx";
	wstring outputFile = output_path + L"ResizeAndMoveChart.xlsx";

	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());
	//Get the chart from the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	Chart* chart = sheet->GetCharts()->Get(0);

	//Set position of the chart
	chart->SetLeftColumn(5);
	chart->SetTopRow(1);

	//Resize the chart
	chart->SetWidth(500);
	chart->SetHeight(350);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}