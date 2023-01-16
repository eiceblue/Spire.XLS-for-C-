#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring  input_path = DATAPATH;
	wstring  output_path = OUTPUTPATH;

	wstring  inputFile = input_path + L"SampeB_4.xlsx";
	wstring  outputFile = output_path + L"ChangeChartSizeAndPosition.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the chart
	Chart* chart = sheet->GetCharts()->Get(0);

	//Change chart size
	chart->SetWidth(600);
	chart->SetHeight(500);

	//Change chart position
	chart->SetLeftColumn(3);
	chart->SetTopRow(7);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}