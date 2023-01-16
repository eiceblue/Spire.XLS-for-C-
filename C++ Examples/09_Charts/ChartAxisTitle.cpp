#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"SampeB_5.xlsx";
	wstring outputFile = output_path + L"ChartAxisTitle.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the chart
	Chart* chart = sheet->GetCharts()->Get(0);

	//Set axis title
	chart->GetPrimaryCategoryAxis()->SetTitle(L"Category Axis");
	chart->GetPrimaryValueAxis()->SetTitle(L"Value axis");

	//Set font size
	chart->GetPrimaryCategoryAxis()->GetFont()->SetSize(12);
	chart->GetPrimaryValueAxis()->GetFont()->SetSize(12);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}