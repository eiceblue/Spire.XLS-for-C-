#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SampeB_4.xlsx";
	wstring outputFile = output_path + L"SetColorForChartArea.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the chart
	Chart* chart = sheet->GetCharts()->Get(0);

	//Set color for chart area
	chart->GetChartArea()->GetFill()->SetForeColor(Spire::Common::Color::GetLightSeaGreen());

	//Set color for plot area
	chart->GetPlotArea()->GetFill()->SetForeColor(Spire::Common::Color::GetLightGray());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
