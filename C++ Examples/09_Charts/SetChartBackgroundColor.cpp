#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring  inputFile = input_path + L"ChartSample1.xlsx";
	wstring  outputFile = output_path + L"SetChartBackgroundColor.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	Chart* chart = sheet->GetCharts()->Get(0);

	//Set background color
	chart->GetChartArea()->SetForeGroundColor(Spire::Common::Color::GetLightYellow());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}