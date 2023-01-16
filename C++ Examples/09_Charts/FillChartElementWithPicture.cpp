#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartSample1.xlsx";
	wstring inputImg = input_path + L"background.png";
	wstring outputFile = output_path + L"FillChartElementWithPicture.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the first chart
	Chart* chart = sheet->GetCharts()->Get(0);
	//int i = chart->GetSeries()->Get(0)->GetDataPoints()->GetCount();
	//Fill chart area with image
	chart->GetChartArea()->GetFill()->CustomPicture(Image::FromFile(inputImg.c_str()), L"None");

	chart->GetPlotArea()->GetFill()->SetTransparency(0.9);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}