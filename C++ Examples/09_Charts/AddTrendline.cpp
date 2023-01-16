#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartSample2.xlsx";
	wstring outputFile = output_path + L"AddTrendline.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//select chart and set logarithmic trendline
	Chart* chart = sheet->GetCharts()->Get(0);
	chart->SetChartTitle(L"Logarithmic Trendline");
	chart->GetSeries()->Get(0)->GetTrendLines()->Add(TrendLineType::Logarithmic);
	//select chart and set moving_average trendline
	Chart* chart1 = sheet->GetCharts()->Get(1);
	chart1->SetChartTitle(L"Moving Average Trendline");
	chart1->GetSeries()->Get(0)->GetTrendLines()->Add(TrendLineType::Moving_Average);
	//select chart and set linear trendline
	Chart* chart2 = sheet->GetCharts()->Get(2);
	chart2->SetChartTitle(L"Linear Trendline");
	chart2->GetSeries()->Get(0)->GetTrendLines()->Add(TrendLineType::Linear);
	//select chart and set exponential trendline
	Chart* chart3 = sheet->GetCharts()->Get(3);
	chart3->SetChartTitle(L"Exponential Trendline");
	chart3->GetSeries()->Get(0)->GetTrendLines()->Add(TrendLineType::Exponential);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
