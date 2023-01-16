#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"ScatterChart.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	sheet->SetName(L"Scatter Chart");

	//Set chart data
	//Add data
	sheet->SetName(L"Demo");
	sheet->GetRange(L"A1")->SetValue(L"Month");
	sheet->GetRange(L"A2")->SetValue(L"Jan.");
	sheet->GetRange(L"A3")->SetValue(L"Feb.");
	sheet->GetRange(L"A4")->SetValue(L"Mar.");
	sheet->GetRange(L"A5")->SetValue(L"Apr.");
	sheet->GetRange(L"A6")->SetValue(L"May.");
	sheet->GetRange(L"A7")->SetValue(L"Jun.");
	sheet->GetRange(L"B1")->SetValue(L"Planned");
	sheet->GetRange(L"B2")->SetNumberValue(3.3);
	sheet->GetRange(L"B3")->SetNumberValue(2.5);
	sheet->GetRange(L"B4")->SetNumberValue(2.0);
	sheet->GetRange(L"B5")->SetNumberValue(3.7);
	sheet->GetRange(L"B6")->SetNumberValue(4.5);
	sheet->GetRange(L"B7")->SetNumberValue(4.0);
	sheet->GetRange(L"C1")->SetValue(L"Actual");
	sheet->GetRange(L"C2")->SetNumberValue(3.8);
	sheet->GetRange(L"C3")->SetNumberValue(3.2);
	sheet->GetRange(L"C4")->SetNumberValue(1.7);
	sheet->GetRange(L"C5")->SetNumberValue(3.5);
	sheet->GetRange(L"C6")->SetNumberValue(4.5);
	sheet->GetRange(L"C7")->SetNumberValue(4.3);

	//Add a chart
	Chart* chart = sheet->GetCharts()->Add(ExcelChartType::ScatterMarkers);

	//Set region of chart data
	chart->SetDataRange(sheet->GetRange(L"B2:B7"));
	chart->SetSeriesDataFromRange(false);

	//Set position of chart
	chart->SetLeftColumn(1);
	chart->SetTopRow(11);
	chart->SetRightColumn(10);
	chart->SetBottomRow(28);

	chart->SetChartTitle(L"Scatter Chart");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);

	chart->GetSeries()->Get(0)->SetCategoryLabels(sheet->GetRange(L"A2:A7"));
	chart->GetSeries()->Get(0)->SetValues(sheet->GetRange(L"B2:B7"));

	//Add a trend line for the first series
	chart->GetSeries()->Get(0)->GetTrendLines()->Add(TrendLineType::Exponential);

	chart->GetPrimaryValueAxis()->SetTitle(L"Month");
	chart->GetPrimaryCategoryAxis()->SetTitle(L"Planned");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}