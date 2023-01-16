#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"CreateCustomChart.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set values
	sheet->GetRange(L"A1")->SetValue(L"60");
	sheet->GetRange(L"A2")->SetValue(L"90");
	sheet->GetRange(L"A3")->SetValue(L"80");
	sheet->GetRange(L"A4")->SetValue(L"85");
	sheet->GetRange(L"B1")->SetValue(L"100");
	sheet->GetRange(L"B2")->SetValue(L"110");
	sheet->GetRange(L"B3")->SetValue(L"80");
	sheet->GetRange(L"B4")->SetValue(L"70");

	//Add a chart based on the data from A1 to B4
	Chart* chart = sheet->GetCharts()->Add();
	chart->SetDataRange(sheet->GetRange(L"A1:B4"));
	chart->SetSeriesDataFromRange(false);

	//Set position of chart
	chart->SetLeftColumn(1);
	chart->SetTopRow(10);
	chart->SetRightColumn(7);
	chart->SetBottomRow(25);

	//Apply different chart type to different series
	auto cs1 = static_cast<ChartSerie*>(chart->GetSeries()->Get(0));
	cs1->SetSerieType(ExcelChartType::ColumnClustered);
	auto cs2 = static_cast<ChartSerie*>(chart->GetSeries()->Get(1));
	cs2->SetSerieType(ExcelChartType::Line);

	chart->SetChartTitle(L"Custom chart");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}