#include "pch.h"
using namespace Spire::Xls;

int main() { 
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"CreateDoughnutChart.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Insert data
	sheet->GetRange(L"A1")->SetValue(L"Country");
	sheet->GetRange(L"A1")->GetStyle()->GetFont()->SetIsBold(true);
	sheet->GetRange(L"A2")->SetValue(L"Cuba");
	sheet->GetRange(L"A3")->SetValue(L"Mexico");
	sheet->GetRange(L"A4")->SetValue(L"France");
	sheet->GetRange(L"A5")->SetValue(L"German");
	sheet->GetRange(L"B1")->SetValue(L"Sales");
	sheet->GetRange(L"B1")->GetStyle()->GetFont()->SetIsBold(true);
	sheet->GetRange(L"B2")->SetNumberValue(6000);
	sheet->GetRange(L"B3")->SetNumberValue(8000);
	sheet->GetRange(L"B4")->SetNumberValue(9000);
	sheet->GetRange(L"B5")->SetNumberValue(8500);

	//Add a new chart, set chart type as doughnut
	Chart* chart = sheet->GetCharts()->Add();
	chart->SetChartType(ExcelChartType::Doughnut);
	chart->SetDataRange(sheet->GetRange(L"A1:B5"));
	chart->SetSeriesDataFromRange(false);

	//Set position of chart
	chart->SetLeftColumn(4);
	chart->SetTopRow(2);
	chart->SetRightColumn(12);
	chart->SetBottomRow(22);

	//Chart title
	chart->SetChartTitle(L"Market share by country");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);

	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		ChartSerie* cs = chart->GetSeries()->Get(i);
		cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasPercentage(true);
	}

	chart->GetLegend()->SetPosition(LegendPositionType::Top);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}