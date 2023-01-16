#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"ExplodedDoughnut.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	sheet->SetName(L"ExplodedDoughnut");

	//Set chart data
	sheet->GetRange(L"A1")->SetValue(L"Country");
	sheet->GetRange(L"A2")->SetValue(L"Cuba");
	sheet->GetRange(L"A3")->SetValue(L"Mexico");
	sheet->GetRange(L"A4")->SetValue(L"France");
	sheet->GetRange(L"A5")->SetValue(L"German");


	sheet->GetRange(L"B1")->SetValue(L"Sales");
	sheet->GetRange(L"B2")->SetNumberValue(6000);
	sheet->GetRange(L"B3")->SetNumberValue(8000);
	sheet->GetRange(L"B4")->SetNumberValue(9000);
	sheet->GetRange(L"B5")->SetNumberValue(8500);

	//Style
	sheet->GetRange(L"A1:B1")->SetRowHeight(15);
	sheet->GetRange(L"A1:B1")->GetStyle()->SetColor(Spire::Common::Color::GetDarkGray());
	sheet->GetRange(L"A1:B1")->GetStyle()->GetFont()->SetColor(Spire::Common::Color::GetWhite());
	sheet->GetRange(L"A1:B1")->GetStyle()->SetVerticalAlignment(VerticalAlignType::Center);
	sheet->GetRange(L"A1:B1")->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Center);

	sheet->GetRange(L"B2:B5")->GetStyle()->SetNumberFormat(L"\"$\"#,##0");

	//Add a chart
	Chart* chart = sheet->GetCharts()->Add();
	chart->SetChartType(ExcelChartType::DoughnutExploded);

	//Set position of chart
	chart->SetLeftColumn(1);
	chart->SetTopRow(6);
	chart->SetRightColumn(11);
	chart->SetBottomRow(29);

	//Set region of chart data
	chart->SetDataRange(sheet->GetRange(L"A1:B5"));
	chart->SetSeriesDataFromRange(false);

	//Chart title
	chart->SetChartTitle(L"Sales market by country");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);

	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		ChartSerie* cs = chart->GetSeries()->Get(i);
		cs->GetFormat()->GetOptions()->SetIsVaryColor(true);
		cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
	}

	chart->GetPlotArea()->GetFill()->SetVisible(false);
	chart->GetLegend()->SetPosition(LegendPositionType::Top);


	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}