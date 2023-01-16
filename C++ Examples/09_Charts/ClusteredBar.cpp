#include "pch.h"
using namespace Spire::Xls;

int main() {
	
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"ClusteredBar.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	sheet->SetName(L"ClusteredBar");

	//Set chart data
	sheet->GetRange(L"A1")->SetValue(L"Country");
	sheet->GetRange(L"A2")->SetValue(L"Cuba");
	sheet->GetRange(L"A3")->SetValue(L"Mexico");
	sheet->GetRange(L"A4")->SetValue(L"France");
	sheet->GetRange(L"A5")->SetValue(L"German");

	sheet->GetRange(L"B1")->SetValue(L"Jun");
	sheet->GetRange(L"B2")->SetNumberValue(6000);
	sheet->GetRange(L"B3")->SetNumberValue(8000);
	sheet->GetRange(L"B4")->SetNumberValue(9000);
	sheet->GetRange(L"B5")->SetNumberValue(8500);

	sheet->GetRange(L"C1")->SetValue(L"Aug");
	sheet->GetRange(L"C2")->SetNumberValue(3000);
	sheet->GetRange(L"C3")->SetNumberValue(2000);
	sheet->GetRange(L"C4")->SetNumberValue(2300);
	sheet->GetRange(L"C5")->SetNumberValue(4200);

	//Style
	sheet->GetRange(L"A1:C1")->SetRowHeight(15);
	sheet->GetRange(L"A1:C1")->GetStyle()->SetColor(Spire::Common::Color::GetDarkGray());
	sheet->GetRange(L"A1:C1")->GetStyle()->GetFont()->SetColor(Spire::Common::Color::GetWhite());
	sheet->GetRange(L"A1:C1")->GetStyle()->SetVerticalAlignment(VerticalAlignType::Center);
	sheet->GetRange(L"A1:C1")->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Center);

	sheet->GetRange(L"B2:C5")->GetStyle()->SetNumberFormat(L"\"$\"#,##0");

	//Add a chart
	Spire::Xls::Chart* chart = sheet->GetCharts()->Add();

	//Set region of chart data
	chart->SetDataRange(sheet->GetRange(L"A1:C5"));
	chart->SetSeriesDataFromRange(false);

	//Set position of chart
	chart->SetLeftColumn(1);
	chart->SetTopRow(6);
	chart->SetRightColumn(11);
	chart->SetBottomRow(29);
	chart->SetChartType(ExcelChartType::BarClustered);

	//Chart title
	chart->SetChartTitle(L"Sales market by country");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);

	chart->GetPrimaryCategoryAxis()->SetTitle(L"Country");
	chart->GetPrimaryCategoryAxis()->GetFont()->SetIsBold(true);
	chart->GetPrimaryCategoryAxis()->GetTitleArea()->SetIsBold(true);
	chart->GetPrimaryCategoryAxis()->GetTitleArea()->SetTextRotationAngle(90);

	chart->GetPrimaryValueAxis()->SetTitle(L"Sales(in Dollars)");
	chart->GetPrimaryValueAxis()->SetHasMajorGridLines(false);
	chart->GetPrimaryValueAxis()->SetMinValue(1000);
	chart->GetPrimaryValueAxis()->GetTitleArea()->SetIsBold(true);

	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		ChartSerie* cs = chart->GetSeries()->Get(i);
		cs->GetFormat()->GetOptions()->SetIsVaryColor(true);
		cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
	}

	chart->GetLegend()->SetPosition(LegendPositionType::Top);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}