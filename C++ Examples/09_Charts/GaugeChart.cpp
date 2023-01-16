#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"GaugeChart.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	sheet->SetName(L"Gauge Chart");

	//Set chart data
	//Set value of specified cell
	sheet->GetRange(L"A1")->SetValue(L"Value");
	sheet->GetRange(L"A2")->SetValue(L"30");
	sheet->GetRange(L"A3")->SetValue(L"60");
	sheet->GetRange(L"A4")->SetValue(L"90");
	sheet->GetRange(L"A5")->SetValue(L"180");
	sheet->GetRange(L"C2")->SetValue(L"value");
	sheet->GetRange(L"C3")->SetValue(L"pointer");
	sheet->GetRange(L"C4")->SetValue(L"End");
	sheet->GetRange(L"D2")->SetValue(L"10");
	sheet->GetRange(L"D3")->SetValue(L"1");
	sheet->GetRange(L"D4")->SetValue(L"189");

	//Add a Doughnut chart
	Chart* chart = sheet->GetCharts()->Add(ExcelChartType::Doughnut);
	chart->SetDataRange(sheet->GetRange(L"A1:A5"));
	chart->SetSeriesDataFromRange(false);
	chart->SetHasLegend(true);

	//Set the position of chart
	chart->SetLeftColumn(2);
	chart->SetTopRow(7);
	chart->SetRightColumn(9);
	chart->SetBottomRow(25);

	//Get the series 1
	ChartSerie* cs1 = static_cast<ChartSerie*>(chart->GetSeries()->Get(L"Value"));
	cs1->GetFormat()->GetOptions()->SetDoughnutHoleSize(60);
	cs1->GetDataFormat()->GetOptions()->SetFirstSliceAngle(270);

	//Set the fill color
	cs1->GetDataPoints()->Get(0)->GetDataFormat()->GetFill()->SetForeColor(Spire::Common::Color::GetYellow());
	cs1->GetDataPoints()->Get(1)->GetDataFormat()->GetFill()->SetForeColor(Spire::Common::Color::GetPaleVioletRed());
	cs1->GetDataPoints()->Get(2)->GetDataFormat()->GetFill()->SetForeColor(Spire::Common::Color::GetDarkViolet());
	cs1->GetDataPoints()->Get(3)->GetDataFormat()->GetFill()->SetVisible(false);

	//Add a series with pie chart
	ChartSerie* cs2 = static_cast<ChartSerie*>(chart->GetSeries()->Add(L"Pointer", ExcelChartType::Pie));

	//Set the value
	cs2->SetValues(sheet->GetRange(L"D2:D4"));
	cs2->SetUsePrimaryAxis(false);
	cs2->GetDataPoints()->Get(0)->GetDataLabels()->SetHasValue(true);
	cs2->GetDataFormat()->GetOptions()->SetFirstSliceAngle(270);
	cs2->GetDataPoints()->Get(0)->GetDataFormat()->GetFill()->SetVisible(false);
	cs2->GetDataPoints()->Get(1)->GetDataFormat()->GetFill()->SetFillType(ShapeFillType::SolidColor);
	cs2->GetDataPoints()->Get(1)->GetDataFormat()->GetFill()->SetForeColor(Spire::Common::Color::GetBlack());
	cs2->GetDataPoints()->Get(2)->GetDataFormat()->GetFill()->SetVisible(false);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}