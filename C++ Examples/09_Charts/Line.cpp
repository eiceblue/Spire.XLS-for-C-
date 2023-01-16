#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"Line.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	sheet->SetName(L"Line Chart");

	//Set chart data
	//Set value of specified cell
	sheet->GetRange(L"A1")->SetValue(L"Country");
	sheet->GetRange(L"A2")->SetValue(L"Cuba");
	sheet->GetRange(L"A3")->SetValue(L"Mexico");
	sheet->GetRange(L"A4")->SetValue(L"France");
	sheet->GetRange(L"A5")->SetValue(L"German");

	sheet->GetRange(L"B1")->SetValue(L"Jun");
	sheet->GetRange(L"B2")->SetNumberValue(3300);
	sheet->GetRange(L"B3")->SetNumberValue(2300);
	sheet->GetRange(L"B4")->SetNumberValue(4500);
	sheet->GetRange(L"B5")->SetNumberValue(6700);

	sheet->GetRange(L"C1")->SetValue(L"Jul");
	sheet->GetRange(L"C2")->SetNumberValue(7500);
	sheet->GetRange(L"C3")->SetNumberValue(2900);
	sheet->GetRange(L"C4")->SetNumberValue(2300);
	sheet->GetRange(L"C5")->SetNumberValue(4200);

	sheet->GetRange(L"D1")->SetValue(L"Aug");
	sheet->GetRange(L"D2")->SetNumberValue(7400);
	sheet->GetRange(L"D3")->SetNumberValue(6900);
	sheet->GetRange(L"D4")->SetNumberValue(7800);
	sheet->GetRange(L"D5")->SetNumberValue(4200);


	sheet->GetRange(L"E1")->SetValue(L"Sep");
	sheet->GetRange(L"E2")->SetNumberValue(8000);
	sheet->GetRange(L"E3")->SetNumberValue(7200);
	sheet->GetRange(L"E4")->SetNumberValue(8500);
	sheet->GetRange(L"E5")->SetNumberValue(5600);

	//Style
	sheet->GetRange(L"A1:E1")->SetRowHeight(15);
	sheet->GetRange(L"A1:E1")->GetStyle()->SetColor(Spire::Common::Color::GetDarkGray());
	sheet->GetRange(L"A1:E1")->GetStyle()->GetFont()->SetColor(Spire::Common::Color::GetWhite());
	sheet->GetRange(L"A1:E1")->GetStyle()->SetVerticalAlignment(VerticalAlignType::Center);
	sheet->GetRange(L"A1:E1")->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Center);

	sheet->GetRange(L"B2:D5")->GetStyle()->SetNumberFormat(L"\"$\"#,##0");

	//Add a chart
	Chart* chart = sheet->GetCharts()->Add();
	chart->SetChartType(ExcelChartType::Line);

	//Set region of chart data
	chart->SetDataRange(sheet->GetRange(L"A1:E5"));

	//Set position of chart
	chart->SetLeftColumn(1);
	chart->SetTopRow(6);
	chart->SetRightColumn(11);
	chart->SetBottomRow(29);

	//Set chart title
	chart->SetChartTitle(L"Sales market by country");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);

	chart->GetPrimaryCategoryAxis()->SetTitle(L"Month");
	chart->GetPrimaryCategoryAxis()->GetFont()->SetIsBold(true);
	chart->GetPrimaryCategoryAxis()->GetTitleArea()->SetIsBold(true);

	chart->GetPrimaryValueAxis()->SetTitle(L"Sales(in Dollars)");
	chart->GetPrimaryValueAxis()->SetHasMajorGridLines(false);
	chart->GetPrimaryValueAxis()->GetTitleArea()->SetTextRotationAngle(90);
	chart->GetPrimaryValueAxis()->SetMinValue(1000);
	chart->GetPrimaryValueAxis()->GetTitleArea()->SetIsBold(true);

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