#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"PyramidColumn.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	sheet->SetName(L"Chart");

	//Set chart data
	sheet->GetRange(L"A1")->SetValue(L"Year");
	sheet->GetRange(L"A2")->SetValue(L"2002");
	sheet->GetRange(L"A3")->SetValue(L"2003");
	sheet->GetRange(L"A4")->SetValue(L"2004");
	sheet->GetRange(L"A5")->SetValue(L"2005");

	sheet->GetRange(L"B1")->SetValue(L"Sales");
	sheet->GetRange(L"B2")->SetNumberValue(4000);
	sheet->GetRange(L"B3")->SetNumberValue(6000);
	sheet->GetRange(L"B4")->SetNumberValue(7000);
	sheet->GetRange(L"B5")->SetNumberValue(8500);

	//Style
	sheet->GetRange(L"A1:B1")->SetRowHeight(15);
	sheet->GetRange(L"A1:B1")->GetStyle()->SetColor(Spire::Common::Color::GetDarkGray());
	sheet->GetRange(L"A1:B1")->GetStyle()->GetFont()->SetColor(Spire::Common::Color::GetWhite());
	sheet->GetRange(L"A1:B1")->GetStyle()->SetVerticalAlignment(VerticalAlignType::Center);
	sheet->GetRange(L"A1:B1")->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Center);

	sheet->GetRange(L"B2:C5")->GetStyle()->SetNumberFormat(L"\"$\"#,##0");

	//Add a chart
	Chart* chart = sheet->GetCharts()->Add();

	//Set region of chart data
	chart->SetDataRange(sheet->GetRange(L"B2:B5"));
	chart->SetSeriesDataFromRange(false);

	//Set position of chart
	chart->SetLeftColumn(1);
	chart->SetTopRow(6);
	chart->SetRightColumn(11);
	chart->SetBottomRow(29);

	chart->SetChartType(ExcelChartType::Pyramid3DClustered);

	//Chart title
	chart->SetChartTitle(L"Sales by year");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);

	chart->GetPrimaryCategoryAxis()->SetTitle(L"Year");
	chart->GetPrimaryCategoryAxis()->GetFont()->SetIsBold(true);
	chart->GetPrimaryCategoryAxis()->GetTitleArea()->SetIsBold(true);

	chart->GetPrimaryValueAxis()->SetTitle(L"Sales(in Dollars)");
	chart->GetPrimaryValueAxis()->SetHasMajorGridLines(false);
	chart->GetPrimaryValueAxis()->SetMinValue(1000);
	chart->GetPrimaryValueAxis()->GetTitleArea()->SetIsBold(true);
	chart->GetPrimaryValueAxis()->GetTitleArea()->SetTextRotationAngle(90);

	ChartSerie* cs = chart->GetSeries()->Get(0);
	cs->SetCategoryLabels(sheet->GetRange(L"A2:A5"));
	cs->GetFormat()->GetOptions()->SetIsVaryColor(true);

	chart->GetLegend()->SetPosition(LegendPositionType::Top);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}