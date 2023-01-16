#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"FormatAxis.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	sheet->SetName(L"FormatAxis");

	//Set chart data
	//Set value of specified cell
	sheet->GetRange(L"A1")->SetValue(L"Month");
	sheet->GetRange(L"A2")->SetValue(L"Jan");
	sheet->GetRange(L"A3")->SetValue(L"Feb");
	sheet->GetRange(L"A4")->SetValue(L"Mar");
	sheet->GetRange(L"A5")->SetValue(L"Apr");
	sheet->GetRange(L"A6")->SetValue(L"May");
	sheet->GetRange(L"A7")->SetValue(L"Jun");
	sheet->GetRange(L"A8")->SetValue(L"Jul");
	sheet->GetRange(L"A9")->SetValue(L"Aug");

	sheet->GetRange(L"B1")->SetValue(L"Planned");
	sheet->GetRange(L"B2")->SetNumberValue(38);
	sheet->GetRange(L"B3")->SetNumberValue(47);
	sheet->GetRange(L"B4")->SetNumberValue(39);
	sheet->GetRange(L"B5")->SetNumberValue(36);
	sheet->GetRange(L"B6")->SetNumberValue(27);
	sheet->GetRange(L"B7")->SetNumberValue(25);
	sheet->GetRange(L"B8")->SetNumberValue(36);
	sheet->GetRange(L"B9")->SetNumberValue(48);

	//Add a chart
	Chart* chart = sheet->GetCharts()->Add(ExcelChartType::ColumnClustered);
	chart->SetDataRange(sheet->GetRange(L"B1:B9"));
	chart->SetSeriesDataFromRange(false);
	chart->GetPlotArea()->SetVisible(false);
	chart->SetTopRow(10);
	chart->SetBottomRow(28);
	chart->SetLeftColumn(2);
	chart->SetRightColumn(10);
	chart->SetChartTitle(L"Chart with Customized Axis");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);
	ChartSerie* cs1 = chart->GetSeries()->Get(0);
	cs1->SetCategoryLabels(sheet->GetRange(L"A2:A9"));

	//Format axis
	chart->GetPrimaryValueAxis()->SetMajorUnit(8);
	chart->GetPrimaryValueAxis()->SetMinorUnit(2);
	chart->GetPrimaryValueAxis()->SetMaxValue(50);
	chart->GetPrimaryValueAxis()->SetMinValue(0);
	chart->GetPrimaryValueAxis()->SetIsReverseOrder(false);
	chart->GetPrimaryValueAxis()->SetMajorTickMark(TickMarkType::TickMarkOutside);
	chart->GetPrimaryValueAxis()->SetMinorTickMark(TickMarkType::TickMarkInside);
	chart->GetPrimaryValueAxis()->SetTickLabelPosition(TickLabelPositionType::TickLabelPositionNextToAxis);
	chart->GetPrimaryValueAxis()->SetCrossesAt(0);

	//Set NumberFormat
	chart->GetPrimaryValueAxis()->SetNumberFormat(L"$#,##0");
	chart->GetPrimaryValueAxis()->SetIsSourceLinked(false);

	ChartSerie* serie = chart->GetSeries()->Get(0);
	Spire::Common::IEnumerator<XlsChartDataPoint>* ie = serie->GetDataPoints()->GetEnumerator();
	while (ie->MoveNext())
	{
		IChartDataPoint* dataPoint = ie->GetCurrent();
		//Format Series
		dataPoint->GetDataFormat()->GetFill()->SetFillType(ShapeFillType::SolidColor);
		dataPoint->GetDataFormat()->GetFill()->SetForeColor(Spire::Common::Color::GetLightGreen());

		//Set transparency
		dataPoint->GetDataFormat()->GetFill()->SetTransparency(0.3);
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}