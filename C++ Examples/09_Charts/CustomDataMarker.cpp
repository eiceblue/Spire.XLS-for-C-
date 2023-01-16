#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"CustomDataMarker.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();
	workbook->CreateEmptySheets(1);

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add some sample data
	sheet->SetName(L"Demo");
	sheet->GetRange(L"A1")->SetValue(L"Tom");
	sheet->GetRange(L"A2")->SetNumberValue(1.5);
	sheet->GetRange(L"A3")->SetNumberValue(2.1);
	sheet->GetRange(L"A4")->SetNumberValue(3.6);
	sheet->GetRange(L"A5")->SetNumberValue(5.2);
	sheet->GetRange(L"A6")->SetNumberValue(7.3);
	sheet->GetRange(L"A7")->SetNumberValue(3.1);
	sheet->GetRange(L"B1")->SetValue(L"Kitty");
	sheet->GetRange(L"B2")->SetNumberValue(2.5);
	sheet->GetRange(L"B3")->SetNumberValue(4.2);
	sheet->GetRange(L"B4")->SetNumberValue(1.3);
	sheet->GetRange(L"B5")->SetNumberValue(3.2);
	sheet->GetRange(L"B6")->SetNumberValue(6.2);
	sheet->GetRange(L"B7")->SetNumberValue(4.7);

	//Create a Scatter-Markers chart based on the sample data
	Chart* chart = sheet->GetCharts()->Add(ExcelChartType::ScatterMarkers);
	chart->SetDataRange(sheet->GetRange(L"A1:B7"));
	chart->GetPlotArea()->SetVisible(false);
	chart->SetSeriesDataFromRange(false);
	chart->SetTopRow(5);
	chart->SetBottomRow(22);
	chart->SetLeftColumn(4);
	chart->SetRightColumn(11);
	chart->SetChartTitle(L"Chart with Markers");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(10);

	//Format the markers in the chart by setting the background color, foreground color, type, size and transparency
	ChartSerie* cs1 = chart->GetSeries()->Get(0);
	cs1->GetDataFormat()->SetMarkerBackgroundColor(Spire::Common::Color::GetRoyalBlue());
	cs1->GetDataFormat()->SetMarkerForegroundColor(Spire::Common::Color::GetWhiteSmoke());
	cs1->GetDataFormat()->SetMarkerSize(7);
	cs1->GetDataFormat()->SetMarkerStyle(ChartMarkerType::PlusSign);
	cs1->GetDataFormat()->SetMarkerTransparencyValue(0.8);

	ChartSerie* cs2 = chart->GetSeries()->Get(1);
	cs2->GetDataFormat()->SetMarkerBackgroundColor(Spire::Common::Color::GetPink());
	cs2->GetDataFormat()->SetMarkerSize(9);
	cs2->GetDataFormat()->SetMarkerStyle(ChartMarkerType::Triangle);
	cs2->GetDataFormat()->SetMarkerTransparencyValue(0.9);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}