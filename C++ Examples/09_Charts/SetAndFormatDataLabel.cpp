#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"SetAndFormatDataLabel.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();
	workbook->CreateEmptySheets(1);

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	sheet->SetName(L"Demo");
	sheet->GetRange(L"A1")->SetValue(L"Month");
	sheet->GetRange(L"A2")->SetValue(L"Jan");
	sheet->GetRange(L"A3")->SetValue(L"Feb");
	sheet->GetRange(L"A4")->SetValue(L"Mar");
	sheet->GetRange(L"A5")->SetValue(L"Apr");
	sheet->GetRange(L"A6")->SetValue(L"May");
	sheet->GetRange(L"A7")->SetValue(L"Jun");
	sheet->GetRange(L"B1")->SetValue(L"Peter");
	sheet->GetRange(L"B2")->SetNumberValue(25);
	sheet->GetRange(L"B3")->SetNumberValue(18);
	sheet->GetRange(L"B4")->SetNumberValue(8);
	sheet->GetRange(L"B5")->SetNumberValue(13);
	sheet->GetRange(L"B6")->SetNumberValue(22);
	sheet->GetRange(L"B7")->SetNumberValue(28);

	Chart* chart = sheet->GetCharts()->Add(ExcelChartType::LineMarkers);
	chart->SetDataRange(sheet->GetRange(L"B1:B7"));
	chart->GetPlotArea()->SetVisible(false);
	chart->SetSeriesDataFromRange(false);
	chart->SetTopRow(5);
	chart->SetBottomRow(26);
	chart->SetLeftColumn(2);
	chart->SetRightColumn(11);
	chart->SetChartTitle(L"Data Labels Demo");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);
	ChartSerie* cs1 = chart->GetSeries()->Get(0);
	cs1->SetCategoryLabels(sheet->GetRange(L"A2:A7"));

	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasLegendKey(false);
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasPercentage(false);
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasSeriesName(true);
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasCategoryName(true);
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetDelimiter(L". ");

	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetSize(9);
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetColor(Spire::Common::Color::GetRed());
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetFontName(L"Calibri");
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetPosition(DataLabelPositionType::Center);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}