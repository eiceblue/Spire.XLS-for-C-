#include "pch.h"
using namespace Spire::Xls;

int main() {	
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"CreateRadarChart.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Initailize worksheet
	workbook->CreateEmptySheets(1);
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	sheet->SetName(L"Chart data");
	sheet->SetGridLinesVisible(false);

	//Writes chart data
	//Product
	sheet->GetRange(L"A1")->SetValue(L"Product");
	sheet->GetRange(L"A2")->SetValue(L"Bikes");
	sheet->GetRange(L"A3")->SetValue(L"Cars");
	sheet->GetRange(L"A4")->SetValue(L"Trucks");
	sheet->GetRange(L"A5")->SetValue(L"Buses");

	//Paris
	sheet->GetRange(L"B1")->SetValue(L"Paris");
	sheet->GetRange(L"B2")->SetNumberValue(4000);
	sheet->GetRange(L"B3")->SetNumberValue(23000);
	sheet->GetRange(L"B4")->SetNumberValue(4000);
	sheet->GetRange(L"B5")->SetNumberValue(30000);

	//New York
	sheet->GetRange(L"C1")->SetValue(L"New York");
	sheet->GetRange(L"C2")->SetNumberValue(30000);
	sheet->GetRange(L"C3")->SetNumberValue(7600);
	sheet->GetRange(L"C4")->SetNumberValue(18000);
	sheet->GetRange(L"C5")->SetNumberValue(8000);

	//Style
	sheet->GetRange(L"A1:C1")->GetStyle()->GetFont()->SetIsBold(true);
	sheet->GetRange(L"A2:C2")->GetStyle()->SetKnownColor(ExcelColors::LightYellow);
	sheet->GetRange(L"A3:C3")->GetStyle()->SetKnownColor(ExcelColors::LightGreen1);
	sheet->GetRange(L"A4:C4")->GetStyle()->SetKnownColor(ExcelColors::LightOrange);
	sheet->GetRange(L"A5:C5")->GetStyle()->SetKnownColor(ExcelColors::LightTurquoise);

	//Border
	sheet->GetRange(L"A1:C5")->GetStyle()->GetBorders()->Get(BordersLineType::EdgeTop)->SetColor(Spire::Common::Color::FromArgb(0, 0, 128));
	sheet->GetRange(L"A1:C5")->GetStyle()->GetBorders()->Get(BordersLineType::EdgeTop)->SetLineStyle(LineStyleType::Thin);
	sheet->GetRange(L"A1:C5")->GetStyle()->GetBorders()->Get(BordersLineType::EdgeBottom)->SetColor(Spire::Common::Color::FromArgb(0, 0, 128));
	sheet->GetRange(L"A1:C5")->GetStyle()->GetBorders()->Get(BordersLineType::EdgeBottom)->SetLineStyle(LineStyleType::Thin);
	sheet->GetRange(L"A1:C5")->GetStyle()->GetBorders()->Get(BordersLineType::EdgeLeft)->SetColor(Spire::Common::Color::FromArgb(0, 0, 128));
	sheet->GetRange(L"A1:C5")->GetStyle()->GetBorders()->Get(BordersLineType::EdgeLeft)->SetLineStyle(LineStyleType::Thin);
	sheet->GetRange(L"A1:C5")->GetStyle()->GetBorders()->Get(BordersLineType::EdgeRight)->SetColor(Spire::Common::Color::FromArgb(0, 0, 128));
	sheet->GetRange(L"A1:C5")->GetStyle()->GetBorders()->Get(BordersLineType::EdgeRight)->SetLineStyle(LineStyleType::Thin);

	sheet->GetRange(L"B2:C5")->GetStyle()->SetNumberFormat(L"\"$\"#,##0");
	//Add a new  chart worsheet to workbook
	Chart* chart = sheet->GetCharts()->Add();

	//Set position of chart
	chart->SetLeftColumn(1);
	chart->SetTopRow(6);
	chart->SetRightColumn(11);
	chart->SetBottomRow(29);

	//Set region of chart data
	chart->SetDataRange(sheet->GetRange(L"A1:C5"));
	chart->SetSeriesDataFromRange(false);

	chart->SetChartType(ExcelChartType::Radar);

	//Chart title
	chart->SetChartTitle(L"Sale market by region");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);

	chart->GetPlotArea()->GetFill()->SetVisible(false);

	chart->GetLegend()->SetPosition(LegendPositionType::Corner);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}