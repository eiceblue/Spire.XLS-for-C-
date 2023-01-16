#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"CreateBubbleChart.xlsx";
	wstring outputFile = output_path + L"CreateBubbleChart_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add a chart
	Chart* chart = sheet->GetCharts()->Add(ExcelChartType::Bubble);

	//Set region of chart data
	chart->SetDataRange(sheet->GetRange(L"A1:C5"));
	chart->SetSeriesDataFromRange(false);
	chart->GetSeries()->Get(0)->SetBubbles(sheet->GetRange(L"C2:C5"));
	//Set position of chart
	chart->SetLeftColumn(7);
	chart->SetTopRow(6);
	chart->SetRightColumn(16);
	chart->SetBottomRow(29);

	chart->SetChartTitle(L"Bubble Chart");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}