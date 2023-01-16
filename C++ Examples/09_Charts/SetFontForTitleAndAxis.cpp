#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring  inputFile = input_path + L"ChartSample1.xlsx";
	wstring  outputFile = output_path + L"SetFontForTitleAndAxis.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	Chart* chart = sheet->GetCharts()->Get(0);

	//Format the font for the chart title
	chart->GetChartTitleArea()->SetColor(Spire::Common::Color::GetBlue());
	chart->GetChartTitleArea()->SetSize(20.0);

	//Format the font for the chart Axis
	chart->GetPrimaryValueAxis()->GetFont()->SetColor(Spire::Common::Color::GetGold());
	chart->GetPrimaryValueAxis()->GetFont()->SetSize(10.0);
	chart->GetPrimaryCategoryAxis()->GetFont()->SetColor(Spire::Common::Color::GetRed());
	chart->GetPrimaryCategoryAxis()->GetFont()->SetSize(20.0);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
