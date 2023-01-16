#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartSample3.xlsx";
	wstring outputFile = output_path + L"SetBorderColorAndStyle.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	Chart* chart = sheet->GetCharts()->Get(0);

	//Set CustomLineWeight property for Series line
	(dynamic_cast<XlsChartBorder*>(chart->GetSeries()->Get(0)->GetDataPoints()->Get(0)->GetDataFormat()->GetLineProperties()))->SetCustomLineWeight(2.5f);
	//Set color property for Series line
	(dynamic_cast<XlsChartBorder*>(chart->GetSeries()->Get(0)->GetDataPoints()->Get(0)->GetDataFormat()->GetLineProperties()))->SetColor(Spire::Common::Color::GetRed());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}