#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChangeSeriesColor.xlsx";
	wstring outputFile = output_path + L"ChangeSeriesColor_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the first chart
	Chart* chart = sheet->GetCharts()->Get(0);

	//Get the second series
	ChartSerie* cs = chart->GetSeries()->Get(1);

	//Set the fill type
	cs->GetFormat()->GetFill()->SetFillType(ShapeFillType::SolidColor);

	//Change the fill color
	cs->GetFormat()->GetFill()->SetForeColor(Spire::Common::Color::GetOrange());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}