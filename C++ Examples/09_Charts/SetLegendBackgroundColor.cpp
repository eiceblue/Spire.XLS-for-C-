#include "pch.h"
using namespace Spire::Xls;

int main() {
                wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ChartSample1.xlsx";
	wstring outputFile = output_path + L"SetLegendBackgroundColor.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	Chart* chart = sheet->GetCharts()->Get(0);

	XlsChartFrameFormat* x = dynamic_cast<XlsChartFrameFormat*>(chart->GetLegend()->GetFrameFormat());
	x->GetFill()->SetFillType(ShapeFillType::SolidColor);
	x->SetForeGroundColor(Spire::Common::Color::GetSkyBlue());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
