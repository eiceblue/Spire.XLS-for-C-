#include "pch.h"
using namespace Spire::Xls;

int main() {
    wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ChartSample1.xlsx";
	wstring outputFile = output_path + L"SetFontForLegendAndDataTable.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	Chart* chart = sheet->GetCharts()->Get(0);

	//Create a font with specified size and color
	ExcelFont* font = workbook->CreateExcelFont();
	font->SetSize(14.0);
	font->SetColor(Spire::Common::Color::GetRed());

	//Apply the font to chart Legend
	(dynamic_cast<ChartTextArea*>(chart->GetLegend()->GetTextArea()))->SetFont(font);

	//Apply the font to chart DataLabel
	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		ChartSerie* cs = chart->GetSeries()->Get(i);
		(dynamic_cast<XlsChartDataLabels*>(cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()))->GetTextArea()->SetFont(font);
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
