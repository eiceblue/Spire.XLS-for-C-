#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SetFont.xlsx";
	wstring outputFile = output_path + L"SetFont.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the first sheet
	Chart* chart = sheet->GetCharts()->Get(0);

	//Create a font
	ExcelFont* font = workbook->CreateExcelFont();
	font->SetSize(15.0);
	font->SetColor(Spire::Common::Color::GetLightSeaGreen());
	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		ChartSerie* cs = chart->GetSeries()->Get(i);
		//Set font
		(dynamic_cast<XlsChartDataLabels*>(cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()))->GetTextArea()->SetFont(font);
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}