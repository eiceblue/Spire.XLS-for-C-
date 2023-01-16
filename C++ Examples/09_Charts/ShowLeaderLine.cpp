#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ShowLeaderLine.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set value of specified range
	sheet->GetRange(L"A1")->SetValue(L"1");
	sheet->GetRange(L"A2")->SetValue(L"2");
	sheet->GetRange(L"A3")->SetValue(L"3");
	sheet->GetRange(L"B1")->SetValue(L"4");
	sheet->GetRange(L"B2")->SetValue(L"5");
	sheet->GetRange(L"B3")->SetValue(L"6");
	sheet->GetRange(L"C1")->SetValue(L"7");
	sheet->GetRange(L"C2")->SetValue(L"8");
	sheet->GetRange(L"C3")->SetValue(L"9");

	Chart* chart = sheet->GetCharts()->Add(ExcelChartType::BarStacked);
	chart->SetDataRange(sheet->GetRange(L"A1:C3"));
	chart->SetTopRow(4);
	chart->SetLeftColumn(2);
	chart->SetWidth(450);
	chart->SetHeight(300);

	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		ChartSerie* cs = chart->GetSeries()->Get(i);
		cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
		cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetShowLeaderLines(true);
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
	
}
