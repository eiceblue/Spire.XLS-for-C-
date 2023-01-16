#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartSample4.xlsx";
	wstring outputFile = output_path + L"ExtractTrendline.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the chart from the first worksheet
	Chart* chart = workbook->GetWorksheets()->Get(0)->GetCharts()->Get(0);

	//Get the trendline of the chart and then extract the equation of the trendline
	IChartTrendLine* trendLine = chart->GetSeries()->Get(1)->GetTrendLines()->GetItem(0);
	wstring formula = trendLine->GetFormula();
	wstring* content = new wstring();
	content->append(L"The equation is: " + formula + L"\r\n");

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}