#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartToImage.xlsx";
	wstring outputFile = output_path + L"GetChartDataPointValues.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the chart
	Chart* chart = sheet->GetCharts()->Get(0);

	//Get the first series of the chart
	ChartSerie* cs = chart->GetSeries()->Get(0);
	wstring* content = new wstring();
	for (int i = 0; i < cs->GetValues()->GetCells()->GetCount(); i++)
	{
		wstring address = cs->GetValues()->GetCells()->GetItem(i)->GetRangeAddress();
		content->append(address + L"\r\n");

		//Get the data point value
		wstring value = cs->GetValues()->GetCells()->GetItem(i)->GetValue();
		content->append(L"The value of the data point is " + value + L"\r\n");
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}