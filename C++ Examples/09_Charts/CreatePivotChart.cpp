#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"PivotTable.xlsx";
	wstring outputFile = output_path + L"CreatePivotChart.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//get the first pivot table in the worksheet
	IPivotTable* pivotTable = sheet->GetPivotTables()->Get(0);

	//create a clustered column chart based on the pivot table
	Chart* chart = sheet->GetCharts()->Add(ExcelChartType::ColumnClustered, pivotTable);
	//set chart position
	chart->SetTopRow(10);
	chart->SetLeftColumn(1);
	chart->SetRightColumn(7);
	chart->SetBottomRow(25);
	//set chart title
	chart->SetChartTitle(L"Pivot Chart");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}