#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"Template_Xls_1.xlsx";
	wstring inputImage = inputFolder + L"Background.png";
	wstring outputFile = outputFolder + L"SetImageOffsetOfChart_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	Worksheet* sheet1 = workbook->GetWorksheets()->Add(L"Contrast");

	//Add chart1 and background image to sheet1
	Chart* chart1 = sheet1->GetCharts()->Add(ExcelChartType::ColumnClustered);
	chart1->SetDataRange(sheet->GetRange(L"D1:E8"));
	chart1->SetSeriesDataFromRange(false);

	//Set chart position
	chart1->SetLeftColumn(1);
	chart1->SetTopRow(11);
	chart1->SetRightColumn(8);
	chart1->SetBottomRow(33);

	//Add picture as background
	chart1->GetChartArea()->GetFill()->CustomPicture(Image::FromFile(inputImage.c_str()), L"None");

	chart1->GetChartArea()->GetFill()->SetTile(false);

	//Set the image offset
	chart1->GetChartArea()->GetFill()->GetPicStretch()->SetLeft(20);
	chart1->GetChartArea()->GetFill()->GetPicStretch()->SetTop(20);
	chart1->GetChartArea()->GetFill()->GetPicStretch()->SetRight(5);
	chart1->GetChartArea()->GetFill()->GetPicStretch()->SetBottom(5);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}