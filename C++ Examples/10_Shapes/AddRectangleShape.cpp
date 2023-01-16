#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ExcelSample_N1.xlsx";
	wstring outputFile = output_path + L"AddRectangleShape.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add rectangle shape 1------Rect
	IRectangleShape* rect1 = sheet->GetRectangleShapes()->AddRectangle(11, 2, 60, 100, RectangleShapeType::Rect);
	rect1->GetLine()->SetWeight(1);
	//Fill shape with solid color
	rect1->GetFill()->SetFillType(ShapeFillType::SolidColor);
	rect1->GetFill()->SetForeColor(Spire::Common::Color::GetDarkGreen());

	//Add rectangle shape 2------RoundRect
	IRectangleShape* rect2 = sheet->GetRectangleShapes()->AddRectangle(11, 5, 60, 100, RectangleShapeType::RoundRect);
	rect2->GetLine()->SetWeight(1);
	rect2->GetFill()->SetFillType(ShapeFillType::SolidColor);
	rect2->GetFill()->SetForeColor(Spire::Common::Color::GetDarkCyan());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();

}
