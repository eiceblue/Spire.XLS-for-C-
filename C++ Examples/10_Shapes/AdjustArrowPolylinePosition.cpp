#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AdjustArrowPolylinePosition.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Draw an elbow arrow
	XlsLineShape* line = dynamic_cast<XlsLineShape*>(sheet->GetTypedLines()->AddLine(5, 5, 100, 100, LineShapeType::ElbowLine));
	line->SetEndArrowHeadStyle(ShapeArrowStyleType::LineNoArrow);
	line->SetBeginArrowHeadStyle(ShapeArrowStyleType::LineArrow);
	GeomertyAdjustValue* ad = line->GetShapeAdjustValues()->AddAdjustValue(GeomertyAdjustValueFormulaType::LiteralValue);

	//When the parameter value is less than 0, the focus of the line is on the left side of the left point, when it is equal to 0, the position is the same as the left point, it is equal to 50 in the middle of the graph, and when it is equal to 100, it is the same as the right point.
	ad->SetFormulaParameter(std::vector<double> {-50});

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
