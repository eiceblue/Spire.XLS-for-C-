#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DrawOneLineThroughTwoPoints.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//1)Draw a line according to relative position
	XlsLineShape* line1 = dynamic_cast<XlsLineShape*>(sheet->GetTypedLines()->AddLine());
	line1->SetLeftColumn(3);
	line1->SetTopRow(3);
	line1->SetLeftColumnOffset(0);
	line1->SetTopRowOffset(0);

	line1->SetRightColumn(4);
	line1->SetBottomRow(5);
	line1->SetRightColumnOffset(0);
	line1->SetBottomRowOffset(0);

	//2)Draw a line according to absolute position(pixels).
	XlsLineShape* line2 = dynamic_cast<XlsLineShape*>(sheet->GetTypedLines()->AddLine());
	Point* startPoint = new Point();
	startPoint->SetX(30), startPoint->SetY(50);
	line2->SetStartPoint(startPoint);
	Point* endPoint = new Point();
	endPoint->SetX(20), endPoint->SetY(80);
	line2->SetEndPoint(endPoint);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
