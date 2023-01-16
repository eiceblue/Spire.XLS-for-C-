#include "pch.h"
using namespace Spire::Xls;

int main() {
                wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ExcelSample_N1.xlsx";
	wstring outputFile = output_path + L"AddOvalShape.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add oval shape1
	IOvalShape* ovalShape1 = sheet->GetOvalShapes()->AddOval(11, 2, 100, 100);
	ovalShape1->GetLine()->SetWeight(0);
	//Fill shape with solid color
	ovalShape1->GetFill()->SetFillType(ShapeFillType::SolidColor);
	ovalShape1->GetFill()->SetForeColor(Spire::Common::Color::GetDarkCyan());

	//Add oval shape2
	IOvalShape* ovalShape2 = sheet->GetOvalShapes()->AddOval(11, 5, 100, 100);
	ovalShape2->GetLine()->SetWeight(1);
	//Fill shape with picture
	ovalShape2->GetLine()->SetDashStyle(ShapeDashLineStyleType::Solid);
	ovalShape2->GetFill()->CustomPicture((input_path + L"Logo.png").c_str());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
