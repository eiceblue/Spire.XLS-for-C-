#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputImg = input_path + L"SpireXls.png";
	wstring outputFile = output_path + L"InsertShapesToExcelSheet.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add a triangle shape.
	IPrstGeomShape* triangle = sheet->GetPrstGeomShapes()->AddPrstGeomShape(2, 2, 100, 100, PrstGeomShapeType::Triangle);
	//Fill the triangle with solid color.
	triangle->GetFill()->SetForeColor(Spire::Common::Color::GetYellow());
	triangle->GetFill()->SetFillType(ShapeFillType::SolidColor);

	//Add a heart shape.
	IPrstGeomShape* heart = sheet->GetPrstGeomShapes()->AddPrstGeomShape(2, 5, 100, 100, PrstGeomShapeType::Heart);
	//Fill the heart with gradient color.
	heart->GetFill()->SetForeColor(Spire::Common::Color::GetRed());
	heart->GetFill()->SetFillType(ShapeFillType::Gradient);

	//Add an arrow shape with default color.
	IPrstGeomShape* arrow = sheet->GetPrstGeomShapes()->AddPrstGeomShape(10, 2, 100, 100, PrstGeomShapeType::CurvedRightArrow);

	//Add a cloud shape.
	IPrstGeomShape* cloud = sheet->GetPrstGeomShapes()->AddPrstGeomShape(10, 5, 100, 100, PrstGeomShapeType::Cloud);
	//Fill the cloud with custom picture
	cloud->GetFill()->CustomPicture(Image::FromFile(inputImg.c_str()), L"SpireXls.png");

	cloud->GetFill()->SetFillType(ShapeFillType::Picture);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}