#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ShapeToImage.xlsx";
	wstring outputFile = output_path + L"ShapeToImage.png";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	sheet->GetPrstGeomShapes()->Get(0);

	//Get the first shape from the first worksheet
	XlsShape* shape = dynamic_cast<XlsShape*>(sheet->GetPrstGeomShapes()->Get(0));
	
	//Save the shape to a image
	Image* img = shape->SaveToImage();
	img->Save(outputFile.c_str(), ImageFormat::GetPng());
}
