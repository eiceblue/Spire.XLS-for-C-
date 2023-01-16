#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SetShapeOrder.xlsx";
	wstring outputFile = output_path + L"SetShapeOrder.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Bring the picture forward one level
	workbook->GetWorksheets()->Get(0)->GetPictures()->Get(0)->ChangeLayer(ShapeLayerChangeType::BringForward);

	//Bring the image in fron of all other objects
	workbook->GetWorksheets()->Get(1)->GetPictures()->Get(0)->ChangeLayer(ShapeLayerChangeType::BringToFront);

	//Send the shape back one level
	XlsShape* shape = dynamic_cast<XlsShape*>(workbook->GetWorksheets()->Get(2)->GetPrstGeomShapes()->Get(1));
	shape->ChangeLayer(ShapeLayerChangeType::SendBackward);

	//Send the shape behind all other objects
	shape = dynamic_cast<XlsShape*>(workbook->GetWorksheets()->Get(3)->GetPrstGeomShapes()->Get(1));
	shape->ChangeLayer(ShapeLayerChangeType::SendToBack);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
