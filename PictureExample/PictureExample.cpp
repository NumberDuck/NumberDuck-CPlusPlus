#include "../NumberDuck.hpp"

#include <stdio.h>
#include <string>

using namespace NumberDuck;

int main(int argc, char **argv)
{
	printf("Picture Example\n");

	// Dino images by Dave Catchpole
	// http://www.flickr.com/photos/yaketyyakyak/sets/72157629976688365/
	// http://creativecommons.org/licenses/by/2.0/deed.en

	printf("Embedding picture!\n");
	// Construct our workbook, and grab the default worksheet
	Workbook* pWorkbook = new Workbook(Workbook::License::AGPL);
	Worksheet* pWorksheet = pWorkbook->GetWorksheetByIndex(0);

	// Create the picture from our source image
	// Number Duck can create pictures from JPG and PNG files
	Picture* pPicture = pWorksheet->CreatePicture("Dino.jpg");

	// By default, the picture is located in the top left corner of the worksheet
	pPicture->SetX(2);
	pPicture->SetY(2);

	// X and Y refer to the cell, to position within the cell, set SubX and SubY
	pPicture->SetSubX(Worksheet::DEFAULT_COLUMN_WIDTH / 2);
	pPicture->SetSubY(Worksheet::DEFAULT_ROW_HEIGHT / 2);

	pWorkbook->Save("PictureExample.xls", Workbook::FileType::XLS);

	delete pWorkbook;


	printf("Extracting picture!\n");
	// Load the excel file with the image we want to extract
	pWorkbook = new Workbook(Workbook::License::AGPL);
	pWorkbook->Load("PictureExample.xls");
	pWorksheet = pWorkbook->GetWorksheetByIndex(0);

	pPicture = pWorksheet->GetPictureByIndex(0);

	// Work out the correct file name to save as based on the picture format
	// Note that while Number Duck can only create PNG and JPEG, it can extract any format
	std::string sFileName = "DinoOut.";
	switch (pPicture->GetFormat())
	{
		case Picture::FORMAT_DIB: sFileName += "bmp";
			break;
		case Picture::FORMAT_EMF: sFileName += "emf";
			break;
		case Picture::FORMAT_JPEG: sFileName += "jpg";
			break;
		case Picture::FORMAT_PICT: sFileName += "pict";
			break;
		case Picture::FORMAT_PNG: sFileName += "png";
			break;
		case Picture::FORMAT_TIFF: sFileName += "tiff";
			break;
		case Picture::FORMAT_WMF: sFileName += "wmf";
			break;
	}

	// now write to disk
	pPicture->GetBlob()->Save(sFileName.c_str());

	delete pWorkbook;

	return 0;
}
