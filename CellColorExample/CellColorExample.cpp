#include "../NumberDuck.hpp"
#include "stb_image.h"

#include <stdio.h>
#include <string>

using namespace NumberDuck;

int main(int argc, char **argv)
{
	printf("Cell Color Example\n");
	printf("Set cell size and background color!\n");

	// Load an image to get the colors we should assign to cells
	int nWidth;
	int nHeight;
	int nColorDepth;
	unsigned char* pImageData = stbi_load("Duck.png", &nWidth, &nHeight, &nColorDepth, 4);

	// the image could be larger than the maximum worksheet bounds, check for that to be safe
	if (nWidth > Worksheet::MAX_COLUMN || nHeight > Worksheet::MAX_ROW)
	{
		printf("Sorry, Duck.png is too large. Max dimensions: %dx%d\n", Worksheet::MAX_COLUMN, Worksheet::MAX_ROW);
		stbi_image_free(pImageData); // free image data
		return 1;
	}
	
	// Now we construct the workbook, and grab the default worksheet. Ready to start messing with cells
	Workbook * pWorkbook = new Workbook();
	Worksheet* pWorksheet = pWorkbook->GetWorksheetByIndex(0);

	// Now that we have our image and our worksheet we'll setup a nice square grid
	// Here we are setting the cell size to 7px by 7px.
	for (int nX = 0; nX < nWidth; nX++)
		pWorksheet->SetColumnWidth(nX, 7);

	for (int nY = 0; nY < nHeight; nY++)
		pWorksheet->SetRowHeight(nY, 7);

	for (int nY = 0; nY < nHeight; nY++)
	{
		for (int nX = 0; nX < nWidth; nX++)
		{
			// Get the data offset into the packed image data and read out the cell color there
			unsigned int nOffset = (nX + nY * nWidth)*4;

			unsigned char nRed = pImageData[nOffset+0];
			unsigned char nGreen = pImageData[nOffset+1];
			unsigned char nBlue = pImageData[nOffset+2];
			unsigned char nAlpha = pImageData[nOffset+3];

			// Skip transparent pixels
			if (nAlpha > 0)
			{
				Color color(nRed, nGreen, nBlue);

				Style* pStyle = NULL;
				// reuse the same style if it has the same background color
				for (int i = 0; i < pWorkbook->GetNumStyle(); i++)
				{
					Style* pTestStyle = pWorkbook->GetStyleByIndex(i);
					if (pTestStyle->GetBackgroundColor(false) && pTestStyle->GetBackgroundColor(false)->Equals(&color))
					{
						pStyle = pTestStyle;
						break;
					}
				}

				// reusable style not found, create a new one
				if (pStyle == NULL)
				{
					pStyle = pWorkbook->CreateStyle();
					pStyle->GetBackgroundColor(true)->SetFromColor(&color);
				}

				// set the cell style (color)
				Cell* pCell = pWorksheet->GetCell(nX, nY);
				pCell->SetStyle(pStyle);
			}
		}
	}
	stbi_image_free(pImageData); // free image data
	
	pWorkbook->Save("CellColorExample.xls", Workbook::FILE_TYPE_XLS);

	delete pWorkbook;
	return 0;
}

