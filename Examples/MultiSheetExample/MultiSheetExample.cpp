#include "../../NumberDuck.hpp"

#include <stdio.h>

using namespace NumberDuck;

int main(int argc, char **argv)
{
	printf("MultiSheet Example\n");
	printf("Create a spreadsheet with multiple worksheets and cross-sheet formulas!\n\n");

	Workbook* pWorkbook = new Workbook(Workbook::License::AGPL);
	
	// Get the default worksheet and rename it
	Worksheet* pWorksheet = pWorkbook->GetWorksheetByIndex(0);
	pWorksheet->SetName("Sales Data");
	
	// Add some sales data to the first worksheet
	pWorksheet->GetCellByAddress("A1")->SetString("Month");
	pWorksheet->GetCellByAddress("B1")->SetString("Sales");
	pWorksheet->GetCellByAddress("C1")->SetString("Expenses");
	
	// Add monthly data
	const char* months[] = {"January", "February", "March", "April", "May", "June"};
	float sales[] = {12000.0f, 15000.0f, 18000.0f, 14000.0f, 16000.0f, 19000.0f};
	float expenses[] = {8000.0f, 9000.0f, 11000.0f, 8500.0f, 9500.0f, 12000.0f};
	
	for (int i = 0; i < 6; i++)
	{
		pWorksheet->GetCell(0, i + 1)->SetString(months[i]);
		pWorksheet->GetCell(1, i + 1)->SetFloat(sales[i]);
		pWorksheet->GetCell(2, i + 1)->SetFloat(expenses[i]);
	}
	
	// Create a second worksheet for calculations
	Worksheet* pCalcWorksheet = pWorkbook->CreateWorksheet();
	pCalcWorksheet->SetName("Calculations");
	
	// Add headers
	pCalcWorksheet->GetCellByAddress("A1")->SetString("Summary");
	pCalcWorksheet->GetCellByAddress("A2")->SetString("Total Sales:");
	pCalcWorksheet->GetCellByAddress("A3")->SetString("Total Expenses:");
	pCalcWorksheet->GetCellByAddress("A4")->SetString("Net Profit:");
	pCalcWorksheet->GetCellByAddress("A5")->SetString("Average Monthly Sales:");
	
	// Add cross-sheet formulas
	pCalcWorksheet->GetCellByAddress("B2")->SetFormula("=SUM('Sales Data'!B2:B7)");
	pCalcWorksheet->GetCellByAddress("B3")->SetFormula("=SUM('Sales Data'!C2:C7)");
	pCalcWorksheet->GetCellByAddress("B4")->SetFormula("=B2-B3");
	pCalcWorksheet->GetCellByAddress("B5")->SetFormula("=AVERAGE('Sales Data'!B2:B7)");
	
	// Create a third worksheet for charts
	Worksheet* pChartWorksheet = pWorkbook->CreateWorksheet();
	pChartWorksheet->SetName("Charts");
	
	// Add a title
	pChartWorksheet->GetCellByAddress("A1")->SetString("Sales vs Expenses Chart");
	
	// Create a chart using data from the first worksheet
	Chart* pChart = pChartWorksheet->CreateChart(Chart::TYPE_COLUMN);
	pChart->SetX(1);
	pChart->SetY(2);
	pChart->SetWidth(Worksheet::DEFAULT_COLUMN_WIDTH * 8);
	pChart->SetHeight(Worksheet::DEFAULT_ROW_HEIGHT * 12);
	
	// Set categories (months)
	pChart->SetCategories("='Sales Data'!A2:A7");
	
	// Create sales series
	Series* pSalesSeries = pChart->CreateSeries("='Sales Data'!B2:B7");
	pSalesSeries->SetName("='Sales Data'!B1");
	
	// Create expenses series
	Series* pExpensesSeries = pChart->CreateSeries("='Sales Data'!C2:C7");
	pExpensesSeries->SetName("='Sales Data'!C1");
	
	// Save the workbook
	pWorkbook->Save("MultiSheetExample.xls", Workbook::FileType::XLS);

	delete pWorkbook;
	return 0;
} 