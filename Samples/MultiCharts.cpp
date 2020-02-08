#include <cmath>
#include <cstdio>
#include <cstdlib>
#include <ctime>
#include <iostream>

#include <Xlsx/Chart.h>
#include <Xlsx/Chartsheet.h>
#include <Xlsx/Workbook.h>

using namespace SimpleXlsx;

int main()
{
    setlocale( LC_ALL, "" );

    time_t CurrentTime = time( NULL );
    srand( static_cast<unsigned int>( CurrentTime ) );

    CWorkbook Book( "Incognito" );

    //data sheet to store data to be referenced
    CWorksheet & FirstDataSheet = Book.AddSheet( "Data" );

    std::vector<CellDataDbl> data1, data2;
    CellDataDbl cellDbl;
    cellDbl.style_id = 0;

    const int DataCellCount = 35;
    // fill data sheet with randomly generated numbers
    for( int i = 0; i < DataCellCount; i++ )
    {
        cellDbl.value = ( rand() % 100 ) / 50.0;
        data1.push_back( cellDbl );
        cellDbl.value = 2.0 - cellDbl.value;
        data2.push_back( cellDbl );
    }

    FirstDataSheet.AddRow( data1 ).AddRow( data2 ); // data can be added by row or by cell

    // adding chart sheet to the workbook the reference to a newly created object is returned
    CChartsheet & ChartSheet = Book.AddChartSheet( "Line Chart", CHART_LINEAR );
    // create series object, that contains most chart settings
    CChart::Series ser;
    // leave category sequence (X axis) not specified (optional) - MS Excel will generate the default sequence automatically
    ser.catSheet = NULL;

    // specify range for values` sequence (Y axis)
    ser.valAxisFrom = CellCoord( 1, 0 );
    ser.valAxisTo = CellCoord( 1, DataCellCount );
    ser.valSheet =  & FirstDataSheet; // don`t forget to set the pointer to the data sheet

    ser.title = "First Line";
    ser.JoinType = CChart::Series::joinSmooth;

    // add series into the chart (you can add as many series as you wish into the same chart)
    ChartSheet.Chart().AddSeries( ser );

    // second line to the chart
    CChart::Series ser2 = ser;
    ser2.valAxisFrom = CellCoord( 2, 0 );
    ser2.valAxisTo = CellCoord( 2, DataCellCount );
    ser2.title = "Second Line";

    // charts in the worksheet
    CChart & OnSheetChart = Book.AddChart( FirstDataSheet, DrawingPoint( 3, 3 ), DrawingPoint( 10, 20 ) );
    OnSheetChart.AddSeries( ser );
    OnSheetChart.AddSeries( ser2 );
    OnSheetChart.SetLegendPos( CChart::POS_BOTTOM );

    CChart & OnSheetChart2 = Book.AddChart( FirstDataSheet, DrawingPoint( 13, 3 ), DrawingPoint( 20, 25 ) );
    OnSheetChart2.AddSeries( ser );
    ser2.Marker.Type = SimpleXlsx::CChart::Series::symDiamond;
    OnSheetChart2.AddSeries( ser2 );
    OnSheetChart2.SetLegendPos( CChart::POS_LEFT_ASIDE );

    CWorksheet & SecondDataSheet = Book.AddSheet( "Data2" );
    for( int i = 0; i < DataCellCount; i++ )
        SecondDataSheet.BeginRow().AddCell( sin( i * 0.5 ) + 1 ).EndRow();

    CChart::Series ser3;
    ser3.catSheet = NULL;
    ser3.valAxisFrom = CellCoord( 1, 0 );
    ser3.valAxisTo = CellCoord( DataCellCount, 0 );
    ser3.valSheet =  & SecondDataSheet; // don`t forget to set the pointer to the data sheet
    ser3.title = "Third Line";
    ser3.JoinType = CChart::Series::joinSmooth;
    ser3.LineColor = "00FF00";
    CChart & OnSheetChart3 = Book.AddChart( SecondDataSheet, DrawingPoint( 3, 3 ), DrawingPoint( 20, 25 ) );
    OnSheetChart3.AddSeries( ser3 );
    ChartSheet.Chart().AddSeries( ser3 );

    Book.SetActiveSheet( ChartSheet );

    if( Book.Save( "MultiCharts.xlsx" ) ) std::cout << "The book has been saved successfully" << std::endl;
    else std::cout << "The book saving has been failed" << std::endl;
    return 0;
}
