#include <iostream>
#include <stdlib.h>
#include <time.h>
#include <math.h>
#include <Xlsx/Chart.h>
#include <Xlsx/Chartsheet.h>
#include <Xlsx/Workbook.h>
#include <XLSXColors/XLSXColorLib.h>

using namespace SimpleXlsx;

static void AddLineChart( CWorkbook & book, CWorksheet & sheet1,
                          XLSXColorLib & cl_lib, XLSXColorLib & gs10_lib )
{
    // create series object, that contains most chart settings
    CChart::Series ser;
    // adding chart to the workbook the reference to a newly created object is returned
    CChartsheet & chart_sheet = book.AddChartSheet( "Line Chart", CHART_LINEAR );
    // leave category sequence (X axis) not specified (optional) - MS Excel will generate the default sequence automatically
    ser.catSheet = NULL;

    // specify range for values` sequence (Y axis)
    ser.valSheet = & sheet1; // don`t forget to set the pointer to the data sheet
    ser.valAxisFrom = CellCoord( 1, 1 );
    ser.valAxisTo = CellCoord( 50, 1 );
    ser.JoinType = CChart::Series::joinSmooth;
    ser.DashType = CChart::Series::dashSolid;
    ser.LineColor = gs10_lib.GetColor( "Black" );

    ser.Marker.Type = CChart::Series::symCircle;
    ser.Marker.Size = 10;
    ser.Marker.FillColor = cl_lib.GetColor( "Gold" );
    ser.Marker.LineColor = cl_lib.GetColor( "Plum" );
    ser.Marker.LineWidth = 1.;
    ser.title = "Line series test";

    // add series into the chart (you can add as many series as you wish into the same chart)
    chart_sheet.Chart().AddSeries( ser );
}

static void AddScatterChart( CWorkbook & book, CWorksheet & sheet, CWorksheet & sheet1,
                             XLSXColorLib & cl_lib, XLSXColorLib & gs10_lib )
{
    // create series object, that contains most chart settings
    CChart::Series ser;
    // adding chart to the workbook the reference to a newly created object is returned
    CChartsheet & chart_sheet = book.AddChartSheet( "Scatter Chart", CHART_SCATTER );
    CChart & chart = chart_sheet.Chart();

    // for scatter charts it is obligatory to specify both category (X axis) and values (Y axis) sequences
    ser.catSheet = & sheet;
    ser.catAxisFrom = CellCoord( 1, 0 );
    ser.catAxisTo = CellCoord( 18, 0 );

    ser.valSheet = & sheet;
    ser.valAxisFrom = CellCoord( 1, 3 );
    ser.valAxisTo = CellCoord( 18, 3 );

    ser.JoinType = CChart::Series::joinSmooth;  // determines whether series will be a smoothed or straight-lined curve
    ser.Marker.Type = CChart::Series::symNone;   // if true add diamond marks in each node of the sequence set

    // add series into the chart (you can add as many series as you wish into the same chart)
    ser.title = "Log Line";
    chart.AddSeries( ser );

    ser.JoinType = CChart::Series::joinNone;
    ser.Marker.Type = CChart::Series::symCircle;
    ser.Marker.Size = 10;
    ser.Marker.FillColor = cl_lib.GetColor( "Gold" );
    ser.Marker.LineColor = cl_lib.GetColor( "Plum" );
    ser.Marker.LineWidth = 1.;

    // add series into the chart (you can add as many series as you wish into the same chart)
    ser.title = "Log Mark";
    chart.AddSeries( ser );

    ser.catAxisFrom = CellCoord( 1, 0 );
    ser.catAxisTo = CellCoord( 18, 0 );
    ser.catSheet = & sheet;

    ser.valAxisFrom = CellCoord( 1, 2 );
    ser.valAxisTo = CellCoord( 18, 2 );
    ser.valSheet = & sheet;

    ser.JoinType = CChart::Series::joinLine;
    ser.Marker.Type = CChart::Series::symNone;

    ser.DashType = CChart::Series::dashDot;

    // add series into the chart (you can add as many series as you wish into the same chart)
    ser.title = "sin Line";
    chart.AddSeries( ser );

    ser.JoinType = CChart::Series::joinNone;  // determines whether series will be joined, and if so- smoothed or straight-lined curve will be drawn
    ser.Marker.Type = CChart::Series::symSquare;    // if true add diamond marks in each node of the sequence set
    ser.Marker.Size = 3;
    ser.Marker.FillColor = gs10_lib.GetColor( "Gray80%" );
    ser.Marker.LineColor = cl_lib.GetColor( "Blue" );
    ser.Marker.LineWidth = 0.75;

    // add series into the chart (you can add as many series as you wish into the same chart)
    ser.title = "sin Mark";
    chart.AddSeries( ser );

    ser.catSheet = & sheet1;
    ser.catAxisFrom = CellCoord( 1, 0 );
    ser.catAxisTo = CellCoord( 50, 0 );

    ser.valSheet = & sheet1;
    ser.valAxisFrom = CellCoord( 1, 1 );
    ser.valAxisTo = CellCoord( 50, 1 );

    // optional parameters
    ser.JoinType = CChart::Series::joinLine;
    ser.Marker.Type = CChart::Series::symNone;
    ser.LineWidth = 0.5;
    ser.DashType = CChart::Series::dashSolid;
    ser.LineColor = gs10_lib.GetColor( "Black" );

    ser.title = "Cos Line";
    chart.AddSeries( ser );

    ser.Marker.Type = CChart::Series::symTriangle;
    ser.Marker.Size = 7;
    ser.Marker.FillColor = cl_lib.GetColor( "Lime" );
    ser.Marker.LineColor = cl_lib.GetColor( "Black" );
    ser.Marker.LineWidth = 0.33;
    ser.JoinType = CChart::Series::joinNone;

    // add series into the chart (you can add as many series as you wish into the same chart)
    ser.title = "Cos Mark";
    chart.AddSeries( ser );

    // optional parameters to set
    chart.SetLegendPos( CChart::POS_RIGHT );
    chart.SetXAxisGrid( CChart::GRID_MAJOR_N_MINOR );
    chart.SetYAxisGrid( CChart::GRID_MAJOR_N_MINOR );
    chart.SetYAxisCross( CChart::CROSS_MIN );
    chart.SetXAxisCross( CChart::CROSS_MIN );
    chart.SetXAxisName( L"\u03D1, sec" );
    chart.SetYAxisName( L"\u046C, \u03A9 /cm\u00B2 " );

    chart.SetDiagrammName( "Scatter curves` chart" );
}

static void AddBarChart( CWorkbook & book, CWorksheet & sheet )
{
    // create series object, that contains most chart settings
    CChart::Series ser;
    CChartsheet & chart_sheet = book.AddChartSheet( "Bar Chart", CHART_BAR );
    CChart & chart = chart_sheet.Chart();

    // leave category sequence (X axis) not specified (optional) - MS Excel will generate the default sequence automatically
    ser.catSheet = NULL;
    // specify range for values` sequence (Y axis)
    ser.valAxisFrom = CellCoord( 1, 2 );
    ser.valAxisTo = CellCoord( 18, 2 );
    ser.valSheet = & sheet; // don`t forget to set the pointer to the data sheet
    ser.title = "Bar series test";
    ser.dataLabels.showPercent = true;

    // optionally it is possible to set some additional parameters for bar chart
    chart.SetBarDirection( CChart::BAR_DIR_HORIZONTAL );
    chart.SetBarGrouping( CChart::BAR_GROUP_CLUSTERED );
    chart.SetTableDataState( CChart::TBL_DATA );
    chart.SetXAxisToCat(); // Switch X axis to cat - nothing about "categorical axis", just specific mode, default is false;
    // add series into the chart (you can add as many series as you wish into the same chart)
    chart.AddSeries( ser );
}

static void AddPieChart( CWorkbook & book, CWorksheet & sheet )
{
    // create series object, that contains most chart settings
    CChart::Series ser;
    CChartsheet & chart_sheet = book.AddChartSheet( "Pie Chart", CHART_PIE );
    CChart & chart = chart_sheet.Chart();

    // leave category sequence (X axis) not specified (optional) - MS Excel will generate the default sequence automatically
    ser.catSheet = NULL;
    // specify range for values` sequence (Y axis)
    ser.valAxisFrom = CellCoord( 1, 2 );
    ser.valAxisTo = CellCoord( 18, 2 );
    ser.valSheet = & sheet; // don`t forget to set the pointer to the data sheet
    ser.title = "Pie series test";
    ser.dataLabels.showPercent = true;

    chart.SetPieFirstSliceAng( 90 );
    // add series into the chart (you can add as many series as you wish into the same chart)
    chart.AddSeries( ser );
}

int main()
{
    // main object - workbook
    CWorkbook book( "me" );

    // data sheet to store data to be referenced
    CWorksheet & sheet = book.AddSheet( "Data" );
    CWorksheet & sheet1 = book.AddSheet( "Data1" );

    for( int i = 0; i < 20; i++ )
    {
        /*  addition cell-by-cell may be useful if there are different data types in one row
            to perform it the code will look like:*/
        sheet.BeginRow();
        sheet.AddCell( i );
        sheet.AddCell( i * i );
        sheet.AddCell( sin( i ) );
        sheet.AddCell( log( i + 1 ) );
        sheet.EndRow();
    }

    for( int i = 0; i < 50; i++ )
    {
        /*  addition cell-by-cell may be useful if there are different data types in one row
            to perform it the code will look like:*/
        sheet1.BeginRow().AddCell( i ).AddCell( cos( i ) ).EndRow();
    }
    // init  color libs.
    XLSXColorLib cl_lib, gs10_lib;
    make_excell_like_named_colors( cl_lib );
    make_grayscale10( gs10_lib );

    AddLineChart( book, sheet1, cl_lib, gs10_lib );
    AddScatterChart( book, sheet, sheet1, cl_lib, gs10_lib );
    AddBarChart( book, sheet );
    AddPieChart( book, sheet );

    // at the end save created workbook wherever you need
    if( book.Save( "Scientific.xlsx" ) ) std::cout << "The book has been saved successfully" << std::endl;
    else std::cout << "The book saving has been failed" << std::endl;
    return 0;
}
