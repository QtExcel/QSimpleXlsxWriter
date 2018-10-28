#include <iostream>
#include <stdlib.h>
#include <time.h>
#include <math.h>
#include <Xlsx/Xlsx/Chart.h>
#include <Xlsx/Chartsheet.h>
#include <Xlsx/Workbook.h>
#include <XLSXColors/XLSXColorLib.h>
using namespace std;
using namespace SimpleXlsx;

int main()
{
    // main object - workbook
    CWorkbook book( "me" );

    // data sheet to store data to be referenced
    CWorksheet & sheet = book.AddSheet( "Data" );
    CWorksheet & sheet1 = book.AddSheet( "Data1" );
    CellDataDbl cellDbl;
    cellDbl.style_id = 0;


    for( int i = 0; i < 20; i++ )
    {

        /*  addition cell-by-cell may be useful if there are different data types in one row
            to perform it the code will look like:*/
        sheet.BeginRow();
        cellDbl.value = i;
        sheet.AddCell( cellDbl );
        cellDbl.value = i * i;
        sheet.AddCell( cellDbl );
        cellDbl.value = sin( i );
        sheet.AddCell( cellDbl );
        cellDbl.value = log( i + 1 );
        sheet.AddCell( cellDbl );

        sheet.EndRow();
    }

    for( int i = 0; i < 50; i++ )
    {

        /*  addition cell-by-cell may be useful if there are different data types in one row
            to perform it the code will look like:*/
        sheet1.BeginRow();
        cellDbl.value = i;
        sheet1.AddCell( cellDbl );
        cellDbl.value = cos( i );
        sheet1.AddCell( cellDbl );
        sheet1.EndRow();
    }
    // init  color libs.
    XLSXColorLib cl_lib, gs10_lib;
    make_excell_like_named_colors( cl_lib );
    make_grayscale10( gs10_lib );

    // create series object, that contains most chart settings
    CChart::Series ser;

    // adding chart to the workbook the reference to a newly created object is returned
    CChartsheet & line_chart = book.AddChartSheet( "Line Chart", CHART_LINEAR );

    // leave category sequence (X axis) not specified (optional) - MS Excel will generate the default sequence automatically
    ser.catSheet =  NULL;

    // specify range for values` sequence (Y axis)
    ser.valSheet =  &sheet1; // don`t forget to set the pointer to the data sheet
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
    line_chart.Chart().AddSeries( ser );

    // adding chart to the workbook the reference to a newly created object is returned
    CChartsheet & scatter_chart = book.AddChartSheet( "Scatter Chart", CHART_SCATTER );

    // for scatter charts it is obligatory to specify both category (X axis) and values (Y axis) sequences
    ser.catSheet =  &sheet;
    ser.catAxisFrom = CellCoord( 1, 0 );
    ser.catAxisTo = CellCoord( 18, 0 );


    ser.valSheet =  &sheet;
    ser.valAxisFrom = CellCoord( 1, 3 );
    ser.valAxisTo = CellCoord( 18, 3 );




    ser.JoinType = CChart::Series::joinSmooth;  // determines whether series will be a smoothed or straight-lined curve
    ser.Marker.Type = CChart::Series::symNone;   // if true add diamond marks in each node of the sequence set




    // add series into the chart (you can add as many series as you wish into the same chart)
    ser.title = "Log Line";
    scatter_chart.Chart().AddSeries( ser );


    ser.JoinType = CChart::Series::joinNone;
    ser.Marker.Type = CChart::Series::symCircle;
    ser.Marker.Size = 10;
    ser.Marker.FillColor = cl_lib.GetColor( "Gold" );
    ser.Marker.LineColor = cl_lib.GetColor( "Plum" );
    ser.Marker.LineWidth = 1.;


    // add series into the chart (you can add as many series as you wish into the same chart)
    ser.title = "Log Mark";
    scatter_chart.Chart().AddSeries( ser );

    ser.catAxisFrom = CellCoord( 1, 0 );
    ser.catAxisTo = CellCoord( 18, 0 );
    ser.catSheet =  &sheet;

    ser.valAxisFrom = CellCoord( 1, 2 );
    ser.valAxisTo = CellCoord( 18, 2 );
    ser.valSheet =  &sheet;

    ser.JoinType = CChart::Series::joinLine;
    ser.Marker.Type = CChart::Series::symNone;

    ser.DashType = CChart::Series::dashDot;
    //scatter_chart.SetScatterStyle(CChartsheet::SCATTER_FILL);

    // add series into the chart (you can add as many series as you wish into the same chart)
    ser.title = "sin Line";
    scatter_chart.Chart().AddSeries( ser );


    ser.JoinType = CChart::Series::joinNone;  // determines whether series will be joined, and if so- smoothed or straight-lined curve will be drawn
    ser.Marker.Type = CChart::Series::symSquare;    // if true add diamond marks in each node of the sequence set
    ser.Marker.Size = 3;
    ser.Marker.FillColor = gs10_lib.GetColor( "Gray80%" );
    ser.Marker.LineColor = cl_lib.GetColor( "Blue" );
    ser.Marker.LineWidth = 0.75;


    // add series into the chart (you can add as many series as you wish into the same chart)
    ser.title = "sin Mark";
    scatter_chart.Chart().AddSeries( ser );

    ser.catSheet =  &sheet1;
    ser.catAxisFrom = CellCoord( 1, 0 );
    ser.catAxisTo = CellCoord( 50, 0 );

    ser.valSheet =  &sheet1;
    ser.valAxisFrom = CellCoord( 1, 1 );
    ser.valAxisTo = CellCoord( 50, 1 );


    // optional parameters
    ;
    ser.JoinType = CChart::Series::joinLine;
    ser.Marker.Type = CChart::Series::symNone;
    ser.LineWidth = 0.5;
    //ser.DashType=CChart::Series::dashShortDash;
    ser.DashType = CChart::Series::dashSolid;
    ser.LineColor = gs10_lib.GetColor( "Black" );


    ser.title = "Cos Line";
    scatter_chart.Chart().AddSeries( ser );


    ser.Marker.Type = CChart::Series::symTriangle;
    ser.Marker.Size = 7;
    ser.Marker.FillColor = cl_lib.GetColor( "Lime" );
    ser.Marker.LineColor = cl_lib.GetColor( "Black" );
    ser.Marker.LineWidth = 0.33;
    ser.JoinType = CChart::Series::joinNone;
    //scatter_chart.SetScatterStyle(CChartsheet::SCATTER_FILL);

    // add series into the chart (you can add as many series as you wish into the same chart)
    ser.title = "Cos Mark";
    scatter_chart.Chart().AddSeries( ser );



    // optional parameters to set
    scatter_chart.Chart().SetLegendPos( CChart::POS_RIGHT );
    scatter_chart.Chart().SetXAxisGrid( CChart::GRID_MAJOR_N_MINOR );
    scatter_chart.Chart().SetYAxisGrid( CChart::GRID_MAJOR_N_MINOR );
    scatter_chart.Chart().SetYAxisCross( CChart::CROSS_MIN );
    scatter_chart.Chart().SetXAxisCross( CChart::CROSS_MIN );
    scatter_chart.Chart().SetXAxisName( "\u03D1, sec" );
    scatter_chart.Chart().SetYAxisName( "\u046C, \u03A9 /cm\u00B2 " );

    scatter_chart.Chart().SetDiagrammName( "Scatter curves` chart" );

    // at the end save created workbook wherever you need
    CChartsheet & bar_chart = book.AddChartSheet( "Bar Chart", CHART_BAR );

    // leave category sequence (X axis) not specified (optional) - MS Excel will generate the default sequence automatically
    ser.catSheet =  NULL;

    // specify range for values` sequence (Y axis)

    ser.valAxisFrom = CellCoord( 1, 2 );
    ser.valAxisTo = CellCoord( 18, 2 );
    ser.valSheet =  &sheet; // don`t forget to set the pointer to the data sheet
    ser.title = "Bar series test";

    // optionally it is possible to set some additional parameters for bar chart
    bar_chart.Chart().SetBarDirection( CChart::BAR_DIR_HORIZONTAL );
    bar_chart.Chart().SetBarGrouping( CChart::BAR_GROUP_CLUSTERED );

    bar_chart.Chart().SetTableDataState( CChart::TBL_DATA );

    bar_chart.Chart().SetXAxisToCat(); // Switch X axis to cat - nothing about "categorical axis", just specific mode, default is false;


    // add series into the chart (you can add as many series as you wish into the same chart)
    bar_chart.Chart().AddSeries( ser );




    if( book.Save( "Scientific.xlsx" ) ) std::cout << "The book has been saved successfully" << std::endl;
    else std::cout << "The book saving has been failed" << std::endl;


    return 0;
}