/*
  SimpleXlsxWriter
  Copyright (C) 2012-2018 Pavel Akimov <oxod.pavel@gmail.com>, Alexandr Belyak <programmeralex@bk.ru>

  This software is provided 'as-is', without any express or implied
  warranty. In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.
*/

#ifndef XLSX_CCHART_H
#define XLSX_CCHART_H

#include "SimpleXlsxDef.h"

// ****************************************************************************
/// @brief  The following namespace contains declarations of a number of classes
///         which allow writing .xlsx files with formulae and charts
/// @note   All classes inside the namepace are used together, so a separate
///         using will not guarantee stability and reliability
// ****************************************************************************
namespace SimpleXlsx
{
    class CWorksheet;

    class PathManager;
    class XMLWriter;

    class CChart
    {
        public:
            /// @brief  General property
            enum EPosition
            {
                POS_NONE = 0,
                POS_LEFT,
                POS_RIGHT,
                POS_TOP,
                POS_BOTTOM,
                POS_LEFT_ASIDE,
                POS_RIGHT_ASIDE,
                POS_TOP_ASIDE,
                POS_BOTTOM_ASIDE
            };

            /// @brief  Axis grid detalization property
            enum EGridLines
            {
                GRID_NONE = 0,
                GRID_MAJOR,
                GRID_MINOR,
                GRID_MAJOR_N_MINOR
            };

            /// @brief  The enumeration contains values that determine diagramm diagram data table property
            enum ETableData
            {
                TBL_DATA_NONE = 0,
                TBL_DATA,
                TBL_DATA_N_KEYS
            };

            /// @brief  The enumeration contains values that demermine point where axis will pass through
            enum EAxisCross
            {
                CROSS_AUTO_ZERO = 0,
                CROSS_MIN,
                CROSS_MAX
            };

            /// @brief	The enumeration determines bars` direction in bar diagramm
            enum EBarDirection
            {
                BAR_DIR_VERTICAL = 0,
                BAR_DIR_HORIZONTAL
            };

            /// @brief	The enumeration determinese several bar relative position
            enum EBarGrouping
            {
                BAR_GROUP_CLUSTERED = 0,
                BAR_GROUP_STACKED,
                BAR_GROUP_PERCENT_STACKED
            };

            /// @brief	The enumeration determines scatter diagram styles
            enum EScatterStyle
            {
                SCATTER_FILL = 0,
                SCATTER_POINT
            };

            /// @brief The enumeration determines empty cell display method
            enum EEmptyCellsDisplayMethod
            {
                EMPTY_CELLS_DISP_GAPS = 0,
                EMPTY_CELLS_DISP_ZERO,
                EMPTY_CELLS_DISP_CONNECT
            };

            /// @brief The enumeration determines plot and chart area fill style
            enum EPlotAreaFillStyle
            {
                PLOT_AREA_FILL_NONE = 0,
                PLOT_AREA_FILL_SOLID,
                PLOT_AREA_FILL_GRADIENT,
                //PLOT_AREA_FILL_PICTURE,
                PLOT_AREA_FILL_PATTERN
            };

            /// @brief The enumeration determines gradient fill direction (for radial and rectangular types)
            enum EGradientFillDirection
            {
                FROM_BOTTOM_RIGHT_CORNER, FROM_BOTTOM_LEFT_CORNER, FROM_CENTER,
                FROM_TOP_RIGHT_CORNER, FROM_TOP_LEFT_CORNER
            };

            /// @brief The enumeration determines pattern fill style
            enum EPatternFillStyle
            {
                PERCENT_5, PERCENT_10, PERCENT_20, PERCENT_25, PERCENT_30, PERCENT_40,
                PERCENT_50, PERCENT_60, PERCENT_70, PERCENT_75, PERCENT_80, PERCENT_90,

                LIGHT_DOWNWARD_DIAGONAL, DARK_DOWNWARD_DIAGONAL, WIDE_DOWNWARD_DIAGONAL,
                LIGHT_UPWARD_DIAGONAL, DARK_UPWARD_DIAGONAL, WIDE_UPWARD_DIAGONAL,

                LIGHT_VERTICAL, NARROW_VERTICAL, DARK_VERTICAL, LIGHT_HORIZONTAL, NARROW_HORIZONTAL, DARK_HORIZONTAL,

                DASHED_DOWNWARD_DIAGONAL, DASHED_UPWARD_DIAGONAL, DASHED_HORIZONTAL, DASHED_VERTICAL, SMALL_CONFETTI, LARGE_CONFETTI,

                ZIG_ZAG, WAVE, DIAGONAL_BRICK, HORIZONTAL_BRICK, WEAVE, PLAID,

                DIVOT, DOTTED_GRID, DOTTED_DIAMOND, SHINGLE, TRELLIS, SPHERE,

                SMALL_GRID, LARGE_GRID, SMALL_CHECKER_BOARD, LARGE_CHECKER_BOARD, OUTLINED_DIAMOND, SOLID_DIAMOND
            };

            /// @brief  Structure with color points to setup a gradient fill of chart
            struct GradientStops
            {
                private:
                    // int - Pos in Gradient stops from 0 to 100 in percent
                    // std::string - Color RGB string like "FF00FF"
                    typedef typename std::map< int, std::string > Container;

                    Container m_Points;

                public:
                    typedef typename Container::const_iterator const_iterator;

                    // *INDENT-OFF*   For AStyle tool
                    inline const_iterator begin() const { return m_Points.begin(); }
                    inline const_iterator end() const   { return m_Points.end(); }

                    inline bool IsValid() const         { return m_Points.size() >= 2; }
                    inline void Clear()                 { m_Points.clear(); }
                    // *INDENT-ON*   For AStyle tool

                    inline void Add( int Percent, const std::string & Color )
                    {
                        assert( ( Percent >= 0 ) && ( Percent <= 100 ) );
                        m_Points.insert( Container::value_type( Percent, Color ) );
                    }
            };

            /// @brief  Structure to setup a chart
            /// @note	category sheet parameters can left unset, then a default category axis is used
            /// @note	if values` sheet is unset the series will not be added to the list of charts
            struct Series
            {
                enum symType {symNone, symDiamond, symCircle, symSquare, symTriangle} ;
                enum joinType {joinNone, joinLine, joinSmooth} ;
                enum dashType {dashSolid, dashDot, dashShortDash, dashDash} ;
                struct stMarker
                {
                    symType Type;			     				///< SymType indicates whether and how nodes are marked
                    size_t Size;                              ///< SymSize 1 ... 10(?)
                    std::string FillColor, LineColor;         ///< Color RGB string like "FF00FF"
                    double LineWidth;                         ///< Like in excell 0.5 ... 3.0(?)
                    stMarker()
                    {
                        Size = 7;
                        Type = symNone;
                        FillColor = "";
                        LineColor = "";
                        LineWidth = 0.5;
                    }
                } Marker;
                CWorksheet * catSheet;							///< catSheet pointer to a sheet contains category axis
                CellCoord catAxisFrom;							///< catAxisFrom data range (from cell (row, col), beginning since 0)
                CellCoord catAxisTo;							///< catAxisTo data range (to cell (row, col), beginning since 0)
                CWorksheet * valSheet;							///< valSheet pointer to a sheet contains values axis
                CellCoord valAxisFrom;							///< valAxisFrom data range (from cell (row, col), beginning since 0)
                CellCoord valAxisTo;							///< valAxisTo data range (to cell (row, col), beginning since 0)

                UniString title;								///< title series title (in the legend)

                joinType JoinType;						     	///< JoinType indicates whether series must be joined and smoothed at rendering
                double LineWidth;                               ///< Like in excell  0.5 ... 3.0 (?)
                std::string LineColor;                          ///< LineGolor RGB string like "FF00FF"
                dashType DashType;							    ///< DashType indicates whether line will be rendered dashed or not

                // Specific for bar chart
                enum BarFillStyle
                {
                    BAR_FILL_NONE = 0,
                    BAR_FILL_SOLID,
                    BAR_FILL_AUTOMATIC
                };
                BarFillStyle    barFillStyle;
                bool            barInvertIfNegative;            ///< Invert fill if negative data value
                std::string     barInvertedColor;               ///< Golor RGB string like "FF00FF" if barInvertIfNegative = true


                Series()
                {
                    catSheet = NULL;
                    valSheet = NULL;
                    LineColor = "";
                    JoinType = joinNone;
                    LineWidth = 1.;
                    DashType = dashSolid;

                    barFillStyle = BAR_FILL_AUTOMATIC;
                    barInvertIfNegative = true;
                    barInvertedColor = "";
                }
            };

        private:
            /// @brief  Structure that describes axis properties
            /// @note	gridLines and sourceLinked properties will have no effect at using with additional axes because of Microsoft Open XML format restrictions
            struct Axis
            {
                uint32_t id;			///< axis axis id
                UniString name;			///< name axis name (that will be depicted)
                uint32_t nameSize;		///< nameSize font size
                EPosition pos;			///< pos axis position
                EGridLines gridLines;	///< gridLines grid detalization
                EAxisCross cross;		///< cross determines axis position relative to null
                bool sourceLinked;		///< sourceLinked indicates if axis has at least one linked source
                std::string minValue;	///< minValue minimum value for axis (type string is used for generality and processing simplicity)
                std::string maxValue;	///< maxValue minimum value for axis (type string is used for generality and processing simplicity)
                int lblSkipInterval;	///< space between two neighbour labels
                int markSkipInterval;	///< space between two neighbour marks
                int lblAngle;			///< axis labels angle in degrees
                bool isVal;            ///< "Val" axis is normal for scatter and line plots
                Axis()
                {
                    id = 0;
                    nameSize = 10;
                    pos = POS_LEFT;
                    gridLines = GRID_NONE;
                    cross = CROSS_AUTO_ZERO;
                    sourceLinked = false;

                    minValue = "";
                    maxValue = "";

                    lblSkipInterval = -1;	// auto
                    markSkipInterval = -1;	// auto
                    lblAngle = -1;			// none
                    isVal = true;
                }
            };

            /// @brief  Structure to setup a gradient fill of chart
            struct GradientFill
            {
                enum fillType { linear, radial, rectangular, path };

                fillType        FillType;
                EGradientFillDirection   FillDirection;
                double          LinearAngle;        // 0 ... 359.9 degree (clockwise)
                bool            LinearScaleAngle;   // Specifies whether the angle scales with the fill region

                GradientStops   ColorPoints;

                GradientFill()
                {
                    FillType = linear;
                    FillDirection = FROM_CENTER;
                    LinearAngle = 0;
                    LinearScaleAngle = true;
                }
            };

            /// @brief  Structure describing the filling for the chart
            struct AreaFill
            {
                EPlotAreaFillStyle Style;
                std::string SolidColor;         ///< SolidlColor RGB string like "FF00FF" for solid fill
                GradientFill Gradient;          ///< Params for a gradient fill
                EPatternFillStyle Pattern;
                std::string PatternFgColor, PatternBgColor;

                void SetLinearGradient( double Angle, bool ScaleAngle, const GradientStops & Stops )
                {
                    Style = PLOT_AREA_FILL_GRADIENT;
                    Gradient.FillType = GradientFill::linear;
                    Gradient.LinearAngle = Angle;
                    Gradient.LinearScaleAngle = ScaleAngle;
                    Gradient.ColorPoints = Stops;
                }
                void SetRadialAndRectGradient( GradientFill::fillType Type, EGradientFillDirection Dir, const GradientStops & Stops )
                {
                    Style = PLOT_AREA_FILL_GRADIENT;
                    Gradient.FillType = Type;
                    Gradient.FillDirection = Dir;
                    Gradient.ColorPoints = Stops;
                }
                void SetPathGradient( const GradientStops & Stops )
                {
                    Style = PLOT_AREA_FILL_GRADIENT;
                    Gradient.FillType = GradientFill::path;
                    Gradient.ColorPoints = Stops;
                }

                void SetPattern( EPatternFillStyle PatternStyle, const std::string & FgColor, const std::string & BgColor )
                {
                    Style = PLOT_AREA_FILL_PATTERN;
                    Pattern = PatternStyle;
                    PatternFgColor = FgColor;
                    PatternBgColor = BgColor;
                }

                AreaFill()
                {
                    Style = PLOT_AREA_FILL_NONE;
                    SolidColor = "FFFFFF";
                }
            };

            /// @brief  Structure describes diagramm properties
            struct Diagramm
            {
                UniString name;				///< name diagramm name (that will be depicted above the chart)
                uint32_t nameSize;			///< nameSize font size
                EPosition legend_pos;		///< legend_pos legend position
                ETableData tableData;		///< tableData table data state
                EChartTypes typeMain;		///< typeMain main chart type
                EChartTypes typeAdditional;	///< typeAdditional additional chart type (optional)

                EBarDirection barDir;
                EBarGrouping barGroup;
                EScatterStyle scatterStyle;

                EEmptyCellsDisplayMethod emptyCellsDisplayMethod;
                bool showDataFromHiddenCells;

                AreaFill plotAreaFill, chartAreaFill;

                Diagramm()
                {
                    nameSize = 18;
                    legend_pos = POS_RIGHT;
                    tableData = TBL_DATA_NONE;
                    typeMain = CHART_LINEAR;
                    typeAdditional = CHART_NONE;

                    barDir = BAR_DIR_VERTICAL;
                    barGroup = BAR_GROUP_CLUSTERED;
                    scatterStyle = SCATTER_POINT;
                    emptyCellsDisplayMethod = EMPTY_CELLS_DISP_GAPS;
                    showDataFromHiddenCells = false;
                }
            };

        private:
            size_t              m_index;            ///< chart ID number

            std::vector<Series> m_seriesSet;        ///< series set for main chart type
            std::vector<Series> m_seriesSetAdd;     ///< series set for additional chart type
            Diagramm            m_diagramm;			///< diagram object
            Axis                m_xAxis;            ///< main X axis object
            Axis                m_yAxis;            ///< main Y axis object
            Axis                m_x2Axis;           ///< additional X axis object
            Axis                m_y2Axis;           ///< additional Y axis object

            UniString         	m_title;            ///< chart sheet title

            PathManager    &    m_pathManager;      ///< reference to XML PathManager

        public:
            // *INDENT-OFF*   For AStyle tool

            // @section    Interclasses internal interface methods
            inline size_t   GetIndex() const                    { return m_index; }

            // @section    User interface
            inline const UniString & GetTitle() const           { return m_title;  }
            inline void SetTitle( const UniString & title )     { if( ! title.empty() ) m_title = title; }

            inline EChartTypes GetMainType() const              { return m_diagramm.typeMain; }
            inline void SetMainType( EChartTypes type )         { m_diagramm.typeMain = type; }

            inline EChartTypes GetAddType() const               { return m_diagramm.typeAdditional; }
            inline void SetAddType( EChartTypes type )          { m_diagramm.typeAdditional = type; }

            inline EEmptyCellsDisplayMethod GetEmptyCellsDisplayingMethod() const
                                                                { return m_diagramm.emptyCellsDisplayMethod; }
            inline void SetEmptyCellsDisplayingMethod( EEmptyCellsDisplayMethod emptyCellsDisplayMethod )
                                                                { m_diagramm.emptyCellsDisplayMethod = emptyCellsDisplayMethod; }

            inline bool GetShowDataFromHiddenCells() const      { return m_diagramm.showDataFromHiddenCells; }
            inline void SetShowDataFromHiddenCells( bool show ) { m_diagramm.showDataFromHiddenCells = show; }



            inline void SetPlotAreaFillNone()                   { m_diagramm.plotAreaFill.Style = PLOT_AREA_FILL_NONE; }
            inline void SetPlotAreaFillSolid( std::string fillColor )
                                                                {
                                                                  m_diagramm.plotAreaFill.Style = PLOT_AREA_FILL_SOLID;
                                                                  m_diagramm.plotAreaFill.SolidColor = fillColor;
                                                                }
            // Angle - 0 ... 359.9 degree (clockwise)
            // ScaleAngle - Specifies whether the angle scales with the fill region )
            inline void SetPlotAreaFillGradientLinear( double Angle, const GradientStops & Stops, bool ScaleAngle = true )
                                                                {
                                                                  assert( Stops.IsValid() );
                                                                  m_diagramm.plotAreaFill.SetLinearGradient( Angle, ScaleAngle, Stops );
                                                                }
            inline void SetPlotAreaFillGradientRadial( EGradientFillDirection Dir, const GradientStops & Stops )
                                                                {
                                                                  assert( Stops.IsValid() );
                                                                  m_diagramm.plotAreaFill.SetRadialAndRectGradient( GradientFill::radial, Dir, Stops );
                                                                }
            inline void SetPlotAreaFillGradientRectangular( EGradientFillDirection Dir, const GradientStops & Stops )
                                                                {
                                                                  assert( Stops.IsValid() );
                                                                  m_diagramm.plotAreaFill.SetRadialAndRectGradient( GradientFill::rectangular, Dir, Stops );
                                                                }
            inline void SetPlotAreaFillGradientPath( const GradientStops & Stops )
                                                                {
                                                                  assert( Stops.IsValid() );
                                                                  m_diagramm.plotAreaFill.SetPathGradient( Stops );
                                                                }
            inline void SetPlotAreaFillPattern( EPatternFillStyle PatternStyle, const std::string & FgColor, const std::string & BgColor )
                                                                {
                                                                  m_diagramm.plotAreaFill.SetPattern( PatternStyle, FgColor, BgColor );
                                                                }


            inline void SetChartAreaFillNone()                  { m_diagramm.chartAreaFill.Style = PLOT_AREA_FILL_NONE; }
            inline void SetChartAreaFillSolid( std::string fillColor )
                                                                {
                                                                  m_diagramm.chartAreaFill.Style = PLOT_AREA_FILL_SOLID;
                                                                  m_diagramm.chartAreaFill.SolidColor = fillColor;
                                                                }
            // Angle - 0 ... 359.9 degree (clockwise)
            // ScaleAngle - Specifies whether the angle scales with the fill region )
            inline void SetChartAreaFillGradientLinear( double Angle, const GradientStops & Stops, bool ScaleAngle = true )
                                                                {
                                                                  assert( Stops.IsValid() );
                                                                  m_diagramm.chartAreaFill.SetLinearGradient( Angle, ScaleAngle, Stops );
                                                                }
            inline void SetChartAreaFillGradientRadial( EGradientFillDirection Dir, const GradientStops & Stops )
                                                                {
                                                                  assert( Stops.IsValid() );
                                                                  m_diagramm.chartAreaFill.SetRadialAndRectGradient( GradientFill::radial, Dir, Stops );
                                                                }
            inline void SetChartAreaFillGradientRectangular( EGradientFillDirection Dir, const GradientStops & Stops )
                                                                {
                                                                  assert( Stops.IsValid() );
                                                                  m_diagramm.chartAreaFill.SetRadialAndRectGradient( GradientFill::rectangular, Dir, Stops );
                                                                }
            inline void SetChartAreaFillGradientPath( const GradientStops & Stops )
                                                                {
                                                                  assert( Stops.IsValid() );
                                                                  m_diagramm.chartAreaFill.SetPathGradient( Stops );
                                                                }
            inline void SetChartAreaFillPattern( EPatternFillStyle PatternStyle, const std::string & FgColor, const std::string & BgColor )
                                                                {
                                                                  m_diagramm.chartAreaFill.SetPattern( PatternStyle, FgColor, BgColor );
                                                                }


            inline void SetDiagrammNameSize( uint32_t size )    { m_diagramm.nameSize = size;   }
            inline void SetDiagrammName( const UniString & name ){ m_diagramm.name = name;       }
            inline void SetTableDataState( ETableData state )   { m_diagramm.tableData = state; }
            inline void SetLegendPos( EPosition pos )           { m_diagramm.legend_pos = pos;  }

            inline void SetBarDirection( EBarDirection barDir ) { m_diagramm.barDir = barDir;      }
            inline void SetBarGrouping( EBarGrouping barGroup ) { m_diagramm.barGroup = barGroup;  }
            inline void SetScatterStyle( EScatterStyle style )  { m_diagramm.scatterStyle = style; }

            inline void SetXAxisLblInterval( int value )            { m_xAxis.lblSkipInterval = value;  }
            inline void SetXAxisMarkInterval( int value )           { m_xAxis.markSkipInterval = value; }
            inline void SetXAxisLblAngle( int degrees )             { m_xAxis.lblAngle = degrees;       }
            inline void SetXAxisMin( const std::string & value )    { m_xAxis.minValue = value;         }
            inline void SetXAxisMax( const std::string & value )    { m_xAxis.maxValue = value;         }
            inline void SetXAxisName( const UniString & name )      { m_xAxis.name = name;              }
            inline void SetXAxisNameSize( uint32_t size )           { m_xAxis.nameSize = size;          }
            inline void SetXAxisPos( EPosition pos )                { m_xAxis.pos = pos;                }
            inline void SetXAxisGrid( EGridLines state )            { m_xAxis.gridLines = state;        }
            inline void SetXAxisCross( EAxisCross cross )           { m_xAxis.cross = cross;            }


            inline void SetYAxisMin( const std::string & value )    { m_yAxis.minValue = value;         }
            inline void SetYAxisMax( const std::string & value )    { m_yAxis.maxValue = value;         }
            inline void SetYAxisName( const UniString & name )      { m_yAxis.name = name;              }
            inline void SetYAxisNameSize( uint32_t size )           { m_yAxis.nameSize = size;          }
            inline void SetYAxisPos( EPosition pos )                { m_yAxis.pos = pos;                }
            inline void SetYAxisGrid( EGridLines state )            { m_yAxis.gridLines = state;        }
            inline void SetYAxisCross( EAxisCross cross )           { m_yAxis.cross = cross;            }


            inline void SetX2AxisLblInterval( int value )           { m_x2Axis.lblSkipInterval = value; }
            inline void SetX2AxisMarkInterval( int value )          { m_x2Axis.markSkipInterval = value;}
            inline void SetX2AxisLblAngle( int degrees )            { m_x2Axis.lblAngle = degrees;      }
            inline void SetX2AxisMin( const std::string & value )   { m_x2Axis.minValue = value;        }
            inline void SetX2AxisMax( const std::string & value )   { m_x2Axis.maxValue = value;        }
            inline void SetX2AxisName( const UniString & name )     { m_x2Axis.name = name;             }
            inline void SetX2AxisNameSize( uint32_t size )          { m_x2Axis.nameSize = size;         }
            inline void SetX2AxisPos( EPosition pos )               { m_x2Axis.pos = pos;               }
            inline void SetX2AxisGrid( EGridLines state )           { m_x2Axis.gridLines = state;       }
            inline void SetX2AxisCross( EAxisCross cross )          { m_x2Axis.cross = cross;           }

            inline void SetY2AxisMin( const std::string & value )   { m_y2Axis.minValue = value;        }
            inline void SetY2AxisMax( const std::string & value )   { m_y2Axis.maxValue = value;        }
            inline void SetY2AxisName( const UniString & name )     { m_y2Axis.name = name;             }
            inline void SetY2AxisNameSize( uint32_t size )          { m_y2Axis.nameSize = size;         }
            inline void SetY2AxisPos( EPosition pos )               { m_y2Axis.pos = pos;               }
            inline void SetY2AxisGrid( EGridLines state )           { m_y2Axis.gridLines = state;       }
            inline void SetY2AxisCross( EAxisCross cross )          { m_y2Axis.cross = cross;           }
///< Patch to val/cat issue E.N.
            inline void SetXAxisToCat(const bool set2val=false )       { m_xAxis.isVal = set2val;  }
            inline void SetYAxisToCat(const bool set2val=false )       { m_yAxis.isVal = set2val;  }
            inline void SetX2AxisToCat(const bool set2val=false )      { m_x2Axis.isVal = set2val;  }
            inline void SetY2AxisToCat(const bool set2val=false )      { m_y2Axis.isVal = set2val;  }

            // *INDENT-ON*    For AStyle tool

            bool AddSeries( const Series & series, bool mainChart = true );

        protected:
            CChart( size_t index, EChartTypes type, PathManager & pathmanager );
            virtual ~CChart();

        private:
            //Disable copy and assignment
            CChart( const CChart & that );
            CChart & operator=( const CChart & );

            bool Save();

            static void AddTitle( XMLWriter & xmlw, const UniString & name, uint32_t size, bool vertPos );
            static void AddTableData( XMLWriter & xmlw , ETableData tableData );
            static void AddLegend( XMLWriter & xmlw , EPosition legend_pos );
            static void AddXAxis( XMLWriter & xmlw, const Axis & x, uint32_t crossAxisId = 0 );
            static void AddYAxis( XMLWriter & xmlw, const Axis & y, uint32_t crossAxisId = 0 );

            static void AddLineChart( XMLWriter & xmlw, Axis & xAxis, uint32_t yAxisId, const std::vector<Series> & series, size_t firstSeriesId );
            static void AddBarChart( XMLWriter & xmlw, Axis & xAxis, uint32_t yAxisId, const std::vector<Series> & series, size_t firstSeriesId, EBarDirection barDir, EBarGrouping barGroup );
            static void AddScatterChart( XMLWriter & xmlw, uint32_t xAxisId, uint32_t yAxisId, const std::vector<Series> & series, size_t firstSeriesId, EScatterStyle style );
            static void AddMarker( XMLWriter & xmlw, const Series & ser , const char * markerID );
            static void AddAreaFill( XMLWriter & xmlw, const AreaFill & areaFill );

            static std::string CellRangeString( const std::string & Title, const CellCoord & CellFrom, const CellCoord & szCellTo );

            static inline char GetCharForPos( EPosition Pos, char DefaultChar )
            {
                switch( Pos )
                {
                    case POS_LEFT   :   return 'l';
                    case POS_RIGHT  :   return 'r';
                    case POS_TOP    :   return 't';
                    case POS_BOTTOM :   return 'b';
                    case POS_LEFT_ASIDE     :   return DefaultChar;
                    case POS_RIGHT_ASIDE    :   return DefaultChar;
                    case POS_TOP_ASIDE      :   return DefaultChar;
                    case POS_BOTTOM_ASIDE   :   return DefaultChar;
                    case POS_NONE           :   return DefaultChar;
                }
                return DefaultChar;
            }

            friend class CWorkbook;
    };

}

#endif // XLSX_CCHART_H
