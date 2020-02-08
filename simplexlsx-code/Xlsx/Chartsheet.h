/*
  SimpleXlsxWriter
  Copyright (C) 2012-2020 Pavel Akimov <oxod.pavel@gmail.com>, Alexandr Belyak <programmeralex@bk.ru>

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

#ifndef XLSX_CHARTSHEET_H
#define XLSX_CHARTSHEET_H

#include <stdlib.h>

#include "SimpleXlsxDef.h"

// ****************************************************************************
/// @brief  The following namespace contains declarations of a number of classes
///         which allow writing .xlsx files with formulae and charts
/// @note   All classes inside the namepace are used together, so a separate
///         using will not guarantee stability and reliability
// ****************************************************************************
namespace SimpleXlsx
{
    class CChart;
    class CDrawing;
    class PathManager;

    // ****************************************************************************
    /// @brief	The class CChartsheet is used for creation sheet with chart.
    /// @see    EChartTypes supplying types of charts
    /// @note   All created chartsheets inside workbook will be allocated on its` own sheet
    // ****************************************************************************
    class CChartsheet : public CSheet
    {
        public:
            //Reference to the Chart
            inline CChart & Chart()
            {
                return m_Chart;
            }

            virtual const UniString & GetTitle() const;

        protected:
            CChartsheet( size_t index, CChart & chart, CDrawing & drawing, PathManager & pathmanager );
            virtual ~CChartsheet();

        private:
            //Disable copy and assignment
            CChartsheet( const CChartsheet & that );
            CChartsheet & operator=( const CChartsheet & );

            CChart       &      m_Chart;            ///< Reference to chart
            CDrawing      &     m_Drawing;          ///< Reference to drawing object
            PathManager    &    m_pathManager;      ///< reference to XML PathManager

            bool Save();

            friend class CWorkbook;
    };

}	// namespace SimpleXlsx

#endif	// XLSX_CHARTSHEET_H
