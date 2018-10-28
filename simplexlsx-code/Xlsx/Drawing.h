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

#ifndef XLSX_DRAWING_H
#define XLSX_DRAWING_H

#include "SimpleXlsxDef.h"

namespace SimpleXlsx
{
    class CChart;

    class PathManager;
    class XMLWriter;

    class CDrawing
    {
        public:
            //Index of Drawing
            inline size_t GetIndex() const
            {
                return m_index;
            }

            //Is no registered Charts
            inline bool IsEmpty() const
            {
                return m_drawings.empty();
            }

        protected:
            CDrawing( size_t index, PathManager & pathmanager );
            virtual ~CDrawing();

        private:
            //Disable copy and assignment
            CDrawing( const CDrawing & that );
            CDrawing & operator=( const CDrawing & );

            size_t              m_index;            ///< Drawing ID number
            PathManager    &    m_pathManager;      ///< reference to XML PathManager

            struct DrawingInfo
            {
                enum AnchorType
                {
                    absoluteAnchor,         //For CChartSheet
                    twoCellAnchor,          //For chart on CWorkSheet
                    imageOneCellAnchor,     //For image on CWorkSheet
                    imageTwoCellAnchor,     //For image on CWorkSheet
                };

                CChart   *  Chart;
                CImage   *  Image;
                AnchorType  AType;
                DrawingPoint  TopLeft;        //For drawing on CWorkSheet
                DrawingPoint  BottomRight;    //For drawing on CWorkSheet
            };

            std::vector<DrawingInfo> m_drawings;


            //Append Chart from Chartsheet
            inline void AppendChart( CChart * Chart )
            {
                DrawingInfo CInfo = { Chart, NULL, DrawingInfo::absoluteAnchor, DrawingPoint(), DrawingPoint() };
                m_drawings.push_back( CInfo );
            }

            //Append Chart from Worksheet
            inline void AppendChart( CChart * Chart, const DrawingPoint & TopLeft, const DrawingPoint & BottomRight )
            {
                DrawingInfo CInfo = { Chart, NULL, DrawingInfo::twoCellAnchor, TopLeft, BottomRight };
                m_drawings.push_back( CInfo );
            }

            //Append Image from Worksheet
            inline void AppendImage( CImage * Image, const DrawingPoint & TopLeft, uint16_t XPercent, uint16_t YPercent )
            {
                DrawingPoint Dimen( uint32_t( Image->Width * XPercent ) * CImage::PointByPixel / 100,
                                    uint32_t( Image->Height * YPercent ) * CImage::PointByPixel / 100 );
                DrawingInfo CInfo = { NULL, Image, DrawingInfo::imageOneCellAnchor, TopLeft, Dimen };
                m_drawings.push_back( CInfo );
            }
            inline void AppendImage( CImage * Image, const DrawingPoint & TopLeft, const DrawingPoint & BottomRight )
            {
                DrawingInfo CInfo = { NULL, Image, DrawingInfo::imageTwoCellAnchor, TopLeft, BottomRight };
                m_drawings.push_back( CInfo );
            }

            bool Save();

            void SaveDrawingRels();
            void SaveDrawing();
            void SaveChartSection( XMLWriter & xmlw, CChart * chart, int rId );
            void SaveImageSection( XMLWriter & xmlw, CImage * image, int rId );
            void SaveChartPoint( XMLWriter & xmlw, const char * Tag, const DrawingPoint & Point );

            friend class CWorkbook;
    };

}

#endif // XLSX_DRAWING_H
