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

#include "Drawing.h"

#include "Chart.h"
#include "XlsxHeaders.h"

#include "../PathManager.hpp"
#include "../XMLWriter.hpp"


namespace SimpleXlsx
{
    CDrawing::CDrawing( size_t index, PathManager & pathmanager ) : m_index( index ), m_pathManager( pathmanager )
    {
    }

    CDrawing::~CDrawing()
    {
    }

    bool CDrawing::Save()
    {
        if( ! IsEmpty() )
        {
            SaveDrawingRels();
            SaveDrawing();
        }
        return true;
    }

    // [- /xl/drawings/_rels/drawingX.xml.rels
    void CDrawing::SaveDrawingRels()
    {
        std::stringstream FileName;
        FileName << "/xl/drawings/_rels/drawing" << m_index << ".xml.rels";

        XMLWriter xmlw( m_pathManager.RegisterXML( FileName.str() ) );
        xmlw.Tag( "Relationships" ).Attr( "xmlns", ns_relationships );
        int rId = 1;
        for( std::vector<DrawingInfo>::const_iterator it = m_drawings.begin(); it != m_drawings.end(); it++, rId++ )
        {
            std::stringstream Target, rIdStream;
            const char * TypeString = NULL;
            switch( ( * it ).AType )
            {
                case DrawingInfo::absoluteAnchor:
                case DrawingInfo::twoCellAnchor:
                {
                    TypeString = type_chart;
                    Target << "../charts/chart" << ( *it ).Chart->GetIndex() << ".xml";
                    break;
                }
                case DrawingInfo::imageOneCellAnchor:
                case DrawingInfo::imageTwoCellAnchor:
                {
                    TypeString = type_image;
                    Target << "../media/" << ( *it ).Image->InternalName;
                    break;
                }
            }
            rIdStream << "rId" << rId;

            xmlw.TagL( "Relationship" ).Attr( "Id", rIdStream.str() ).Attr( "Type", TypeString ).Attr( "Target", Target.str() ).EndL();
        }
        xmlw.End( "Relationships" );
    }

    // [- /xl/drawings/drawingX.xml
    void CDrawing::SaveDrawing()
    {
        std::stringstream FileName;
        FileName << "/xl/drawings/drawing" << m_index << ".xml";

        XMLWriter xmlw( m_pathManager.RegisterXML( FileName.str() ) );
        xmlw.Tag( "xdr:wsDr" ).Attr( "xmlns:xdr", ns_xdr ).Attr( "xmlns:a", ns_a );

        int rId = 1;
        for( std::vector< DrawingInfo >::const_iterator it = m_drawings.begin(); it != m_drawings.end(); it++, rId++ )
        {
            switch( ( * it ).AType )
            {
                case DrawingInfo::absoluteAnchor:
                {
                    xmlw.Tag( "xdr:absoluteAnchor" );
                    xmlw.TagL( "xdr:pos" ).Attr( "x", 0 ).Attr( "y", 0 ).EndL();
                    xmlw.TagL( "xdr:ext" ).Attr( "cx", 9312088 ).Attr( "cy", 6084794 ).EndL();
                    SaveChartSection( xmlw, ( * it ).Chart, rId );
                    xmlw.End( "xdr:absoluteAnchor" );
                    break;
                }
                case DrawingInfo::twoCellAnchor:
                {
                    xmlw.Tag( "xdr:twoCellAnchor" );
                    SaveChartPoint( xmlw, "xdr:from", ( * it ).TopLeft );
                    SaveChartPoint( xmlw, "xdr:to", ( * it ).BottomRight );
                    SaveChartSection( xmlw, ( * it ).Chart, rId );
                    xmlw.End( "xdr:twoCellAnchor" );
                    break;
                }
                case DrawingInfo::imageOneCellAnchor:
                {
                    xmlw.Tag( "xdr:oneCellAnchor" ).Attr( "editAs", "oneCell" );
                    SaveChartPoint( xmlw, "xdr:from", ( * it ).TopLeft );
                    xmlw.TagL( "xdr:ext" ).Attr( "cx", ( * it ).BottomRight.col ).Attr( "cy", ( * it ).BottomRight.row ).EndL();
                    SaveImageSection( xmlw, ( * it ).Image, rId );
                    xmlw.End( "xdr:oneCellAnchor" );
                    break;
                }
                case DrawingInfo::imageTwoCellAnchor:
                {
                    xmlw.Tag( "xdr:twoCellAnchor" );
                    SaveChartPoint( xmlw, "xdr:from", ( * it ).TopLeft );
                    SaveChartPoint( xmlw, "xdr:to", ( * it ).BottomRight );
                    SaveImageSection( xmlw, ( * it ).Image, rId );
                    xmlw.End( "xdr:twoCellAnchor" );
                    break;
                }
            }
        }
        xmlw.End( "xdr:wsDr" );
    }

    void CDrawing::SaveChartSection( XMLWriter & xmlw, CChart * chart, int rId )
    {
        std::stringstream rIdStream;
        rIdStream << "rId" << rId;

        xmlw.Tag( "xdr:graphicFrame" ).Attr( "macro", "" ).Tag( "xdr:nvGraphicFramePr" );
        xmlw.TagL( "xdr:cNvPr" ).Attr( "id", rId ).Attr( "name", chart->GetTitle() ).EndL();
        xmlw.Tag( "xdr:cNvGraphicFramePr" ).TagL( "a:graphicFrameLocks" ).Attr( "noGrp", 1 ).EndL().End( "xdr:cNvGraphicFramePr" );
        xmlw.End( "xdr:nvGraphicFramePr" );

        xmlw.Tag( "xdr:xfrm" );
        xmlw.TagL( "a:off" ).Attr( "x", 0 ).Attr( "y", 0 ).EndL();
        xmlw.TagL( "a:ext" ).Attr( "cx", 0 ).Attr( "cy", 0 ).EndL();
        xmlw.End( "xdr:xfrm" );

        xmlw.Tag( "a:graphic" ).Tag( "a:graphicData" ).Attr( "uri", ns_c );
        xmlw.TagL( "c:chart" ).Attr( "xmlns:c", ns_c ).Attr( "xmlns:r", ns_book_r ).Attr( "r:id", rIdStream.str() ).EndL();
        xmlw.End( "a:graphicData" ).End( "a:graphic" );

        xmlw.End( "xdr:graphicFrame" );
        xmlw.TagL( "xdr:clientData" ).EndL();
    }

    void CDrawing::SaveImageSection( XMLWriter & xmlw, CImage * image, int rId )
    {
        std::stringstream rIdStream;
        rIdStream << "rId" << rId;

        xmlw.Tag( "xdr:pic" ).Tag( "xdr:nvPicPr" );
        xmlw.TagL( "xdr:cNvPr" ).Attr( "id", rId ).Attr( "name", image->InternalName ).EndL();
        xmlw.Tag( "xdr:cNvPicPr" ).TagL( "a:picLocks" ).Attr( "noChangeAspect", 1 ).EndL().End( "xdr:cNvPicPr" );
        xmlw.End( "xdr:nvPicPr" );

        xmlw.Tag( "xdr:blipFill" );
        xmlw.Tag( "a:blip" ).Attr( "xmlns:r", ns_relationships_chart ).Attr( "r:embed", rIdStream.str() );
        xmlw.End( "a:blip" );
        xmlw.Tag( "a:stretch" ).TagL( "a:fillRect" ).EndL().End( "a:stretch" );
        xmlw.End( "xdr:blipFill" );

        xmlw.Tag( "xdr:spPr" );
        xmlw.Tag( "a:prstGeom" ).Attr( "prst", "rect" ).TagL( "a:avLst" ).EndL().End( "a:prstGeom" );
        xmlw.End( "xdr:spPr" );

        xmlw.End( "xdr:pic" );

        xmlw.TagL( "xdr:clientData" ).EndL();
    }

    void CDrawing::SaveChartPoint( XMLWriter & xmlw, const char * Tag, const DrawingPoint & Point )
    {
        xmlw.Tag( Tag );
        xmlw.TagOnlyContent( "xdr:col", Point.col );
        xmlw.TagOnlyContent( "xdr:colOff", Point.colOff );
        xmlw.TagOnlyContent( "xdr:row", Point.row );
        xmlw.TagOnlyContent( "xdr:rowOff", Point.rowOff );
        xmlw.End( Tag );
    }
}
