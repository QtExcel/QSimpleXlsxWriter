#ifndef XLSXCOLORLIB_H_INCLUDED
#define XLSXCOLORLIB_H_INCLUDED
#include <map>
#include "clsRGBColorRecord.h"
namespace SimpleXlsx
{
 class
   XLSXColorLib{
   std::map<std::string, clsRGBColorRecord> lib;
   public:
     XLSXColorLib(){};
     virtual ~XLSXColorLib();
     void AddColor(const char * id, const clsRGBColorRecord & cl){ lib[id]=cl;  };
     const char * GetColor(const char * id) { return lib.at(id).Get(); };
     void Clear() { lib.clear(); };
   };
 extern void make_grayscale10(XLSXColorLib & xlib);
 extern void make_excell_like_named_colors(XLSXColorLib & xlib);
}
#endif // XLSXCOLORLIB_H_INCLUDED
