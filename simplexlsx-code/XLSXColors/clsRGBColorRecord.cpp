#include <stdio.h>
#include <stdlib.h>
#include <inttypes.h>
#include <algorithm>
#include "clsRGBColorRecord.h"
#include <string.h>
namespace SimpleXlsx
{


clsRGBColorRecord::clsRGBColorRecord()
{
    //ctor
    snprintf(colstr,sizeof(colstr),"%02X%02X%02X",0,0,0);

}

clsRGBColorRecord::~clsRGBColorRecord()
{
    //dtor
}

clsRGBColorRecord::clsRGBColorRecord(const clsRGBColorRecord& other)
{
    //copy ctor
    Copy(other);
}

clsRGBColorRecord& clsRGBColorRecord::operator=(const clsRGBColorRecord& rhs)
{
    if (this == &rhs) return *this; // handle self assignment
    Copy(rhs);
    return *this;
}
void clsRGBColorRecord::Copy(const clsRGBColorRecord& p){
   //std::move(p.colstr,p.colstr+sizeof(p.colstr),colstr);
   memcpy(colstr,p.colstr,7*sizeof(char));
 };

 void clsRGBColorRecord::Set(const char (&sharpn)[7]){
   //std::move(sharpn,sharpn+sizeof(sharpn),colstr);
   memcpy(colstr,sharpn,7*sizeof(char));
   };

 void clsRGBColorRecord::Set(const unsigned char r, const unsigned char g, const unsigned char b){
  snprintf(colstr,sizeof(colstr),"%02X%02X%02X",r,g,b);
  };
  void clsRGBColorRecord::Set(const unsigned char gs){
   snprintf(colstr,sizeof(colstr),"%02X%02X%02X",gs,gs,gs);
   };
  void clsRGBColorRecord::Set(const double dgs){
   unsigned char gs=(255*dgs)*0.01;
   Set(gs);
  };
};
