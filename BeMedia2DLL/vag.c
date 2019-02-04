#include "types.h"

void __stdcall DecodeVAG(
  uint8  *src,
  sint16 *outbuf,
  sint32 *p1,
  sint32 *p2
) {
  static sint32 coef[5][2] = {
    {  0,   0},
    { 60,   0},
    {115, -52},
    { 98, -55},
    {122, -60}
  };
  sint32 c;
  uint32 i;
  uint32 p = 0;

    uint8  fm = *src++;
    uint32 filter    = (fm >> 4) & 0xF;
    uint32 magnitude = (fm     ) & 0xF;
    if(magnitude > 12 || filter > 4) {
      magnitude = 12;
      filter = 0;
    }
    src++;
    for(i = 0; i < 14; i++) {
      uint32 d = *src++;
      sint32 d1 = (d & 0x0F) << (12 + 16);
      sint32 d2 = (d & 0xF0) << ( 8 + 16);
      d1 >>= magnitude + 16;
      d2 >>= magnitude + 16;

      c = d1 +
        ((((*p1) * coef[filter][0]) +
          ((*p2) * coef[filter][1])) >> 6);
      if(c < -32768) c = -32768;
      if(c >  32767) c =  32767;

      *outbuf++ = c;
      (*p2) = (*p1);
      (*p1) = c;

      c = d2 +
        ((((*p1) * coef[filter][0]) +
          ((*p2) * coef[filter][1])) >> 6);

      if(c < -32768) c = -32768;
      if(c >  32767) c =  32767;

      *outbuf++ = c;
      (*p2) = (*p1);
      (*p1) = c;
    }

}
