#include "types.h"

void __stdcall DecodeXAADPCM(
  uint8 *in,
  sint16 *out
) {

int i,j,r,f,x,y,z;
int K0[4]={ 0x0000, 0x00f0, 0x01cc, 0x0188};
int K1[4]={ 0x0000, 0x0000,-0x00d0,-0x00dc};
for(i=0;i<4;i++){

	r=12-(in[i*2+4]&15);
	f=in[i*2+4]>>4;
	for(j=0;j<28;j++){
		x=in[i+j*4+16]&15;
		if(x>=8) x-=16;
		x<<=r;
		x+=(y*K0[f]+z*K1[f])/256;
		if(x<-0x8000) x=-0x8000;
		if(x>=0x8000) x=0x7fff;
		out[i*56+j]=x;
		z=y; y=x;
	}

	r=12-(in[i*2+5]&15);
	f=in[i*2+5]>>4;
	for(j=0;j<28;j++){
		x=in[i+j*4+16]>>4;
		if(x>=8) x-=16;
		x<<=r;
		x+=(y*K0[f]+z*K1[f])/256;
		if(x<-0x8000) x=-0x8000;
		if(x>=0x8000) x=0x7fff;
		out[i*56+28+j]=x;
		z=y; y=x;
	}
}

}