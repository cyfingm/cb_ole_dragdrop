#ifndef _MYDROPSOURCE_H_
#define _MYDROPSOURCE_H_

#include <stdio.h>
#include "IDragDemo.h"

class MyDropSource : public IDropSource
{
public:

	//IUnknown implementation
	ULONG __stdcall AddRef();
	ULONG __stdcall Release();
	STDMETHODIMP 	QueryInterface(REFIID riid, void **ppvObject);

    //IDropSource members
	STDMETHODIMP QueryContinueDrag(BOOL fEscapePressed, DWORD grfKeyState);
    STDMETHODIMP GiveFeedback(DWORD dwEffect);

	//Cons/Destructors
	MyDropSource();
	~MyDropSource();
private:
    LONG refcount;
};
#endif
