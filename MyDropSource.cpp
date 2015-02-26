#include "MyDropSource.h"

//Constructors
MyDropSource::MyDropSource() : refcount(1)
{
	fprintf(stdout, "Initiate a drop source!\n");
}

//Destructors
MyDropSource::~MyDropSource()
{
	refcount = 0;
	fprintf(stdout, "Destruct a drop source!\n");
}

//IUnkown implementation
ULONG __stdcall MyDropSource::AddRef()
{
	return InterlockedIncrement(&refcount);
}

ULONG __stdcall MyDropSource::Release()
{
	ULONG nRefCount = InterlockedDecrement(&refcount);

	if(nRefCount == 0)
		delete this;

	return nRefCount;
}

STDMETHODIMP MyDropSource::QueryInterface(REFIID riid, void **ppvObject) {
	// check to see what interface has been requested
    if (!ppvObject)
        return E_POINTER;

    if (riid == IID_IDataObject)
        *ppvObject = (IDataObject*)this;
    else if (riid == IID_IUnknown)
        *ppvObject = (IUnknown*)this;
    else
    {
        *ppvObject = 0;
        return E_NOINTERFACE;
    }

    AddRef();
	return S_OK;
}

//DropSource implement
STDMETHODIMP __stdcall MyDropSource::QueryContinueDrag(BOOL fEscapePressed, DWORD grfKeyState)
{
	// if the Escape key has been pressed since the last call, cancel the drop
    if(fEscapePressed == TRUE)
        return DRAGDROP_S_CANCEL;

    // if the LeftMouse button has been released, then do the drop!
    if((grfKeyState & MK_LBUTTON) == 0)
        return DRAGDROP_S_DROP;

    // continue with the drag-drop
    return S_OK;
}

STDMETHODIMP MyDropSource::GiveFeedback(DWORD dwEffect)
{
	UNREFERENCED_PARAMETER(dwEffect);
    return DRAGDROP_S_USEDEFAULTCURSORS;
}
