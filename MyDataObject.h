#ifndef _MYDATAOBJECT_H_
#define _MYDATAOBJECT_H_

#include <stdio.h>
#include "IDragDemo.h"
#include "MyDropSource.h"

#ifndef SAFE_DELETE
#define SAFE_DELETE(p) { if(p){delete(p);  (p)=NULL;} }
#endif

HRESULT CreateEnumFormatEtc(UINT cfmt, FORMATETC *afmt, IEnumFORMATETC **ppEnumFormatEtc);

class MyDataObject : public IDataObject
{
public:
    //IUnknown implementation
	ULONG __stdcall AddRef();
	ULONG __stdcall Release();
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);

    //IDataObject members
	STDMETHODIMP GetData               (FORMATETC *pformatetcIn, STGMEDIUM *pmedium);
	STDMETHODIMP GetDataHere           (FORMATETC *pformatetc,   STGMEDIUM *pmedium);
    STDMETHODIMP QueryGetData          (FORMATETC *pformatetc);
	STDMETHODIMP GetCanonicalFormatEtc (FORMATETC *pformatectIn, FORMATETC *pformatetcOut);
	STDMETHODIMP SetData               (FORMATETC *pformatetc,   STGMEDIUM *pmedium,  BOOL fRelease);
	STDMETHODIMP EnumFormatEtc         (DWORD     dwDirection,   IEnumFORMATETC **ppenumFormatEtc);
	STDMETHODIMP DAdvise               (FORMATETC *pformatetc,   DWORD advf, IAdviseSink *pAdvSink, DWORD *pdwConnection);
	STDMETHODIMP DUnadvise             (DWORD     dwConnection);
    STDMETHODIMP EnumDAdvise           (IEnumSTATDATA **ppenumAdvise);
public:
	MyDataObject();
	~MyDataObject();
private:
	LONG refcount;

	//自定义的接受格式
	FORMATETC* m_AcceptFormat;
	STGMEDIUM* m_StorageMedium;

	HGLOBAL DupGlobalMem(HGLOBAL hMem);

	//Helper function
	HRESULT CopyMedium(STGMEDIUM* pMedDest, STGMEDIUM* pMedSrc, FORMATETC* pFmtSrc);
	HRESULT SetBlob(CLIPFORMAT cf, const void *pvBlob, UINT cbBlob);

	//需要一个DropSource去查询
//	MyDropSource* m_DropSource;

	//引用次数
	LONG m_RefCount;
};
//----------------MyEnumFormatEtc-----------------------------------------------------------
class MyEnumFormatEtc : public IEnumFORMATETC
{
public:
    // IUnknown members
    HRESULT __stdcall  QueryInterface (REFIID iid, void ** ppv)
    {
        if((iid==IID_IUnknown)||(iid==IID_IEnumFORMATETC))
        {
            *ppv=this;
            AddRef();
            return S_OK;
        }
        else
        {
            *ppv=NULL;
            return E_NOINTERFACE;
        }
    }
    ULONG   __stdcall AddRef (void) { return ++_iRefCount; }
    ULONG   __stdcall Release (void)  { if(--_iRefCount==0){delete this;  return 0;} return _iRefCount; }

    // IEnumFormatEtc members
    HRESULT __stdcall  Next  (ULONG celt, FORMATETC * rgelt, ULONG * pceltFetched);
    HRESULT __stdcall  Skip  (ULONG celt)
    {
        _nIndex += celt;
        return (_nIndex <= _nNumFormats) ? S_OK : S_FALSE;
    }
    HRESULT __stdcall  Reset (void)
    {
        _nIndex = 0;
        return S_OK;
    }
    HRESULT __stdcall  Clone (IEnumFORMATETC ** ppEnumFormatEtc)
    {
        HRESULT hResult;
        hResult = CreateEnumFormatEtc(_nNumFormats, _pFormatEtc, ppEnumFormatEtc);
        if(hResult == S_OK)
        {
            ((MyEnumFormatEtc *)*ppEnumFormatEtc)->_nIndex = _nIndex;
        }
        return hResult;
    }

    // Construction / Destruction
	MyEnumFormatEtc(FORMATETC *pFormatEtc, ULONG nNumFormats);
    ~MyEnumFormatEtc();

private:
    LONG        _iRefCount;        // Reference count for this COM interface
    ULONG       _nIndex;           // current enumerator index
    ULONG       _nNumFormats;      // number of FORMATETC members
    FORMATETC * _pFormatEtc;       // array of FORMATETC objects
};
//---------------------------------------------------------------------------

#endif
