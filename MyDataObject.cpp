 #include "MyDataObject.h"
 #include "MyDropSource.h"
 #include <Urlmon.h>

//Constructors
MyDataObject::MyDataObject()
{
	m_RefCount = 0;
}

//Destructors
MyDataObject::~MyDataObject()
{
	refcount = 0;

	SAFE_DELETE(m_StorageMedium);
	SAFE_DELETE(m_AcceptFormat);
}

//IUnkown implementation
ULONG __stdcall MyDataObject::AddRef()
{
	return InterlockedIncrement(&m_RefCount);
}

ULONG __stdcall MyDataObject::Release()
{
	ULONG nRefCount = InterlockedDecrement(&m_RefCount);

	if (nRefCount == 0)
		delete this;

	return nRefCount;
}

STDMETHODIMP MyDataObject::QueryInterface(REFIID riid, void **ppvObject) {

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

//Data Object implement
//呈现指定的 pformatetcIn 结构中描述的数据，并将它传输给 pmedium
STDMETHODIMP MyDataObject::GetData(FORMATETC *pformatetcIn, STGMEDIUM *pmedium)
{
	//入参检查
	if ( (NULL == pformatetcIn) || (NULL == pmedium) )
    {
        return E_INVALIDARG;
    }

    pmedium->hGlobal = NULL;

	//要求的格式符合要求即可传递给数据消费者
	if( (pformatetcIn->tymed & m_AcceptFormat->tymed) &&
		(pformatetcIn->dwAspect == m_AcceptFormat->dwAspect) &&
		(pformatetcIn->cfFormat == m_AcceptFormat->cfFormat) )
		{
            return CopyMedium(pmedium, m_StorageMedium, m_AcceptFormat);
        }

    return DV_E_FORMATETC;
}

STDMETHODIMP MyDataObject::GetDataHere(FORMATETC *pformatetc, STGMEDIUM *pmedium)
{
	return E_NOTIMPL;
}

STDMETHODIMP MyDataObject::QueryGetData(FORMATETC *pformatetc)
{
	if(NULL == pformatetc )
    {
        return E_INVALIDARG;
    }
    if(!(DVASPECT_CONTENT & pformatetc->dwAspect))
    {
        return DV_E_DVASPECT;
    }
    HRESULT hr = DV_E_TYMED;

	if(m_AcceptFormat->tymed & pformatetc->tymed )
        {
		if(m_AcceptFormat->cfFormat == pformatetc->cfFormat )
            {
			return S_OK;
            }
		else
            {
			hr = DV_E_CLIPFORMAT;
            }
        }
	else
        {
            hr = DV_E_TYMED;
		}
    return hr;
}

STDMETHODIMP MyDataObject::GetCanonicalFormatEtc(FORMATETC *pformatectIn, FORMATETC *pformatetcOut)
{
	pformatetcOut->ptd = NULL;
	return E_NOTIMPL;
}

STDMETHODIMP MyDataObject::SetData(FORMATETC* pformatetc, STGMEDIUM* pmedium, BOOL fRelease)
{
	//入参检查
	if ( (NULL == pformatetc) || (NULL == pmedium) )
        return E_INVALIDARG;


	if ( pformatetc->tymed != pmedium->tymed )
		return E_FAIL;

	//将传入的格式描述结构传给对象
	m_AcceptFormat = new FORMATETC;
	m_StorageMedium = new STGMEDIUM;
	ZeroMemory(m_AcceptFormat, sizeof(FORMATETC));
	ZeroMemory(m_StorageMedium, sizeof(STGMEDIUM));

    if(TRUE == fRelease)
	{
		*m_StorageMedium = *pmedium;
    }
	else
    {
		CopyMedium(m_StorageMedium, pmedium, pformatetc);
	}

    return S_OK;
}

STDMETHODIMP MyDataObject::EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC **ppenumFormatEtc)
{
	if(NULL == ppenumFormatEtc)
    {
        return E_INVALIDARG;
    }
    *ppenumFormatEtc = NULL;
    HRESULT hr = E_NOTIMPL;
    if (DATADIR_GET == dwDirection )
    {
        FORMATETC rgfmtetc[] =
        {
			{ CF_HDROP, NULL, DVASPECT_CONTENT, 0, TYMED_HGLOBAL },
//			{ CF_TEXT, NULL, DVASPECT_CONTENT, 0, TYMED_HGLOBAL },
        };
        hr = CreateEnumFormatEtc(ARRAYSIZE(rgfmtetc), rgfmtetc, ppenumFormatEtc);
    }
    return hr;
}
//Advises:OLE_E_ADVISENOTSUPPORTED
STDMETHODIMP MyDataObject::DAdvise(FORMATETC *pformatetc, DWORD advf, IAdviseSink *pAdvSink, DWORD *pdwConnection)
{
	UNREFERENCED_PARAMETER(pformatetc);
    UNREFERENCED_PARAMETER(advf);
    UNREFERENCED_PARAMETER(pAdvSink);
    UNREFERENCED_PARAMETER(pdwConnection);
    return E_NOTIMPL;
}

STDMETHODIMP MyDataObject::DUnadvise(DWORD dwConnection)
{
	UNREFERENCED_PARAMETER(dwConnection);
	return E_NOTIMPL;
}

STDMETHODIMP MyDataObject::EnumDAdvise(IEnumSTATDATA **ppenumAdvise)
{
	UNREFERENCED_PARAMETER(ppenumAdvise);
	return E_NOTIMPL;
}
//Advises:OLE_E_ADVISENOTSUPPORTED

HGLOBAL MyDataObject::DupGlobalMem(HGLOBAL hMem)
{
	DWORD   len    = GlobalSize(hMem);
    PVOID   source = GlobalLock(hMem);
    PVOID   dest   = GlobalAlloc(GMEM_ZEROINIT | GMEM_MOVEABLE | GMEM_DDESHARE, len);

    memcpy(dest, source, len);
    GlobalUnlock(hMem);
    return dest;
}

HRESULT MyDataObject::CopyMedium(STGMEDIUM* pMedDest, STGMEDIUM* pMedSrc, FORMATETC* pFmtSrc)
{
    if ( (NULL == pMedDest) || (NULL ==pMedSrc) || (NULL == pFmtSrc) )
    {
        return E_INVALIDARG;
    }
    switch(pMedSrc->tymed)
    {
    case TYMED_HGLOBAL:
        pMedDest->hGlobal = (HGLOBAL)OleDuplicateData(pMedSrc->hGlobal, pFmtSrc->cfFormat, NULL);
        break;
    case TYMED_GDI:
        pMedDest->hBitmap = (HBITMAP)OleDuplicateData(pMedSrc->hBitmap, pFmtSrc->cfFormat, NULL);
        break;
    case TYMED_MFPICT:
        pMedDest->hMetaFilePict = (HMETAFILEPICT)OleDuplicateData(pMedSrc->hMetaFilePict, pFmtSrc->cfFormat, NULL);
        break;
    case TYMED_ENHMF:
        pMedDest->hEnhMetaFile = (HENHMETAFILE)OleDuplicateData(pMedSrc->hEnhMetaFile, pFmtSrc->cfFormat, NULL);
        break;
    case TYMED_FILE:
        pMedSrc->lpszFileName = (LPOLESTR)OleDuplicateData(pMedSrc->lpszFileName, pFmtSrc->cfFormat, NULL);
        break;
    case TYMED_ISTREAM:
        pMedDest->pstm = pMedSrc->pstm;
        pMedSrc->pstm->AddRef();
        break;
    case TYMED_ISTORAGE:
        pMedDest->pstg = pMedSrc->pstg;
        pMedSrc->pstg->AddRef();
        break;
    case TYMED_NULL:
    default:
        break;
    }
    pMedDest->tymed = pMedSrc->tymed;
    pMedDest->pUnkForRelease = NULL;
    if(pMedSrc->pUnkForRelease != NULL)
    {
        pMedDest->pUnkForRelease = pMedSrc->pUnkForRelease;
        pMedSrc->pUnkForRelease->AddRef();
    }
    return S_OK;
}
HRESULT MyDataObject::SetBlob(CLIPFORMAT cf, const void *pvBlob, UINT cbBlob)
{
    void *pv = GlobalAlloc(GPTR, cbBlob);
    HRESULT hr = pv ? S_OK : E_OUTOFMEMORY;
    if ( SUCCEEDED(hr) )
    {
        CopyMemory(pv, pvBlob, cbBlob);
        FORMATETC fmte = {cf, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
        // The STGMEDIUM structure is used to define how to handle a global memory transfer.
        // This structure includes a flag, tymed, which indicates the medium
        // to be used, and a union comprising pointers and a handle for getting whichever
        // medium is specified in tymed.
        STGMEDIUM medium = {0};
        medium.tymed = TYMED_HGLOBAL;
        medium.hGlobal = pv;
        hr = this->SetData(&fmte, &medium, TRUE);
        if (FAILED(hr))
        {
            GlobalFree(pv);
        }
    }
    return hr;
}

HRESULT CreateEnumFormatEtc(UINT cfmt, FORMATETC *afmt, IEnumFORMATETC **ppEnumFormatEtc)
{
	if (cfmt == 0 || afmt == 0 || ppEnumFormatEtc == 0)
		return E_INVALIDARG;

	*ppEnumFormatEtc = new MyEnumFormatEtc(afmt, cfmt);
    return (*ppEnumFormatEtc) ? S_OK: E_OUTOFMEMORY;
}

void DeepCopyFormatEtc(FORMATETC *dest, FORMATETC *source)
{
    // copy the source FORMATETC into dest
    *dest = *source;
    if(source->ptd)
    {
        // allocate memory for the DVTARGETDEVICE if necessary
        dest->ptd = (DVTARGETDEVICE*)CoTaskMemAlloc(sizeof(DVTARGETDEVICE));
        // copy the contents of the source DVTARGETDEVICE into dest->ptd
        *(dest->ptd) = *(source->ptd);
    }
}

MyEnumFormatEtc::MyEnumFormatEtc(FORMATETC *pFormatEtc, ULONG nNumFormats)
    :_iRefCount(1),_nIndex(0),_nNumFormats(nNumFormats)
{
    _pFormatEtc  = new FORMATETC[nNumFormats];
    // make a new copy of each FORMATETC structure
    for(ULONG i = 0; i < nNumFormats; i++)
    {
        DeepCopyFormatEtc (&_pFormatEtc[i], &pFormatEtc[i]);
    }
}
MyEnumFormatEtc::~MyEnumFormatEtc()
{
    // first free any DVTARGETDEVICE structures
    for(ULONG i = 0; i < _nNumFormats; i++)
    {
        if(_pFormatEtc[i].ptd)
            CoTaskMemFree(_pFormatEtc[i].ptd);
    }
    // now free the main array
    delete[] _pFormatEtc;
}

HRESULT __stdcall MyEnumFormatEtc::Next(ULONG celt, FORMATETC *pFormatEtc, ULONG *pceltFetched)
{
    ULONG copied = 0;
    // copy the FORMATETC structures into the caller's buffer
    while (_nIndex < _nNumFormats && copied < celt)
    {
        DeepCopyFormatEtc (&pFormatEtc [copied], &_pFormatEtc [_nIndex]);
        copied++;
        _nIndex++;
    }
    // store result
    if(pceltFetched != 0)
        *pceltFetched = copied;
    // did we copy all that was requested?
    return (copied == celt) ? S_OK : S_FALSE;
}

