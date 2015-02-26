//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop

#include "MainWindow.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Label1StartDrag(TObject *Sender, TDragObject *&DragObject)
{
	Label1->Caption = "Start drag";

	//Source file
	char tFileName[256] = "D:\\119.dat";

	//Prepare FOTMATETC
	fmtetc.cfFormat = CF_HDROP;
	fmtetc.dwAspect = DVASPECT_CONTENT;
	fmtetc.lindex = -1;
	fmtetc.ptd = (void*)0;
	fmtetc.tymed = TYMED_HGLOBAL;

	//Prepare DROPFILES
	DROPFILES* tDropFiles;
	//Fill the filename
	HGLOBAL hGblFiles;
	LPSTR lpData;
	stgmed.hGlobal = GlobalAlloc(GHND, sizeof(DROPFILES)+strlen(tFileName)+ 2);
	if(0 == stgmed.hGlobal)
		MessageBoxA(NULL, "OUT_OF_MEMORY!!!", "OUT_OF_MEMORY", 0);

	tDropFiles = (DROPFILES*)GlobalLock(stgmed.hGlobal);
	ZeroMemory(tDropFiles, sizeof(DROPFILES)+strlen(tFileName)+2);
	strcpy(((char*)tDropFiles)+sizeof(DROPFILES), tFileName);
	GlobalUnlock(stgmed.hGlobal);

	tDropFiles->fNC 	= true;
	tDropFiles->fWide 	= false;
	tDropFiles->pFiles 	= sizeof(DROPFILES);
	tDropFiles->pt.x 	= 0;
	tDropFiles->pt.y 	= 0;

	//set hGlobal
	stgmed.tymed = TYMED_HGLOBAL;
	stgmed.hGlobal = tDropFiles;
	stgmed.pUnkForRelease = 0;

//	//为CF_TEXT创建格式描述
//	FORMATETC textFormatEtc = {0};
//	textFormatEtc.cfFormat = CF_TEXT;             // Clipboard format=CF_TEXT
//	textFormatEtc.ptd = NULL;                     // Target device=Screen
//	textFormatEtc.dwAspect = DVASPECT_CONTENT;    // Level of detail=Full content
//	textFormatEtc.lindex = -1;                    // Index=Not applicable
//	textFormatEtc.tymed = TYMED_HGLOBAL;          // Storage medium=Memory
//
//	//为CF_TEXT创建存储格式
//    // Copy the text string to a global memory block.
//	char szText[] = "Hello, World";
//	HANDLE hData = GlobalAlloc(GMEM_MOVEABLE, strlen(szText)+1);
//	LPSTR pData = (LPSTR)GlobalLock(hData);
//	lstrcpy(pData, szText);
//	GlobalUnlock(hData);
//
//	STGMEDIUM textStgMed = {0};
//	textStgMed.hGlobal = hData;
//	textStgMed.tymed = TYMED_HGLOBAL;
//	textStgMed.pUnkForRelease = NULL;

	//Create Instance of IDropSource and IDataObject
	pDropSource  = new MyDropSource();
	pDropSource->AddRef();
	pDataObject  = new MyDataObject();
	pDataObject->AddRef();

	//SetData
	pDataObject->SetData(&fmtetc, &stgmed, true);
//	pDataObject->SetData(&textFormatEtc, &textStgMed, false);

	OleInitialize(0);
	//注册自定义格式
//	UINT uDropEffect = RegisterClipboardFormat("Preferred DropEffect");
//	HGLOBAL hGblEffect = GlobalAlloc(GMEM_ZEROINIT | GMEM_MOVEABLE | GMEM_DDESHARE,sizeof(DWORD));
//	LPDWORD lpdDropEffect = (LPDWORD)GlobalLock(hGblEffect);
//	*lpdDropEffect = DROPEFFECT_COPY;

	//Ole剪贴板设置
//	HRESULT tSetClipboardOK = OleSetClipboard(pDataObject);
//	if(S_OK != tSetClipboardOK)
//	{
//		if(CLIPBRD_E_CANT_OPEN == tSetClipboardOK)
//			MessageBoxA(NULL, "CLIPBRD_E_CANT_OPEN!!!", "CLIPBRD_E_CANT_OPEN", 0);
//		if(CLIPBRD_E_CANT_EMPTY == tSetClipboardOK)
//			MessageBoxA(NULL, "CLIPBRD_E_CANT_EMPTY!!!", "CLIPBRD_E_CANT_EMPTY", 0);
//		if(CLIPBRD_E_CANT_CLOSE == tSetClipboardOK)
//			MessageBoxA(NULL, "CLIPBRD_E_CANT_CLOSE!!!", "CLIPBRD_E_CANT_CLOSE", 0);
//		if(CLIPBRD_E_CANT_SET == tSetClipboardOK)
//			MessageBoxA(NULL, "CLIPBRD_E_CANT_SET!!!", "CLIPBRD_E_CANT_SET", 0);
////		MessageBoxA(NULL, "tSetClipboardOK::CLIPBRD_E_UNKNOWN!!!", "tSetClipboardOK::CLIPBRD_E_UNKNOWN", 0);
//		return;
//	}
//
//	//检测Ole剪贴板中是否存在目标
//	HRESULT tIsInOleClipboard = OleIsCurrentClipboard(pDataObject);
//	if(S_FALSE == tIsInOleClipboard)
//	{
//	   MessageBoxA(NULL, "NOT IN CLIPBOARD!!!", "tIsInOleClipboard", 0);
//	   return;
//	}

	//Invoke DoDragDrop
    DWORD dwEffect;
	HRESULT tResult = DoDragDrop((IDataObject*)pDataObject, (IDropSource*)pDropSource, DROPEFFECT_MOVE, &dwEffect);

	//Ckeck drag&drop result
	if(tResult != DRAGDROP_S_DROP)
		{
		if(tResult == DRAGDROP_S_CANCEL)
			MessageBoxA(NULL, "DRAGDROP_S_CANCEL!", "DRAGDROP_S_DROP", 0);
		else
			MessageBoxA(NULL, "E_UNSPEC!", "DRAGDROP_S_DROP", 0);
		return;
		}

	if((dwEffect & DROPEFFECT_MOVE) == DROPEFFECT_MOVE)
		MessageBoxA(NULL, "Ole drag&drop OK!!", "DRAGDROP_S_DROP", 0);
	else
		{
		if((dwEffect & DROPEFFECT_NONE) == DROPEFFECT_NONE)
			MessageBoxA(NULL, "DROPEFFECT_NONE!!", "DRAGDROP_S_DROP", 0);
		}

	//检测Ole剪贴板中是否存在目标
//	tIsInOleClipboard = OleIsCurrentClipboard(pDataObject);
//	if(S_FALSE == tIsInOleClipboard)
//	{
//	   MessageBoxA(NULL, "NOT IN CLIPBOARD!!!", "tIsInOleClipboard", 0);
//	   return;
//	}

	//Clean
	pDropSource->Release();
	pDataObject->Release();

	OleUninitialize();

	return;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Label1EndDrag(TObject *Sender, TObject *Target, int X, int Y)
{
	Label1->Caption = "Finish drag";
	return;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormCreate(TObject* Sender)
{
	DragAcceptFiles(this->Handle, true);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormDestroy(TObject* Sender)
{
	;
}
//---------------------------------------------------------------------------

