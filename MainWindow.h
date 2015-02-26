//---------------------------------------------------------------------------

#ifndef MainWindowH
#define MainWindowH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Ole2.h>
#include "MyDataObject.h"
//#include "CDataObject.h"
//#include "CDropSource.h"
#include "MyDropSource.h"
//#include "Lib.h"
//#include "SDKDataObject.h"
//#include "SDKDropSource.h"
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
	TLabel *Label1;
	void __fastcall Label1StartDrag(TObject *Sender, TDragObject *&DragObject);
	void __fastcall Label1EndDrag(TObject *Sender, TObject *Target, int X, int Y);
	void __fastcall FormCreate(TObject *Sender);
	void __fastcall FormDestroy(TObject *Sender);
private:	// User declarations
	//准备两个接口实例
	IDataObject *pDataObject;
	IDropSource *pDropSource;

	//
	STGMEDIUM stgmed;
	FORMATETC fmtetc;

public:		// User declarations
	__fastcall TForm1(TComponent* Owner);

	int TCopyFileToClipboard(char szFileName[]);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
