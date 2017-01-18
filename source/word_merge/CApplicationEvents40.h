// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CApplicationEvents40 包装类

class CApplicationEvents40 : public COleDispatchDriver
{
public:
	CApplicationEvents40(){} // 调用 COleDispatchDriver 默认构造函数
	CApplicationEvents40(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplicationEvents40(const CApplicationEvents40& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IApplicationEvents4 方法
public:
	STDMETHOD(Startup)()
	{
		HRESULT result;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(Quit)()
	{
		HRESULT result;
		InvokeHelper(0x2, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(DocumentChange)()
	{
		HRESULT result;
		InvokeHelper(0x3, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(DocumentOpen)(Document * Doc)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x4, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc);
		return result;
	}
	STDMETHOD(DocumentBeforeClose)(Document * Doc, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x6, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc, Cancel);
		return result;
	}
	STDMETHOD(DocumentBeforePrint)(Document * Doc, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x7, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc, Cancel);
		return result;
	}
	STDMETHOD(DocumentBeforeSave)(Document * Doc, BOOL * SaveAsUI, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL VTS_PBOOL ;
		InvokeHelper(0x8, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc, SaveAsUI, Cancel);
		return result;
	}
	STDMETHOD(NewDocument)(Document * Doc)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc);
		return result;
	}
	STDMETHOD(WindowActivate)(Document * Doc, Window * Wn)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0xa, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc, Wn);
		return result;
	}
	STDMETHOD(WindowDeactivate)(Document * Doc, Window * Wn)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0xb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc, Wn);
		return result;
	}
	STDMETHOD(WindowSelectionChange)(Selection * Sel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xc, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Sel);
		return result;
	}
	STDMETHOD(WindowBeforeRightClick)(Selection * Sel, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0xd, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Sel, Cancel);
		return result;
	}
	STDMETHOD(WindowBeforeDoubleClick)(Selection * Sel, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0xe, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Sel, Cancel);
		return result;
	}
	STDMETHOD(EPostagePropertyDialog)(Document * Doc)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xf, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc);
		return result;
	}
	STDMETHOD(EPostageInsert)(Document * Doc)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x10, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc);
		return result;
	}
	STDMETHOD(MailMergeAfterMerge)(Document * Doc, Document * DocResult)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x11, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc, DocResult);
		return result;
	}
	STDMETHOD(MailMergeAfterRecordMerge)(Document * Doc)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x12, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc);
		return result;
	}
	STDMETHOD(MailMergeBeforeMerge)(Document * Doc, long StartRecord, long EndRecord, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_I4 VTS_I4 VTS_PBOOL ;
		InvokeHelper(0x13, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc, StartRecord, EndRecord, Cancel);
		return result;
	}
	STDMETHOD(MailMergeBeforeRecordMerge)(Document * Doc, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x14, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc, Cancel);
		return result;
	}
	STDMETHOD(MailMergeDataSourceLoad)(Document * Doc)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x15, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc);
		return result;
	}
	STDMETHOD(MailMergeDataSourceValidate)(Document * Doc, BOOL * Handled)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x16, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc, Handled);
		return result;
	}
	STDMETHOD(MailMergeWizardSendToCustom)(Document * Doc)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x17, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc);
		return result;
	}
	STDMETHOD(MailMergeWizardStateChange)(Document * Doc, int * FromState, int * ToState, BOOL * Handled)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PI4 VTS_PI4 VTS_PBOOL ;
		InvokeHelper(0x18, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc, FromState, ToState, Handled);
		return result;
	}
	STDMETHOD(WindowSize)(Document * Doc, Window * Wn)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x19, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc, Wn);
		return result;
	}
	STDMETHOD(XMLSelectionChange)(Selection * Sel, XMLNode * OldXMLNode, XMLNode * NewXMLNode, long * Reason)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_DISPATCH VTS_PI4 ;
		InvokeHelper(0x1a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Sel, OldXMLNode, NewXMLNode, Reason);
		return result;
	}
	STDMETHOD(XMLValidationError)(XMLNode * XMLNode)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x1b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, XMLNode);
		return result;
	}
	STDMETHOD(DocumentSync)(Document * Doc, MsoSyncEventType SyncEventType)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_I4 ;
		InvokeHelper(0x1c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc, SyncEventType);
		return result;
	}
	STDMETHOD(EPostageInsertEx)(Document * Doc, int cpDeliveryAddrStart, int cpDeliveryAddrEnd, int cpReturnAddrStart, int cpReturnAddrEnd, int xaWidth, int yaHeight, LPCTSTR bstrPrinterName, LPCTSTR bstrPaperFeed, BOOL fPrint, BOOL * fCancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_BSTR VTS_BSTR VTS_BOOL VTS_PBOOL ;
		InvokeHelper(0x1d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc, cpDeliveryAddrStart, cpDeliveryAddrEnd, cpReturnAddrStart, cpReturnAddrEnd, xaWidth, yaHeight, bstrPrinterName, bstrPaperFeed, fPrint, fCancel);
		return result;
	}
	STDMETHOD(MailMergeDataSourceValidate2)(Document * Doc, BOOL * Handled)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x1e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Doc, Handled);
		return result;
	}

	// IApplicationEvents4 属性
public:

};
