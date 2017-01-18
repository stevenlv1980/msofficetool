// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CApplicationEvents30 包装类

class CApplicationEvents30 : public COleDispatchDriver
{
public:
	CApplicationEvents30(){} // 调用 COleDispatchDriver 默认构造函数
	CApplicationEvents30(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplicationEvents30(const CApplicationEvents30& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IApplicationEvents3 方法
public:
	void Startup()
	{
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Quit()
	{
		InvokeHelper(0x2, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void DocumentChange()
	{
		InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void DocumentOpen(LPDISPATCH Doc)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc);
	}
	void DocumentBeforeClose(LPDISPATCH Doc, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, Cancel);
	}
	void DocumentBeforePrint(LPDISPATCH Doc, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, Cancel);
	}
	void DocumentBeforeSave(LPDISPATCH Doc, BOOL * SaveAsUI, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL VTS_PBOOL ;
		InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, SaveAsUI, Cancel);
	}
	void NewDocument(LPDISPATCH Doc)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc);
	}
	void WindowActivate(LPDISPATCH Doc, LPDISPATCH Wn)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, Wn);
	}
	void WindowDeactivate(LPDISPATCH Doc, LPDISPATCH Wn)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, Wn);
	}
	void WindowSelectionChange(LPDISPATCH Sel)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sel);
	}
	void WindowBeforeRightClick(LPDISPATCH Sel, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0xd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sel, Cancel);
	}
	void WindowBeforeDoubleClick(LPDISPATCH Sel, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0xe, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sel, Cancel);
	}
	void EPostagePropertyDialog(LPDISPATCH Doc)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xf, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc);
	}
	void EPostageInsert(LPDISPATCH Doc)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x10, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc);
	}
	void MailMergeAfterMerge(LPDISPATCH Doc, LPDISPATCH DocResult)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x11, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, DocResult);
	}
	void MailMergeAfterRecordMerge(LPDISPATCH Doc)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x12, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc);
	}
	void MailMergeBeforeMerge(LPDISPATCH Doc, long StartRecord, long EndRecord, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_I4 VTS_I4 VTS_PBOOL ;
		InvokeHelper(0x13, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, StartRecord, EndRecord, Cancel);
	}
	void MailMergeBeforeRecordMerge(LPDISPATCH Doc, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x14, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, Cancel);
	}
	void MailMergeDataSourceLoad(LPDISPATCH Doc)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x15, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc);
	}
	void MailMergeDataSourceValidate(LPDISPATCH Doc, BOOL * Handled)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x16, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, Handled);
	}
	void MailMergeWizardSendToCustom(LPDISPATCH Doc)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x17, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc);
	}
	void MailMergeWizardStateChange(LPDISPATCH Doc, long * FromState, long * ToState, BOOL * Handled)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PI4 VTS_PI4 VTS_PBOOL ;
		InvokeHelper(0x18, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, FromState, ToState, Handled);
	}
	void WindowSize(LPDISPATCH Doc, LPDISPATCH Wn)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x19, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, Wn);
	}

	// IApplicationEvents3 属性
public:

};
