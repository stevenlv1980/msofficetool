// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CWorkbookEvents0 包装类

class CWorkbookEvents0 : public COleDispatchDriver
{
public:
	CWorkbookEvents0(){} // 调用 COleDispatchDriver 默认构造函数
	CWorkbookEvents0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWorkbookEvents0(const CWorkbookEvents0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// WorkbookEvents 方法
public:
	void Open()
	{
		InvokeHelper(0x783, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Activate()
	{
		InvokeHelper(0x130, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Deactivate()
	{
		InvokeHelper(0x5fa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void BeforeClose(BOOL * Cancel)
	{
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x60a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Cancel);
	}
	void BeforeSave(BOOL SaveAsUI, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_BOOL VTS_PBOOL ;
		InvokeHelper(0x60b, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SaveAsUI, Cancel);
	}
	void BeforePrint(BOOL * Cancel)
	{
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x60d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Cancel);
	}
	void NewSheet(LPDISPATCH Sh)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x60e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sh);
	}
	void AddinInstall()
	{
		InvokeHelper(0x610, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void AddinUninstall()
	{
		InvokeHelper(0x611, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void WindowResize(LPDISPATCH Wn)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x612, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wn);
	}
	void WindowActivate(LPDISPATCH Wn)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x614, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wn);
	}
	void WindowDeactivate(LPDISPATCH Wn)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x615, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wn);
	}
	void SheetSelectionChange(LPDISPATCH Sh, LPDISPATCH Target)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x616, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sh, Target);
	}
	void SheetBeforeDoubleClick(LPDISPATCH Sh, LPDISPATCH Target, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x617, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sh, Target, Cancel);
	}
	void SheetBeforeRightClick(LPDISPATCH Sh, LPDISPATCH Target, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x618, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sh, Target, Cancel);
	}
	void SheetActivate(LPDISPATCH Sh)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x619, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sh);
	}
	void SheetDeactivate(LPDISPATCH Sh)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x61a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sh);
	}
	void SheetCalculate(LPDISPATCH Sh)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x61b, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sh);
	}
	void SheetChange(LPDISPATCH Sh, LPDISPATCH Target)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x61c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sh, Target);
	}
	void SheetFollowHyperlink(LPDISPATCH Sh, LPDISPATCH Target)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x73e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sh, Target);
	}
	void SheetPivotTableUpdate(LPDISPATCH Sh, LPDISPATCH Target)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x86d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sh, Target);
	}
	void PivotTableCloseConnection(LPDISPATCH Target)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x86e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Target);
	}
	void PivotTableOpenConnection(LPDISPATCH Target)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x86f, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Target);
	}
	void Sync(long SyncEventType)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x8da, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SyncEventType);
	}
	void BeforeXmlImport(LPDISPATCH Map, LPCTSTR Url, BOOL IsRefresh, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_BOOL VTS_PBOOL ;
		InvokeHelper(0x8eb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Map, Url, IsRefresh, Cancel);
	}
	void AfterXmlImport(LPDISPATCH Map, BOOL IsRefresh, long Result)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BOOL VTS_I4 ;
		InvokeHelper(0x8ed, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Map, IsRefresh, Result);
	}
	void BeforeXmlExport(LPDISPATCH Map, LPCTSTR Url, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_PBOOL ;
		InvokeHelper(0x8ef, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Map, Url, Cancel);
	}
	void AfterXmlExport(LPDISPATCH Map, LPCTSTR Url, long Result)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_I4 ;
		InvokeHelper(0x8f0, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Map, Url, Result);
	}
	void RowsetComplete(LPCTSTR Description, LPCTSTR Sheet, BOOL Success)
	{
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BOOL ;
		InvokeHelper(0xa32, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Description, Sheet, Success);
	}

	// WorkbookEvents 属性
public:

};
