// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CAppEvents 包装类

class CAppEvents : public COleDispatchDriver
{
public:
	CAppEvents(){} // 调用 COleDispatchDriver 默认构造函数
	CAppEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAppEvents(const CAppEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// AppEvents 方法
public:
	void NewWorkbook(LPDISPATCH Wb)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x61d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb);
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
	void WorkbookOpen(LPDISPATCH Wb)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x61f, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb);
	}
	void WorkbookActivate(LPDISPATCH Wb)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x620, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb);
	}
	void WorkbookDeactivate(LPDISPATCH Wb)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x621, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb);
	}
	void WorkbookBeforeClose(LPDISPATCH Wb, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x622, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, Cancel);
	}
	void WorkbookBeforeSave(LPDISPATCH Wb, BOOL SaveAsUI, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BOOL VTS_PBOOL ;
		InvokeHelper(0x623, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, SaveAsUI, Cancel);
	}
	void WorkbookBeforePrint(LPDISPATCH Wb, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x624, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, Cancel);
	}
	void WorkbookNewSheet(LPDISPATCH Wb, LPDISPATCH Sh)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x625, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, Sh);
	}
	void WorkbookAddinInstall(LPDISPATCH Wb)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x626, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb);
	}
	void WorkbookAddinUninstall(LPDISPATCH Wb)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x627, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb);
	}
	void WindowResize(LPDISPATCH Wb, LPDISPATCH Wn)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x612, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, Wn);
	}
	void WindowActivate(LPDISPATCH Wb, LPDISPATCH Wn)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x614, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, Wn);
	}
	void WindowDeactivate(LPDISPATCH Wb, LPDISPATCH Wn)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x615, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, Wn);
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
	void WorkbookPivotTableCloseConnection(LPDISPATCH Wb, LPDISPATCH Target)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x870, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, Target);
	}
	void WorkbookPivotTableOpenConnection(LPDISPATCH Wb, LPDISPATCH Target)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x871, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, Target);
	}
	void WorkbookSync(LPDISPATCH Wb, long SyncEventType)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_I4 ;
		InvokeHelper(0x8f1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, SyncEventType);
	}
	void WorkbookBeforeXmlImport(LPDISPATCH Wb, LPDISPATCH Map, LPCTSTR Url, BOOL IsRefresh, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_BSTR VTS_BOOL VTS_PBOOL ;
		InvokeHelper(0x8f2, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, Map, Url, IsRefresh, Cancel);
	}
	void WorkbookAfterXmlImport(LPDISPATCH Wb, LPDISPATCH Map, BOOL IsRefresh, long Result)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_BOOL VTS_I4 ;
		InvokeHelper(0x8f3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, Map, IsRefresh, Result);
	}
	void WorkbookBeforeXmlExport(LPDISPATCH Wb, LPDISPATCH Map, LPCTSTR Url, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_BSTR VTS_PBOOL ;
		InvokeHelper(0x8f4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, Map, Url, Cancel);
	}
	void WorkbookAfterXmlExport(LPDISPATCH Wb, LPDISPATCH Map, LPCTSTR Url, long Result)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_BSTR VTS_I4 ;
		InvokeHelper(0x8f5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, Map, Url, Result);
	}
	void WorkbookRowsetComplete(LPDISPATCH Wb, LPCTSTR Description, LPCTSTR Sheet, BOOL Success)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_BSTR VTS_BOOL ;
		InvokeHelper(0xa33, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Wb, Description, Sheet, Success);
	}
	void AfterCalculate()
	{
		InvokeHelper(0xa34, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// AppEvents 属性
public:

};
