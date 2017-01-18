// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CWorkbookEvents 包装类

class CWorkbookEvents : public COleDispatchDriver
{
public:
	CWorkbookEvents(){} // 调用 COleDispatchDriver 默认构造函数
	CWorkbookEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWorkbookEvents(const CWorkbookEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IWorkbookEvents 方法
public:
	STDMETHOD(Open)()
	{
		HRESULT result;
		InvokeHelper(0x783, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(Activate)()
	{
		HRESULT result;
		InvokeHelper(0x130, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(Deactivate)()
	{
		HRESULT result;
		InvokeHelper(0x5fa, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(BeforeClose)(BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x60a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Cancel);
		return result;
	}
	STDMETHOD(BeforeSave)(BOOL SaveAsUI, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BOOL VTS_PBOOL ;
		InvokeHelper(0x60b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, SaveAsUI, Cancel);
		return result;
	}
	STDMETHOD(BeforePrint)(BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x60d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Cancel);
		return result;
	}
	STDMETHOD(NewSheet)(LPDISPATCH Sh)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x60e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Sh);
		return result;
	}
	STDMETHOD(AddinInstall)()
	{
		HRESULT result;
		InvokeHelper(0x610, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(AddinUninstall)()
	{
		HRESULT result;
		InvokeHelper(0x611, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(WindowResize)(Window * Wn)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x612, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wn);
		return result;
	}
	STDMETHOD(WindowActivate)(Window * Wn)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x614, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wn);
		return result;
	}
	STDMETHOD(WindowDeactivate)(Window * Wn)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x615, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wn);
		return result;
	}
	STDMETHOD(SheetSelectionChange)(LPDISPATCH Sh, Range * Target)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x616, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Sh, Target);
		return result;
	}
	STDMETHOD(SheetBeforeDoubleClick)(LPDISPATCH Sh, Range * Target, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x617, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Sh, Target, Cancel);
		return result;
	}
	STDMETHOD(SheetBeforeRightClick)(LPDISPATCH Sh, Range * Target, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x618, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Sh, Target, Cancel);
		return result;
	}
	STDMETHOD(SheetActivate)(LPDISPATCH Sh)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x619, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Sh);
		return result;
	}
	STDMETHOD(SheetDeactivate)(LPDISPATCH Sh)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x61a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Sh);
		return result;
	}
	STDMETHOD(SheetCalculate)(LPDISPATCH Sh)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x61b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Sh);
		return result;
	}
	STDMETHOD(SheetChange)(LPDISPATCH Sh, Range * Target)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x61c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Sh, Target);
		return result;
	}
	STDMETHOD(SheetFollowHyperlink)(LPDISPATCH Sh, Hyperlink * Target)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x73e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Sh, Target);
		return result;
	}
	STDMETHOD(SheetPivotTableUpdate)(LPDISPATCH Sh, PivotTable * Target)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x86d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Sh, Target);
		return result;
	}
	STDMETHOD(PivotTableCloseConnection)(PivotTable * Target)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x86e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Target);
		return result;
	}
	STDMETHOD(PivotTableOpenConnection)(PivotTable * Target)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x86f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Target);
		return result;
	}
	STDMETHOD(Sync)(MsoSyncEventType SyncEventType)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x8da, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, SyncEventType);
		return result;
	}
	STDMETHOD(BeforeXmlImport)(XmlMap * Map, LPCTSTR Url, BOOL IsRefresh, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_BOOL VTS_PBOOL ;
		InvokeHelper(0x8eb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Map, Url, IsRefresh, Cancel);
		return result;
	}
	STDMETHOD(AfterXmlImport)(XmlMap * Map, BOOL IsRefresh, XlXmlImportResult Result)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_BOOL VTS_I4 ;
		InvokeHelper(0x8ed, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Map, IsRefresh, Result);
		return result;
	}
	STDMETHOD(BeforeXmlExport)(XmlMap * Map, LPCTSTR Url, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_PBOOL ;
		InvokeHelper(0x8ef, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Map, Url, Cancel);
		return result;
	}
	STDMETHOD(AfterXmlExport)(XmlMap * Map, LPCTSTR Url, XlXmlExportResult Result)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_I4 ;
		InvokeHelper(0x8f0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Map, Url, Result);
		return result;
	}
	STDMETHOD(RowsetComplete)(LPCTSTR Description, LPCTSTR Sheet, BOOL Success)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BOOL ;
		InvokeHelper(0xa32, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Description, Sheet, Success);
		return result;
	}

	// IWorkbookEvents 属性
public:

};
