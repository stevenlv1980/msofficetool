// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CAppEvents0 包装类

class CAppEvents0 : public COleDispatchDriver
{
public:
	CAppEvents0(){} // 调用 COleDispatchDriver 默认构造函数
	CAppEvents0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAppEvents0(const CAppEvents0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IAppEvents 方法
public:
	STDMETHOD(NewWorkbook)(Workbook * Wb)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x61d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb);
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
	STDMETHOD(WorkbookOpen)(Workbook * Wb)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x61f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb);
		return result;
	}
	STDMETHOD(WorkbookActivate)(Workbook * Wb)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x620, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb);
		return result;
	}
	STDMETHOD(WorkbookDeactivate)(Workbook * Wb)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x621, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb);
		return result;
	}
	STDMETHOD(WorkbookBeforeClose)(Workbook * Wb, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x622, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, Cancel);
		return result;
	}
	STDMETHOD(WorkbookBeforeSave)(Workbook * Wb, BOOL SaveAsUI, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_BOOL VTS_PBOOL ;
		InvokeHelper(0x623, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, SaveAsUI, Cancel);
		return result;
	}
	STDMETHOD(WorkbookBeforePrint)(Workbook * Wb, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x624, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, Cancel);
		return result;
	}
	STDMETHOD(WorkbookNewSheet)(Workbook * Wb, LPDISPATCH Sh)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x625, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, Sh);
		return result;
	}
	STDMETHOD(WorkbookAddinInstall)(Workbook * Wb)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x626, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb);
		return result;
	}
	STDMETHOD(WorkbookAddinUninstall)(Workbook * Wb)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x627, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb);
		return result;
	}
	STDMETHOD(WindowResize)(Workbook * Wb, Window * Wn)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x612, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, Wn);
		return result;
	}
	STDMETHOD(WindowActivate)(Workbook * Wb, Window * Wn)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x614, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, Wn);
		return result;
	}
	STDMETHOD(WindowDeactivate)(Workbook * Wb, Window * Wn)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x615, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, Wn);
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
	STDMETHOD(WorkbookPivotTableCloseConnection)(Workbook * Wb, PivotTable * Target)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x870, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, Target);
		return result;
	}
	STDMETHOD(WorkbookPivotTableOpenConnection)(Workbook * Wb, PivotTable * Target)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH ;
		InvokeHelper(0x871, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, Target);
		return result;
	}
	STDMETHOD(WorkbookSync)(Workbook * Wb, MsoSyncEventType SyncEventType)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_I4 ;
		InvokeHelper(0x8f1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, SyncEventType);
		return result;
	}
	STDMETHOD(WorkbookBeforeXmlImport)(Workbook * Wb, XmlMap * Map, LPCTSTR Url, BOOL IsRefresh, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_BSTR VTS_BOOL VTS_PBOOL ;
		InvokeHelper(0x8f2, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, Map, Url, IsRefresh, Cancel);
		return result;
	}
	STDMETHOD(WorkbookAfterXmlImport)(Workbook * Wb, XmlMap * Map, BOOL IsRefresh, XlXmlImportResult Result)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_BOOL VTS_I4 ;
		InvokeHelper(0x8f3, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, Map, IsRefresh, Result);
		return result;
	}
	STDMETHOD(WorkbookBeforeXmlExport)(Workbook * Wb, XmlMap * Map, LPCTSTR Url, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_BSTR VTS_PBOOL ;
		InvokeHelper(0x8f4, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, Map, Url, Cancel);
		return result;
	}
	STDMETHOD(WorkbookAfterXmlExport)(Workbook * Wb, XmlMap * Map, LPCTSTR Url, XlXmlExportResult Result)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_BSTR VTS_I4 ;
		InvokeHelper(0x8f5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, Map, Url, Result);
		return result;
	}
	STDMETHOD(WorkbookRowsetComplete)(Workbook * Wb, LPCTSTR Description, LPCTSTR Sheet, BOOL Success)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_BSTR VTS_BOOL ;
		InvokeHelper(0xa33, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Wb, Description, Sheet, Success);
		return result;
	}
	STDMETHOD(AfterCalculate)()
	{
		HRESULT result;
		InvokeHelper(0xa34, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}

	// IAppEvents 属性
public:

};
