// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CDialogSheet0 包装类

class CDialogSheet0 : public COleDispatchDriver
{
public:
	CDialogSheet0(){} // 调用 COleDispatchDriver 默认构造函数
	CDialogSheet0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDialogSheet0(const CDialogSheet0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IDialogSheet 方法
public:
	STDMETHOD(get_Application)(Application * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Creator)(XlCreator * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Parent)(LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Activate)(long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x130, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, lcid);
		return result;
	}
	STDMETHOD(Copy)(VARIANT Before, VARIANT After, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x227, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Before, &After, lcid);
		return result;
	}
	STDMETHOD(Delete)(long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x75, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, lcid);
		return result;
	}
	STDMETHOD(get_CodeName)(BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBSTR ;
		InvokeHelper(0x55d, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get__CodeName)(BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBSTR ;
		InvokeHelper(0x80010000, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put__CodeName)(LPCTSTR RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x80010000, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Index)(long lcid, long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PI4 ;
		InvokeHelper(0x1e6, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(Move)(VARIANT Before, VARIANT After, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x27d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Before, &After, lcid);
		return result;
	}
	STDMETHOD(get_Name)(BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBSTR ;
		InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Name)(LPCTSTR RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x6e, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Next)(LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x1f6, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_OnDoubleClick)(long lcid, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PBSTR ;
		InvokeHelper(0x274, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(put_OnDoubleClick)(long lcid, LPCTSTR RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_BSTR ;
		InvokeHelper(0x274, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, lcid, newValue);
		return result;
	}
	STDMETHOD(get_OnSheetActivate)(long lcid, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PBSTR ;
		InvokeHelper(0x407, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(put_OnSheetActivate)(long lcid, LPCTSTR RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_BSTR ;
		InvokeHelper(0x407, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, lcid, newValue);
		return result;
	}
	STDMETHOD(get_OnSheetDeactivate)(long lcid, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PBSTR ;
		InvokeHelper(0x439, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(put_OnSheetDeactivate)(long lcid, LPCTSTR RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_BSTR ;
		InvokeHelper(0x439, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, lcid, newValue);
		return result;
	}
	STDMETHOD(get_PageSetup)(PageSetup * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x3e6, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Previous)(LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x1f7, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(__PrintOut)(VARIANT From, VARIANT To, VARIANT Copies, VARIANT Preview, VARIANT ActivePrinter, VARIANT PrintToFile, VARIANT Collate, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x389, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate, lcid);
		return result;
	}
	STDMETHOD(PrintPreview)(VARIANT EnableChanges, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x119, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &EnableChanges, lcid);
		return result;
	}
	STDMETHOD(_Protect)(VARIANT Password, VARIANT DrawingObjects, VARIANT Contents, VARIANT Scenarios, VARIANT UserInterfaceOnly, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x11a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Password, &DrawingObjects, &Contents, &Scenarios, &UserInterfaceOnly, lcid);
		return result;
	}
	STDMETHOD(get_ProtectContents)(long lcid, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PBOOL ;
		InvokeHelper(0x124, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(get_ProtectDrawingObjects)(long lcid, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PBOOL ;
		InvokeHelper(0x125, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(get_ProtectionMode)(long lcid, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PBOOL ;
		InvokeHelper(0x487, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(get_ProtectScenarios)(long lcid, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PBOOL ;
		InvokeHelper(0x126, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(_SaveAs)(LPCTSTR Filename, VARIANT FileFormat, VARIANT Password, VARIANT WriteResPassword, VARIANT ReadOnlyRecommended, VARIANT CreateBackup, VARIANT AddToMru, VARIANT TextCodepage, VARIANT TextVisualLayout, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x11c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Filename, &FileFormat, &Password, &WriteResPassword, &ReadOnlyRecommended, &CreateBackup, &AddToMru, &TextCodepage, &TextVisualLayout, lcid);
		return result;
	}
	STDMETHOD(Select)(VARIANT Replace, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 ;
		InvokeHelper(0xeb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Replace, lcid);
		return result;
	}
	STDMETHOD(Unprotect)(VARIANT Password, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x11d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Password, lcid);
		return result;
	}
	STDMETHOD(get_Visible)(long lcid, XlSheetVisibility * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PI4 ;
		InvokeHelper(0x22e, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(put_Visible)(long lcid, XlSheetVisibility RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0x22e, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, lcid, newValue);
		return result;
	}
	STDMETHOD(get_Shapes)(Shapes * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x561, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	void _Dummy29()
	{
		InvokeHelper(0x1001d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(Arcs)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x2f8, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	void _Dummy31()
	{
		InvokeHelper(0x1001f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy32()
	{
		InvokeHelper(0x10020, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(Buttons)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x22d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	void _Dummy34()
	{
		InvokeHelper(0x10022, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_EnableCalculation)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x590, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_EnableCalculation)(BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x590, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	void _Dummy36()
	{
		InvokeHelper(0x10024, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(ChartObjects)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x424, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	STDMETHOD(CheckBoxes)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x338, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	STDMETHOD(CheckSpelling)(VARIANT CustomDictionary, VARIANT IgnoreUppercase, VARIANT AlwaysSuggest, VARIANT SpellLang, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x1f9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &CustomDictionary, &IgnoreUppercase, &AlwaysSuggest, &SpellLang, lcid);
		return result;
	}
	void _Dummy40()
	{
		InvokeHelper(0x10028, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy41()
	{
		InvokeHelper(0x10029, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy42()
	{
		InvokeHelper(0x1002a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy43()
	{
		InvokeHelper(0x1002b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy44()
	{
		InvokeHelper(0x1002c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy45()
	{
		InvokeHelper(0x1002d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_DisplayAutomaticPageBreaks)(long lcid, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PBOOL ;
		InvokeHelper(0x283, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(put_DisplayAutomaticPageBreaks)(long lcid, BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_BOOL ;
		InvokeHelper(0x283, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, lcid, newValue);
		return result;
	}
	STDMETHOD(Drawings)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x304, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	STDMETHOD(DrawingObjects)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x58, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	STDMETHOD(DropDowns)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x344, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	STDMETHOD(get_EnableAutoFilter)(long lcid, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PBOOL ;
		InvokeHelper(0x484, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(put_EnableAutoFilter)(long lcid, BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_BOOL ;
		InvokeHelper(0x484, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, lcid, newValue);
		return result;
	}
	STDMETHOD(get_EnableSelection)(XlEnableSelection * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x591, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_EnableSelection)(XlEnableSelection RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x591, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_EnableOutlining)(long lcid, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PBOOL ;
		InvokeHelper(0x485, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(put_EnableOutlining)(long lcid, BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_BOOL ;
		InvokeHelper(0x485, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, lcid, newValue);
		return result;
	}
	STDMETHOD(get_EnablePivotTable)(long lcid, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PBOOL ;
		InvokeHelper(0x486, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(put_EnablePivotTable)(long lcid, BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_BOOL ;
		InvokeHelper(0x486, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, lcid, newValue);
		return result;
	}
	STDMETHOD(Evaluate)(VARIANT Name, long lcid, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Name, lcid, RHS);
		return result;
	}
	STDMETHOD(_Evaluate)(VARIANT Name, long lcid, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0xfffffffb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Name, lcid, RHS);
		return result;
	}
	void _Dummy56()
	{
		InvokeHelper(0x10038, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(ResetAllPageBreaks)()
	{
		HRESULT result;
		InvokeHelper(0x592, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(GroupBoxes)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x342, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	STDMETHOD(GroupObjects)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x459, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	STDMETHOD(Labels)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x349, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	STDMETHOD(Lines)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x2ff, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	STDMETHOD(ListBoxes)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x340, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	STDMETHOD(get_Names)(Names * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x1ba, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(OLEObjects)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x31f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	void _Dummy65()
	{
		InvokeHelper(0x10041, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy66()
	{
		InvokeHelper(0x10042, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy67()
	{
		InvokeHelper(0x10043, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(OptionButtons)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x33a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	void _Dummy69()
	{
		InvokeHelper(0x10045, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(Ovals)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x321, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	STDMETHOD(Paste)(VARIANT Destination, VARIANT Link, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0xd3, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Destination, &Link, lcid);
		return result;
	}
	STDMETHOD(_PasteSpecial)(VARIANT Format, VARIANT Link, VARIANT DisplayAsIcon, VARIANT IconFileName, VARIANT IconIndex, VARIANT IconLabel, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x403, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Format, &Link, &DisplayAsIcon, &IconFileName, &IconIndex, &IconLabel, lcid);
		return result;
	}
	STDMETHOD(Pictures)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x303, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	void _Dummy74()
	{
		InvokeHelper(0x1004a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy75()
	{
		InvokeHelper(0x1004b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy76()
	{
		InvokeHelper(0x1004c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(Rectangles)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x306, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	void _Dummy78()
	{
		InvokeHelper(0x1004e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy79()
	{
		InvokeHelper(0x1004f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_ScrollArea)(BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBSTR ;
		InvokeHelper(0x599, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_ScrollArea)(LPCTSTR RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x599, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(ScrollBars)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x33e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	void _Dummy82()
	{
		InvokeHelper(0x10052, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy83()
	{
		InvokeHelper(0x10053, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(Spinners)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x346, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	void _Dummy85()
	{
		InvokeHelper(0x10055, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy86()
	{
		InvokeHelper(0x10056, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(TextBoxes)(VARIANT Index, long lcid, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x309, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, lcid, RHS);
		return result;
	}
	void _Dummy88()
	{
		InvokeHelper(0x10058, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy89()
	{
		InvokeHelper(0x10059, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy90()
	{
		InvokeHelper(0x1005a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_HPageBreaks)(HPageBreaks * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x58a, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_VPageBreaks)(VPageBreaks * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x58b, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_QueryTables)(QueryTables * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x59a, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_DisplayPageBreaks)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x59b, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_DisplayPageBreaks)(BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x59b, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Comments)(Comments * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x23f, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Hyperlinks)(Hyperlinks * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x571, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(ClearCircles)()
	{
		HRESULT result;
		InvokeHelper(0x59c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(CircleInvalid)()
	{
		HRESULT result;
		InvokeHelper(0x59d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(get__DisplayRightToLeft)(long lcid, long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PI4 ;
		InvokeHelper(0x288, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(put__DisplayRightToLeft)(long lcid, long RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0x288, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, lcid, newValue);
		return result;
	}
	STDMETHOD(get_AutoFilter)(AutoFilter * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x319, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_DisplayRightToLeft)(long lcid, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PBOOL ;
		InvokeHelper(0x6ee, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(put_DisplayRightToLeft)(long lcid, BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_BOOL ;
		InvokeHelper(0x6ee, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, lcid, newValue);
		return result;
	}
	STDMETHOD(get_Scripts)(Scripts * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x718, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(_PrintOut)(VARIANT From, VARIANT To, VARIANT Copies, VARIANT Preview, VARIANT ActivePrinter, VARIANT PrintToFile, VARIANT Collate, VARIANT PrToFileName, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x6ec, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate, &PrToFileName, lcid);
		return result;
	}
	STDMETHOD(_CheckSpelling)(VARIANT CustomDictionary, VARIANT IgnoreUppercase, VARIANT AlwaysSuggest, VARIANT SpellLang, VARIANT IgnoreFinalYaa, VARIANT SpellScript, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x719, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &CustomDictionary, &IgnoreUppercase, &AlwaysSuggest, &SpellLang, &IgnoreFinalYaa, &SpellScript, lcid);
		return result;
	}
	STDMETHOD(get_Tab)(Tab * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x411, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_MailEnvelope)(MsoEnvelope * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x7e5, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(SaveAs)(LPCTSTR Filename, VARIANT FileFormat, VARIANT Password, VARIANT WriteResPassword, VARIANT ReadOnlyRecommended, VARIANT CreateBackup, VARIANT AddToMru, VARIANT TextCodepage, VARIANT TextVisualLayout, VARIANT Local)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x785, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Filename, &FileFormat, &Password, &WriteResPassword, &ReadOnlyRecommended, &CreateBackup, &AddToMru, &TextCodepage, &TextVisualLayout, &Local);
		return result;
	}
	STDMETHOD(get_CustomProperties)(CustomProperties * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x7ee, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_SmartTags)(SmartTags * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x7e0, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Protection)(Protection * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0xb0, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(PasteSpecial)(VARIANT Format, VARIANT Link, VARIANT DisplayAsIcon, VARIANT IconFileName, VARIANT IconIndex, VARIANT IconLabel, VARIANT NoHTMLFormatting, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x788, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Format, &Link, &DisplayAsIcon, &IconFileName, &IconIndex, &IconLabel, &NoHTMLFormatting, lcid);
		return result;
	}
	STDMETHOD(Protect)(VARIANT Password, VARIANT DrawingObjects, VARIANT Contents, VARIANT Scenarios, VARIANT UserInterfaceOnly, VARIANT AllowFormattingCells, VARIANT AllowFormattingColumns, VARIANT AllowFormattingRows, VARIANT AllowInsertingColumns, VARIANT AllowInsertingRows, VARIANT AllowInsertingHyperlinks, VARIANT AllowDeletingColumns, VARIANT AllowDeletingRows, VARIANT AllowSorting, VARIANT AllowFiltering, VARIANT AllowUsingPivotTables)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x7ed, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Password, &DrawingObjects, &Contents, &Scenarios, &UserInterfaceOnly, &AllowFormattingCells, &AllowFormattingColumns, &AllowFormattingRows, &AllowInsertingColumns, &AllowInsertingRows, &AllowInsertingHyperlinks, &AllowDeletingColumns, &AllowDeletingRows, &AllowSorting, &AllowFiltering, &AllowUsingPivotTables);
		return result;
	}
	void _Dummy113()
	{
		InvokeHelper(0x10071, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy114()
	{
		InvokeHelper(0x10072, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy115()
	{
		InvokeHelper(0x10073, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(PrintOut)(VARIANT From, VARIANT To, VARIANT Copies, VARIANT Preview, VARIANT ActivePrinter, VARIANT PrintToFile, VARIANT Collate, VARIANT PrToFileName)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x939, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate, &PrToFileName);
		return result;
	}
	STDMETHOD(get_EnableFormatConditionsCalculation)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x9cf, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_EnableFormatConditionsCalculation)(BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x9cf, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Sort)(Sort * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x370, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(ExportAsFixedFormat)(XlFixedFormatType Type, VARIANT Filename, VARIANT Quality, VARIANT IncludeDocProperties, VARIANT IgnorePrintAreas, VARIANT From, VARIANT To, VARIANT OpenAfterPublish, VARIANT FixedFormatExtClassPtr)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x9bd, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Type, &Filename, &Quality, &IncludeDocProperties, &IgnorePrintAreas, &From, &To, &OpenAfterPublish, &FixedFormatExtClassPtr);
		return result;
	}
	STDMETHOD(get_DefaultButton)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x359, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_DefaultButton)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x359, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_DialogFrame)(DialogFrame * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x347, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(EditBoxes)(VARIANT Index, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x33c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, RHS);
		return result;
	}
	STDMETHOD(get_Focus)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x32e, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Focus)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x32e, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(Hide)(VARIANT Cancel, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBOOL ;
		InvokeHelper(0x32d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Cancel, RHS);
		return result;
	}
	STDMETHOD(Show)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x1f0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}

	// IDialogSheet 属性
public:

};
