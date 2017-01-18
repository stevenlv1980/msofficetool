// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CModule 包装类

class CModule : public COleDispatchDriver
{
public:
	CModule(){} // 调用 COleDispatchDriver 默认构造函数
	CModule(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CModule(const CModule& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IModule 方法
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
	void _Dummy18()
	{
		InvokeHelper(0x10012, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
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
	void _Dummy21()
	{
		InvokeHelper(0x10015, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_ProtectionMode)(long lcid, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PBOOL ;
		InvokeHelper(0x487, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	void _Dummy23()
	{
		InvokeHelper(0x10017, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
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
	STDMETHOD(InsertFile)(VARIANT Filename, VARIANT Merge, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x248, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Filename, &Merge, RHS);
		return result;
	}
	STDMETHOD(SaveAs)(LPCTSTR Filename, VARIANT FileFormat, VARIANT Password, VARIANT WriteResPassword, VARIANT ReadOnlyRecommended, VARIANT CreateBackup, VARIANT AddToMru, VARIANT TextCodepage, VARIANT TextVisualLayout)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x785, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Filename, &FileFormat, &Password, &WriteResPassword, &ReadOnlyRecommended, &CreateBackup, &AddToMru, &TextCodepage, &TextVisualLayout);
		return result;
	}
	STDMETHOD(Protect)(VARIANT Password, VARIANT DrawingObjects, VARIANT Contents, VARIANT Scenarios, VARIANT UserInterfaceOnly)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x7ed, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Password, &DrawingObjects, &Contents, &Scenarios, &UserInterfaceOnly);
		return result;
	}
	STDMETHOD(_PrintOut)(VARIANT From, VARIANT To, VARIANT Copies, VARIANT Preview, VARIANT ActivePrinter, VARIANT PrintToFile, VARIANT Collate)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x6ec, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate);
		return result;
	}
	STDMETHOD(PrintOut)(VARIANT From, VARIANT To, VARIANT Copies, VARIANT Preview, VARIANT ActivePrinter, VARIANT PrintToFile, VARIANT Collate)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x939, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate);
		return result;
	}

	// IModule 属性
public:

};
