// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CGlobal 包装类

class CGlobal : public COleDispatchDriver
{
public:
	CGlobal(){} // 调用 COleDispatchDriver 默认构造函数
	CGlobal(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CGlobal(const CGlobal& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// _Global 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ActiveCell()
	{
		LPDISPATCH result;
		InvokeHelper(0x131, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ActiveChart()
	{
		LPDISPATCH result;
		InvokeHelper(0xb7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ActiveDialog()
	{
		LPDISPATCH result;
		InvokeHelper(0x32f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ActiveMenuBar()
	{
		LPDISPATCH result;
		InvokeHelper(0x2f6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_ActivePrinter()
	{
		CString result;
		InvokeHelper(0x132, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_ActivePrinter(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x132, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_ActiveSheet()
	{
		LPDISPATCH result;
		InvokeHelper(0x133, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ActiveWindow()
	{
		LPDISPATCH result;
		InvokeHelper(0x2f7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ActiveWorkbook()
	{
		LPDISPATCH result;
		InvokeHelper(0x134, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_AddIns()
	{
		LPDISPATCH result;
		InvokeHelper(0x225, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Assistant()
	{
		LPDISPATCH result;
		InvokeHelper(0x59e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void Calculate()
	{
		InvokeHelper(0x117, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH get_Cells()
	{
		LPDISPATCH result;
		InvokeHelper(0xee, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Charts()
	{
		LPDISPATCH result;
		InvokeHelper(0x79, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Columns()
	{
		LPDISPATCH result;
		InvokeHelper(0xf1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_CommandBars()
	{
		LPDISPATCH result;
		InvokeHelper(0x59f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_DDEAppReturnCode()
	{
		long result;
		InvokeHelper(0x14c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void DDEExecute(long Channel, LPCTSTR String)
	{
		static BYTE parms[] = VTS_I4 VTS_BSTR ;
		InvokeHelper(0x14d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Channel, String);
	}
	long DDEInitiate(LPCTSTR App, LPCTSTR Topic)
	{
		long result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR ;
		InvokeHelper(0x14e, DISPATCH_METHOD, VT_I4, (void*)&result, parms, App, Topic);
		return result;
	}
	void DDEPoke(long Channel, VARIANT Item, VARIANT Data)
	{
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x14f, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Channel, &Item, &Data);
	}
	VARIANT DDERequest(long Channel, LPCTSTR Item)
	{
		VARIANT result;
		static BYTE parms[] = VTS_I4 VTS_BSTR ;
		InvokeHelper(0x150, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, Channel, Item);
		return result;
	}
	void DDETerminate(long Channel)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x151, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Channel);
	}
	LPDISPATCH get_DialogSheets()
	{
		LPDISPATCH result;
		InvokeHelper(0x2fc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	VARIANT Evaluate(VARIANT Name)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Name);
		return result;
	}
	VARIANT _Evaluate(VARIANT Name)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xfffffffb, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Name);
		return result;
	}
	VARIANT ExecuteExcel4Macro(LPCTSTR String)
	{
		VARIANT result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x15e, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, String);
		return result;
	}
	LPDISPATCH Intersect(LPDISPATCH Arg1, LPDISPATCH Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x2fe, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Arg1, Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	LPDISPATCH get_MenuBars()
	{
		LPDISPATCH result;
		InvokeHelper(0x24d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Modules()
	{
		LPDISPATCH result;
		InvokeHelper(0x246, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Names()
	{
		LPDISPATCH result;
		InvokeHelper(0x1ba, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Range(VARIANT Cell1, VARIANT Cell2)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0xc5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, &Cell1, &Cell2);
		return result;
	}
	LPDISPATCH get_Rows()
	{
		LPDISPATCH result;
		InvokeHelper(0x102, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	VARIANT Run(VARIANT Macro, VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x103, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Macro, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	VARIANT _Run2(VARIANT Macro, VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x326, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Macro, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	LPDISPATCH get_Selection()
	{
		LPDISPATCH result;
		InvokeHelper(0x93, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void SendKeys(VARIANT Keys, VARIANT Wait)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x17f, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Keys, &Wait);
	}
	LPDISPATCH get_Sheets()
	{
		LPDISPATCH result;
		InvokeHelper(0x1e5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ShortcutMenus(long Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x308, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH get_ThisWorkbook()
	{
		LPDISPATCH result;
		InvokeHelper(0x30a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Toolbars()
	{
		LPDISPATCH result;
		InvokeHelper(0x228, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Union(LPDISPATCH Arg1, LPDISPATCH Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x30b, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Arg1, Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	LPDISPATCH get_Windows()
	{
		LPDISPATCH result;
		InvokeHelper(0x1ae, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Workbooks()
	{
		LPDISPATCH result;
		InvokeHelper(0x23c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_WorksheetFunction()
	{
		LPDISPATCH result;
		InvokeHelper(0x5a0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Worksheets()
	{
		LPDISPATCH result;
		InvokeHelper(0x1ee, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Excel4IntlMacroSheets()
	{
		LPDISPATCH result;
		InvokeHelper(0x245, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Excel4MacroSheets()
	{
		LPDISPATCH result;
		InvokeHelper(0x243, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// _Global 属性
public:

};
