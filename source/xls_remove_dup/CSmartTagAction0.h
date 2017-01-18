// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CSmartTagAction0 包装类

class CSmartTagAction0 : public COleDispatchDriver
{
public:
	CSmartTagAction0(){} // 调用 COleDispatchDriver 默认构造函数
	CSmartTagAction0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSmartTagAction0(const CSmartTagAction0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// SmartTagAction 方法
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
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void Execute()
	{
		InvokeHelper(0x8a3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	CString get__Default()
	{
		CString result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	BOOL get_PresentInPane()
	{
		BOOL result;
		InvokeHelper(0x8f9, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_ExpandHelp()
	{
		BOOL result;
		InvokeHelper(0x8fa, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ExpandHelp(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x8fa, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_CheckboxState()
	{
		BOOL result;
		InvokeHelper(0x8fb, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_CheckboxState(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x8fb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_TextboxText()
	{
		CString result;
		InvokeHelper(0x8fc, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_TextboxText(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x8fc, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_ListSelection()
	{
		long result;
		InvokeHelper(0x8fd, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ListSelection(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x8fd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_RadioGroupSelection()
	{
		long result;
		InvokeHelper(0x8fe, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_RadioGroupSelection(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x8fe, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_ActiveXControl()
	{
		LPDISPATCH result;
		InvokeHelper(0x8ff, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// SmartTagAction 属性
public:

};
