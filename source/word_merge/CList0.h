// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CList0 包装类

class CList0 : public COleDispatchDriver
{
public:
	CList0(){} // 调用 COleDispatchDriver 默认构造函数
	CList0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CList0(const CList0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// List 方法
public:
	LPDISPATCH get_Range()
	{
		LPDISPATCH result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ListParagraphs()
	{
		LPDISPATCH result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_SingleListTemplate()
	{
		BOOL result;
		InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void ConvertNumbersToText(VARIANT * NumberType)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NumberType);
	}
	void RemoveNumbers(VARIANT * NumberType)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NumberType);
	}
	long CountNumberedItems(VARIANT * NumberType, VARIANT * Level)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x67, DISPATCH_METHOD, VT_I4, (void*)&result, parms, NumberType, Level);
		return result;
	}
	void ApplyListTemplateOld(LPDISPATCH ListTemplate, VARIANT * ContinuePreviousList)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT ;
		InvokeHelper(0x68, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ListTemplate, ContinuePreviousList);
	}
	long CanContinuePreviousList(LPDISPATCH ListTemplate)
	{
		long result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x69, DISPATCH_METHOD, VT_I4, (void*)&result, parms, ListTemplate);
		return result;
	}
	void ApplyListTemplate(LPDISPATCH ListTemplate, VARIANT * ContinuePreviousList, VARIANT * DefaultListBehavior)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x6a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ListTemplate, ContinuePreviousList, DefaultListBehavior);
	}
	CString get_StyleName()
	{
		CString result;
		InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void ApplyListTemplateWithLevel(LPDISPATCH ListTemplate, VARIANT * ContinuePreviousList, VARIANT * DefaultListBehavior, VARIANT * ApplyLevel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x6b, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ListTemplate, ContinuePreviousList, DefaultListBehavior, ApplyLevel);
	}

	// List 属性
public:

};
