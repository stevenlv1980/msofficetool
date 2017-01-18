// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CListFormat 包装类

class CListFormat : public COleDispatchDriver
{
public:
	CListFormat(){} // 调用 COleDispatchDriver 默认构造函数
	CListFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CListFormat(const CListFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ListFormat 方法
public:
	long get_ListLevelNumber()
	{
		long result;
		InvokeHelper(0x44, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ListLevelNumber(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x44, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_List()
	{
		LPDISPATCH result;
		InvokeHelper(0x45, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ListTemplate()
	{
		LPDISPATCH result;
		InvokeHelper(0x46, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_ListValue()
	{
		long result;
		InvokeHelper(0x47, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	BOOL get_SingleList()
	{
		BOOL result;
		InvokeHelper(0x48, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_SingleListTemplate()
	{
		BOOL result;
		InvokeHelper(0x49, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	long get_ListType()
	{
		long result;
		InvokeHelper(0x4a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	CString get_ListString()
	{
		CString result;
		InvokeHelper(0x4b, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
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
	long CanContinuePreviousList(LPDISPATCH ListTemplate)
	{
		long result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xb8, DISPATCH_METHOD, VT_I4, (void*)&result, parms, ListTemplate);
		return result;
	}
	void RemoveNumbers(VARIANT * NumberType)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xb9, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NumberType);
	}
	void ConvertNumbersToText(VARIANT * NumberType)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xba, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NumberType);
	}
	long CountNumberedItems(VARIANT * NumberType, VARIANT * Level)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0xbb, DISPATCH_METHOD, VT_I4, (void*)&result, parms, NumberType, Level);
		return result;
	}
	void ApplyBulletDefaultOld()
	{
		InvokeHelper(0xbc, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ApplyNumberDefaultOld()
	{
		InvokeHelper(0xbd, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ApplyOutlineNumberDefaultOld()
	{
		InvokeHelper(0xbe, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ApplyListTemplateOld(LPDISPATCH ListTemplate, VARIANT * ContinuePreviousList, VARIANT * ApplyTo)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0xbf, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ListTemplate, ContinuePreviousList, ApplyTo);
	}
	void ListOutdent()
	{
		InvokeHelper(0xd2, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ListIndent()
	{
		InvokeHelper(0xd3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ApplyBulletDefault(VARIANT * DefaultListBehavior)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xd4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, DefaultListBehavior);
	}
	void ApplyNumberDefault(VARIANT * DefaultListBehavior)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xd5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, DefaultListBehavior);
	}
	void ApplyOutlineNumberDefault(VARIANT * DefaultListBehavior)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xd6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, DefaultListBehavior);
	}
	void ApplyListTemplate(LPDISPATCH ListTemplate, VARIANT * ContinuePreviousList, VARIANT * ApplyTo, VARIANT * DefaultListBehavior)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0xd7, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ListTemplate, ContinuePreviousList, ApplyTo, DefaultListBehavior);
	}
	LPDISPATCH get_ListPictureBullet()
	{
		LPDISPATCH result;
		InvokeHelper(0x4c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void ApplyListTemplateWithLevel(LPDISPATCH ListTemplate, VARIANT * ContinuePreviousList, VARIANT * ApplyTo, VARIANT * DefaultListBehavior, VARIANT * ApplyLevel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0xd8, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ListTemplate, ContinuePreviousList, ApplyTo, DefaultListBehavior, ApplyLevel);
	}

	// ListFormat 属性
public:

};
