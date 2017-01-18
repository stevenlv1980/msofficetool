// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CWindows 包装类

class CWindows : public COleDispatchDriver
{
public:
	CWindows(){} // 调用 COleDispatchDriver 默认构造函数
	CWindows(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWindows(const CWindows& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Windows 方法
public:
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
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
	LPDISPATCH Item(VARIANT * Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH Add(VARIANT * Window)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xa, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Window);
		return result;
	}
	void Arrange(VARIANT * ArrangeStyle)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ArrangeStyle);
	}
	BOOL CompareSideBySideWith(VARIANT * Document)
	{
		BOOL result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xc, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Document);
		return result;
	}
	BOOL BreakSideBySide()
	{
		BOOL result;
		InvokeHelper(0xd, DISPATCH_METHOD, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void ResetPositionsSideBySide()
	{
		InvokeHelper(0xe, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	BOOL get_SyncScrollingSideBySide()
	{
		BOOL result;
		InvokeHelper(0x3eb, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_SyncScrollingSideBySide(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x3eb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// Windows 属性
public:

};
