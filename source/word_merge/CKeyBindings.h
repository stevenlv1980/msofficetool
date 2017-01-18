// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CKeyBindings 包装类

class CKeyBindings : public COleDispatchDriver
{
public:
	CKeyBindings(){} // 调用 COleDispatchDriver 默认构造函数
	CKeyBindings(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CKeyBindings(const CKeyBindings& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// KeyBindings 方法
public:
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
	LPDISPATCH get_Context()
	{
		LPDISPATCH result;
		InvokeHelper(0xa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(long Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH Add(long KeyCategory, LPCTSTR Command, long KeyCode, VARIANT * KeyCode2, VARIANT * CommandParameter)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_BSTR VTS_I4 VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, KeyCategory, Command, KeyCode, KeyCode2, CommandParameter);
		return result;
	}
	void ClearAll()
	{
		InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH Key(long KeyCode, VARIANT * KeyCode2)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x6e, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, KeyCode, KeyCode2);
		return result;
	}

	// KeyBindings 属性
public:

};
