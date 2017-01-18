// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CHyperlinks 包装类

class CHyperlinks : public COleDispatchDriver
{
public:
	CHyperlinks(){} // 调用 COleDispatchDriver 默认构造函数
	CHyperlinks(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CHyperlinks(const CHyperlinks& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Hyperlinks 方法
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
	long get_Count()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(VARIANT * Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH _Add(LPDISPATCH Anchor, VARIANT * Address, VARIANT * SubAddress)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x64, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Anchor, Address, SubAddress);
		return result;
	}
	LPDISPATCH Add(LPDISPATCH Anchor, VARIANT * Address, VARIANT * SubAddress, VARIANT * ScreenTip, VARIANT * TextToDisplay, VARIANT * Target)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Anchor, Address, SubAddress, ScreenTip, TextToDisplay, Target);
		return result;
	}

	// Hyperlinks 属性
public:

};
