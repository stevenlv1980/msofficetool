// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CAllowEditRanges 包装类

class CAllowEditRanges : public COleDispatchDriver
{
public:
	CAllowEditRanges(){} // 调用 COleDispatchDriver 默认构造函数
	CAllowEditRanges(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAllowEditRanges(const CAllowEditRanges& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// AllowEditRanges 方法
public:
	long get_Count()
	{
		long result;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Item(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xaa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPDISPATCH Add(LPCTSTR Title, LPDISPATCH Range, VARIANT Password)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_DISPATCH VTS_VARIANT ;
		InvokeHelper(0xb5, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Title, Range, &Password);
		return result;
	}
	LPDISPATCH get__Default(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}

	// AllowEditRanges 属性
public:

};
