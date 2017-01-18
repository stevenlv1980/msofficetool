// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CSort0 包装类

class CSort0 : public COleDispatchDriver
{
public:
	CSort0(){} // 调用 COleDispatchDriver 默认构造函数
	CSort0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSort0(const CSort0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Sort 方法
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
	LPDISPATCH get_Rng()
	{
		LPDISPATCH result;
		InvokeHelper(0xabc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Header()
	{
		long result;
		InvokeHelper(0x37f, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Header(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x37f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_MatchCase()
	{
		BOOL result;
		InvokeHelper(0x1aa, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_MatchCase(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x1aa, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Orientation()
	{
		long result;
		InvokeHelper(0x86, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Orientation(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x86, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_SortMethod()
	{
		long result;
		InvokeHelper(0x381, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_SortMethod(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x381, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_SortFields()
	{
		LPDISPATCH result;
		InvokeHelper(0xabd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void SetRange(LPDISPATCH Rng)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xabe, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Rng);
	}
	void Apply()
	{
		InvokeHelper(0x68b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// Sort 属性
public:

};
