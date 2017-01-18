// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CTop100 包装类

class CTop100 : public COleDispatchDriver
{
public:
	CTop100(){} // 调用 COleDispatchDriver 默认构造函数
	CTop100(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTop100(const CTop100& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Top10 方法
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
	long get_Priority()
	{
		long result;
		InvokeHelper(0x3d9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Priority(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x3d9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_StopIfTrue()
	{
		BOOL result;
		InvokeHelper(0xa41, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_StopIfTrue(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa41, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_AppliesTo()
	{
		LPDISPATCH result;
		InvokeHelper(0xa42, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_TopBottom()
	{
		long result;
		InvokeHelper(0xaa8, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_TopBottom(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xaa8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Rank()
	{
		long result;
		InvokeHelper(0x50a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Rank(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x50a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_Percent()
	{
		BOOL result;
		InvokeHelper(0xaa9, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_Percent(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xaa9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Interior()
	{
		LPDISPATCH result;
		InvokeHelper(0x81, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Borders()
	{
		LPDISPATCH result;
		InvokeHelper(0x1b3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Font()
	{
		LPDISPATCH result;
		InvokeHelper(0x92, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	VARIANT get_NumberFormat()
	{
		VARIANT result;
		InvokeHelper(0xc1, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_NumberFormat(VARIANT newValue)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xc1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
	}
	void SetFirstPriority()
	{
		InvokeHelper(0xa45, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SetLastPriority()
	{
		InvokeHelper(0xa46, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Delete()
	{
		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ModifyAppliesToRange(LPDISPATCH Range)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xa43, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Range);
	}
	BOOL get_PTCondition()
	{
		BOOL result;
		InvokeHelper(0xa47, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	long get_ScopeType()
	{
		long result;
		InvokeHelper(0xa37, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ScopeType(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xa37, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_CalcFor()
	{
		long result;
		InvokeHelper(0xaaa, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_CalcFor(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xaaa, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// Top10 属性
public:

};
