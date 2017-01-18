// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CDatabar 包装类

class CDatabar : public COleDispatchDriver
{
public:
	CDatabar(){} // 调用 COleDispatchDriver 默认构造函数
	CDatabar(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDatabar(const CDatabar& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Databar 方法
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
	LPDISPATCH get_AppliesTo()
	{
		LPDISPATCH result;
		InvokeHelper(0xa42, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_MinPoint()
	{
		LPDISPATCH result;
		InvokeHelper(0xa9e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_MaxPoint()
	{
		LPDISPATCH result;
		InvokeHelper(0xa9f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_PercentMin()
	{
		long result;
		InvokeHelper(0xaa0, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_PercentMin(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xaa0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_PercentMax()
	{
		long result;
		InvokeHelper(0xaa1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_PercentMax(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xaa1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_BarColor()
	{
		LPDISPATCH result;
		InvokeHelper(0xaa2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_ShowValue()
	{
		BOOL result;
		InvokeHelper(0x7e8, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ShowValue(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x7e8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_Formula()
	{
		CString result;
		InvokeHelper(0x105, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Formula(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x105, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
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

	// Databar 属性
public:

};
