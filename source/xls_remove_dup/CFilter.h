// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CFilter 包装类

class CFilter : public COleDispatchDriver
{
public:
	CFilter(){} // 调用 COleDispatchDriver 默认构造函数
	CFilter(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CFilter(const CFilter& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Filter 方法
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
	BOOL get_On()
	{
		BOOL result;
		InvokeHelper(0x652, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	VARIANT get_Criteria1()
	{
		VARIANT result;
		InvokeHelper(0x31c, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	long get__Operator()
	{
		long result;
		InvokeHelper(0xa51, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	VARIANT get_Criteria2()
	{
		VARIANT result;
		InvokeHelper(0x31e, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	long get_Operator()
	{
		long result;
		InvokeHelper(0x31d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Operator(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x31d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}

	// Filter 属性
public:

};
