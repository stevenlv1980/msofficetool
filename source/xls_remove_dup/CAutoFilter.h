// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CAutoFilter 包装类

class CAutoFilter : public COleDispatchDriver
{
public:
	CAutoFilter(){} // 调用 COleDispatchDriver 默认构造函数
	CAutoFilter(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAutoFilter(const CAutoFilter& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// AutoFilter 方法
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
	LPDISPATCH get_Range()
	{
		LPDISPATCH result;
		InvokeHelper(0xc5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Filters()
	{
		LPDISPATCH result;
		InvokeHelper(0x651, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_FilterMode()
	{
		BOOL result;
		InvokeHelper(0x320, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Sort()
	{
		LPDISPATCH result;
		InvokeHelper(0x370, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void ApplyFilter()
	{
		InvokeHelper(0xa50, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ShowAllData()
	{
		InvokeHelper(0x31a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// AutoFilter 属性
public:

};
