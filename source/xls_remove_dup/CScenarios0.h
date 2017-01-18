// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CScenarios0 包装类

class CScenarios0 : public COleDispatchDriver
{
public:
	CScenarios0(){} // 调用 COleDispatchDriver 默认构造函数
	CScenarios0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CScenarios0(const CScenarios0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Scenarios 方法
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
	LPDISPATCH Add(LPCTSTR Name, VARIANT ChangingCells, VARIANT Values, VARIANT Comment, VARIANT Locked, VARIANT Hidden)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0xb5, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, &ChangingCells, &Values, &Comment, &Locked, &Hidden);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	VARIANT CreateSummary(long ReportType, VARIANT ResultCells)
	{
		VARIANT result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT ;
		InvokeHelper(0x391, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, ReportType, &ResultCells);
		return result;
	}
	LPDISPATCH Item(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xaa, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	VARIANT Merge(VARIANT Source)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x234, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Source);
		return result;
	}
	LPUNKNOWN _NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_METHOD, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}

	// Scenarios 属性
public:

};
