// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CChartView 包装类

class CChartView : public COleDispatchDriver
{
public:
	CChartView(){} // 调用 COleDispatchDriver 默认构造函数
	CChartView(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CChartView(const CChartView& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ChartView 方法
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
	LPDISPATCH get_Sheet()
	{
		LPDISPATCH result;
		InvokeHelper(0x2ef, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// ChartView 属性
public:

};
