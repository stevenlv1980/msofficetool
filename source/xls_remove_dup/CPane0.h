// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CPane0 包装类

class CPane0 : public COleDispatchDriver
{
public:
	CPane0(){} // 调用 COleDispatchDriver 默认构造函数
	CPane0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CPane0(const CPane0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Pane 方法
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
	BOOL Activate()
	{
		BOOL result;
		InvokeHelper(0x130, DISPATCH_METHOD, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	long get_Index()
	{
		long result;
		InvokeHelper(0x1e6, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	VARIANT LargeScroll(VARIANT Down, VARIANT Up, VARIANT ToRight, VARIANT ToLeft)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x223, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Down, &Up, &ToRight, &ToLeft);
		return result;
	}
	long get_ScrollColumn()
	{
		long result;
		InvokeHelper(0x28e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ScrollColumn(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x28e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_ScrollRow()
	{
		long result;
		InvokeHelper(0x28f, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ScrollRow(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x28f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	VARIANT SmallScroll(VARIANT Down, VARIANT Up, VARIANT ToRight, VARIANT ToLeft)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x224, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Down, &Up, &ToRight, &ToLeft);
		return result;
	}
	LPDISPATCH get_VisibleRange()
	{
		LPDISPATCH result;
		InvokeHelper(0x45e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void ScrollIntoView(long Left, long Top, long Width, long Height, VARIANT Start)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_VARIANT ;
		InvokeHelper(0x6f5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Left, Top, Width, Height, &Start);
	}
	long PointsToScreenPixelsX(long Points)
	{
		long result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x6f0, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Points);
		return result;
	}
	long PointsToScreenPixelsY(long Points)
	{
		long result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x6f1, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Points);
		return result;
	}

	// Pane 属性
public:

};
