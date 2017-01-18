// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CChartEvents 包装类

class CChartEvents : public COleDispatchDriver
{
public:
	CChartEvents(){} // 调用 COleDispatchDriver 默认构造函数
	CChartEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CChartEvents(const CChartEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ChartEvents 方法
public:
	void Activate()
	{
		InvokeHelper(0x130, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Deactivate()
	{
		InvokeHelper(0x5fa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Resize()
	{
		InvokeHelper(0x100, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void MouseDown(long Button, long Shift, long x, long y)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 ;
		InvokeHelper(0x5fb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Button, Shift, x, y);
	}
	void MouseUp(long Button, long Shift, long x, long y)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 ;
		InvokeHelper(0x5fc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Button, Shift, x, y);
	}
	void MouseMove(long Button, long Shift, long x, long y)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 ;
		InvokeHelper(0x5fd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Button, Shift, x, y);
	}
	void BeforeRightClick(BOOL * Cancel)
	{
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x5fe, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Cancel);
	}
	void DragPlot()
	{
		InvokeHelper(0x5ff, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void DragOver()
	{
		InvokeHelper(0x600, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void BeforeDoubleClick(long ElementID, long Arg1, long Arg2, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 VTS_PBOOL ;
		InvokeHelper(0x601, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ElementID, Arg1, Arg2, Cancel);
	}
	void Select(long ElementID, long Arg1, long Arg2)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 ;
		InvokeHelper(0xeb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ElementID, Arg1, Arg2);
	}
	void SeriesChange(long SeriesIndex, long PointIndex)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0x602, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SeriesIndex, PointIndex);
	}
	void Calculate()
	{
		InvokeHelper(0x117, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// ChartEvents 属性
public:

};
