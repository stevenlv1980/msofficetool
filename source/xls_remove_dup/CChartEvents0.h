// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CChartEvents0 包装类

class CChartEvents0 : public COleDispatchDriver
{
public:
	CChartEvents0(){} // 调用 COleDispatchDriver 默认构造函数
	CChartEvents0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CChartEvents0(const CChartEvents0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IChartEvents 方法
public:
	STDMETHOD(Activate)()
	{
		HRESULT result;
		InvokeHelper(0x130, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(Deactivate)()
	{
		HRESULT result;
		InvokeHelper(0x5fa, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(Resize)()
	{
		HRESULT result;
		InvokeHelper(0x100, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(MouseDown)(long Button, long Shift, long x, long y)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 ;
		InvokeHelper(0x5fb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Button, Shift, x, y);
		return result;
	}
	STDMETHOD(MouseUp)(long Button, long Shift, long x, long y)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 ;
		InvokeHelper(0x5fc, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Button, Shift, x, y);
		return result;
	}
	STDMETHOD(MouseMove)(long Button, long Shift, long x, long y)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 ;
		InvokeHelper(0x5fd, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Button, Shift, x, y);
		return result;
	}
	STDMETHOD(BeforeRightClick)(BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x5fe, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Cancel);
		return result;
	}
	STDMETHOD(DragPlot)()
	{
		HRESULT result;
		InvokeHelper(0x5ff, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(DragOver)()
	{
		HRESULT result;
		InvokeHelper(0x600, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(BeforeDoubleClick)(long ElementID, long Arg1, long Arg2, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 VTS_PBOOL ;
		InvokeHelper(0x601, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, ElementID, Arg1, Arg2, Cancel);
		return result;
	}
	STDMETHOD(Select)(long ElementID, long Arg1, long Arg2)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 ;
		InvokeHelper(0xeb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, ElementID, Arg1, Arg2);
		return result;
	}
	STDMETHOD(SeriesChange)(long SeriesIndex, long PointIndex)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0x602, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, SeriesIndex, PointIndex);
		return result;
	}
	STDMETHOD(Calculate)()
	{
		HRESULT result;
		InvokeHelper(0x117, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}

	// IChartEvents 属性
public:

};
