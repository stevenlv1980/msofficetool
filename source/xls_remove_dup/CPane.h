// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CPane 包装类

class CPane : public COleDispatchDriver
{
public:
	CPane(){} // 调用 COleDispatchDriver 默认构造函数
	CPane(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CPane(const CPane& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IPane 方法
public:
	STDMETHOD(get_Application)(Application * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Creator)(XlCreator * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Parent)(LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Activate)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x130, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Index)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x1e6, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(LargeScroll)(VARIANT Down, VARIANT Up, VARIANT ToRight, VARIANT ToLeft, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x223, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Down, &Up, &ToRight, &ToLeft, RHS);
		return result;
	}
	STDMETHOD(get_ScrollColumn)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x28e, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_ScrollColumn)(long RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x28e, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_ScrollRow)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x28f, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_ScrollRow)(long RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x28f, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(SmallScroll)(VARIANT Down, VARIANT Up, VARIANT ToRight, VARIANT ToLeft, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x224, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Down, &Up, &ToRight, &ToLeft, RHS);
		return result;
	}
	STDMETHOD(get_VisibleRange)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x45e, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(ScrollIntoView)(long Left, long Top, long Width, long Height, VARIANT Start)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_VARIANT ;
		InvokeHelper(0x6f5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Left, Top, Width, Height, &Start);
		return result;
	}
	STDMETHOD(PointsToScreenPixelsX)(long Points, long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PI4 ;
		InvokeHelper(0x6f0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Points, RHS);
		return result;
	}
	STDMETHOD(PointsToScreenPixelsY)(long Points, long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PI4 ;
		InvokeHelper(0x6f1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Points, RHS);
		return result;
	}

	// IPane 属性
public:

};
