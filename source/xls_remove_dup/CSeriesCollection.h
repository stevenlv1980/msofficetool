// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CSeriesCollection 包装类

class CSeriesCollection : public COleDispatchDriver
{
public:
	CSeriesCollection(){} // 调用 COleDispatchDriver 默认构造函数
	CSeriesCollection(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSeriesCollection(const CSeriesCollection& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ISeriesCollection 方法
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
	STDMETHOD(Add)(VARIANT Source, XlRowCol Rowcol, VARIANT SeriesLabels, VARIANT CategoryLabels, VARIANT Replace, Series * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xb5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Source, Rowcol, &SeriesLabels, &CategoryLabels, &Replace, RHS);
		return result;
	}
	STDMETHOD(get_Count)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Extend)(VARIANT Source, VARIANT Rowcol, VARIANT CategoryLabels, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0xe3, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Source, &Rowcol, &CategoryLabels, RHS);
		return result;
	}
	STDMETHOD(Item)(VARIANT Index, Series * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xaa, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, RHS);
		return result;
	}
	STDMETHOD(_NewEnum)(LPUNKNOWN * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PUNKNOWN ;
		InvokeHelper(0xfffffffc, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Paste)(XlRowCol Rowcol, VARIANT SeriesLabels, VARIANT CategoryLabels, VARIANT Replace, VARIANT NewSeries, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0xd3, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Rowcol, &SeriesLabels, &CategoryLabels, &Replace, &NewSeries, RHS);
		return result;
	}
	STDMETHOD(NewSeries)(Series * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x45d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(_Default)(VARIANT Index, Series * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, RHS);
		return result;
	}

	// ISeriesCollection 属性
public:

};
