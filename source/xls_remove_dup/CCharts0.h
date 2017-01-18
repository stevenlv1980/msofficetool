// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CCharts0 包装类

class CCharts0 : public COleDispatchDriver
{
public:
	CCharts0(){} // 调用 COleDispatchDriver 默认构造函数
	CCharts0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCharts0(const CCharts0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ICharts 方法
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
	STDMETHOD(Add)(VARIANT Before, VARIANT After, VARIANT Count, Chart * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xb5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Before, &After, &Count, RHS);
		return result;
	}
	STDMETHOD(Copy)(VARIANT Before, VARIANT After, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x227, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Before, &After, lcid);
		return result;
	}
	STDMETHOD(get_Count)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Delete)(long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x75, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, lcid);
		return result;
	}
	void _Dummy7()
	{
		InvokeHelper(0x10007, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_Item)(VARIANT Index, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xaa, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &Index, RHS);
		return result;
	}
	STDMETHOD(Move)(VARIANT Before, VARIANT After, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x27d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Before, &After, lcid);
		return result;
	}
	STDMETHOD(get__NewEnum)(LPUNKNOWN * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PUNKNOWN ;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(__PrintOut)(VARIANT From, VARIANT To, VARIANT Copies, VARIANT Preview, VARIANT ActivePrinter, VARIANT PrintToFile, VARIANT Collate, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x389, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate, lcid);
		return result;
	}
	STDMETHOD(PrintPreview)(VARIANT EnableChanges, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x119, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &EnableChanges, lcid);
		return result;
	}
	STDMETHOD(Select)(VARIANT Replace, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 ;
		InvokeHelper(0xeb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Replace, lcid);
		return result;
	}
	STDMETHOD(get_HPageBreaks)(HPageBreaks * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x58a, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_VPageBreaks)(VPageBreaks * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x58b, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Visible)(long lcid, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x22e, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(put_Visible)(long lcid, VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT ;
		InvokeHelper(0x22e, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, lcid, &newValue);
		return result;
	}
	STDMETHOD(get__Default)(VARIANT Index, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &Index, RHS);
		return result;
	}
	STDMETHOD(_PrintOut)(VARIANT From, VARIANT To, VARIANT Copies, VARIANT Preview, VARIANT ActivePrinter, VARIANT PrintToFile, VARIANT Collate, VARIANT PrToFileName, long lcid)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x6ec, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate, &PrToFileName, lcid);
		return result;
	}
	STDMETHOD(PrintOut)(VARIANT From, VARIANT To, VARIANT Copies, VARIANT Preview, VARIANT ActivePrinter, VARIANT PrintToFile, VARIANT Collate, VARIANT PrToFileName)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x939, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate, &PrToFileName);
		return result;
	}

	// ICharts 属性
public:

};
