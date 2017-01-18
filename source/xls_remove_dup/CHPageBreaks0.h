// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CHPageBreaks0 包装类

class CHPageBreaks0 : public COleDispatchDriver
{
public:
	CHPageBreaks0(){} // 调用 COleDispatchDriver 默认构造函数
	CHPageBreaks0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CHPageBreaks0(const CHPageBreaks0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IHPageBreaks 方法
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
	STDMETHOD(get_Count)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Item)(long Index, HPageBreak * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0xaa, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, Index, RHS);
		return result;
	}
	STDMETHOD(get__Default)(long Index, HPageBreak * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, Index, RHS);
		return result;
	}
	STDMETHOD(get__NewEnum)(LPUNKNOWN * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PUNKNOWN ;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Add)(LPDISPATCH Before, HPageBreak * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PDISPATCH ;
		InvokeHelper(0xb5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Before, RHS);
		return result;
	}

	// IHPageBreaks 属性
public:

};
