// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CAllowEditRanges0 包装类

class CAllowEditRanges0 : public COleDispatchDriver
{
public:
	CAllowEditRanges0(){} // 调用 COleDispatchDriver 默认构造函数
	CAllowEditRanges0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAllowEditRanges0(const CAllowEditRanges0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IAllowEditRanges 方法
public:
	STDMETHOD(get_Count)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Item)(VARIANT Index, AllowEditRange * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xaa, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &Index, RHS);
		return result;
	}
	STDMETHOD(Add)(LPCTSTR Title, Range * Range, VARIANT Password, AllowEditRange * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_DISPATCH VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xb5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Title, Range, &Password, RHS);
		return result;
	}
	STDMETHOD(get__Default)(VARIANT Index, AllowEditRange * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &Index, RHS);
		return result;
	}
	STDMETHOD(get__NewEnum)(LPUNKNOWN * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PUNKNOWN ;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}

	// IAllowEditRanges 属性
public:

};
