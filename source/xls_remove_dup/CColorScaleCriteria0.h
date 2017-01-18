// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CColorScaleCriteria0 包装类

class CColorScaleCriteria0 : public COleDispatchDriver
{
public:
	CColorScaleCriteria0(){} // 调用 COleDispatchDriver 默认构造函数
	CColorScaleCriteria0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CColorScaleCriteria0(const CColorScaleCriteria0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IColorScaleCriteria 方法
public:
	STDMETHOD(get_Count)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get__Default)(VARIANT Index, ColorScaleCriterion * * RHS)
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
	STDMETHOD(get_Item)(VARIANT Index, ColorScaleCriterion * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xaa, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &Index, RHS);
		return result;
	}

	// IColorScaleCriteria 属性
public:

};
