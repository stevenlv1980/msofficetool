// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CHeaderFooter0 包装类

class CHeaderFooter0 : public COleDispatchDriver
{
public:
	CHeaderFooter0(){} // 调用 COleDispatchDriver 默认构造函数
	CHeaderFooter0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CHeaderFooter0(const CHeaderFooter0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IHeaderFooter 方法
public:
	STDMETHOD(get_Text)(BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBSTR ;
		InvokeHelper(0x8a, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Text)(LPCTSTR RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x8a, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Picture)(Graphic * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x1df, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}

	// IHeaderFooter 属性
public:

};
