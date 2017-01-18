// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CListObjects 包装类

class CListObjects : public COleDispatchDriver
{
public:
	CListObjects(){} // 调用 COleDispatchDriver 默认构造函数
	CListObjects(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CListObjects(const CListObjects& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IListObjects 方法
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
	STDMETHOD(_Add)(XlListObjectSourceType SourceType, VARIANT Source, VARIANT LinkSource, XlYesNoGuess XlListObjectHasHeaders, VARIANT Destination, ListObject * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x825, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, SourceType, &Source, &LinkSource, XlListObjectHasHeaders, &Destination, RHS);
		return result;
	}
	STDMETHOD(get__Default)(VARIANT Index, ListObject * * RHS)
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
	STDMETHOD(get_Item)(VARIANT Index, ListObject * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xaa, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &Index, RHS);
		return result;
	}
	STDMETHOD(get_Count)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Add)(XlListObjectSourceType SourceType, VARIANT Source, VARIANT LinkSource, XlYesNoGuess XlListObjectHasHeaders, VARIANT Destination, VARIANT TableStyleName, ListObject * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xb5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, SourceType, &Source, &LinkSource, XlListObjectHasHeaders, &Destination, &TableStyleName, RHS);
		return result;
	}

	// IListObjects 属性
public:

};
