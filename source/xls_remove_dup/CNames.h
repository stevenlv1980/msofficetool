// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CNames 包装类

class CNames : public COleDispatchDriver
{
public:
	CNames(){} // 调用 COleDispatchDriver 默认构造函数
	CNames(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CNames(const CNames& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// INames 方法
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
	STDMETHOD(Add)(VARIANT Name, VARIANT RefersTo, VARIANT Visible, VARIANT MacroType, VARIANT ShortcutKey, VARIANT Category, VARIANT NameLocal, VARIANT RefersToLocal, VARIANT CategoryLocal, VARIANT RefersToR1C1, VARIANT RefersToR1C1Local, Name * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xb5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Name, &RefersTo, &Visible, &MacroType, &ShortcutKey, &Category, &NameLocal, &RefersToLocal, &CategoryLocal, &RefersToR1C1, &RefersToR1C1Local, RHS);
		return result;
	}
	STDMETHOD(Item)(VARIANT Index, VARIANT IndexLocal, VARIANT RefersTo, long lcid, Name * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0xaa, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, &IndexLocal, &RefersTo, lcid, RHS);
		return result;
	}
	STDMETHOD(_Default)(VARIANT Index, VARIANT IndexLocal, VARIANT RefersTo, long lcid, Name * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, &IndexLocal, &RefersTo, lcid, RHS);
		return result;
	}
	STDMETHOD(get_Count)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get__NewEnum)(LPUNKNOWN * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PUNKNOWN ;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}

	// INames 属性
public:

};
