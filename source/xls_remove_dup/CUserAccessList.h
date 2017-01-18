// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CUserAccessList 包装类

class CUserAccessList : public COleDispatchDriver
{
public:
	CUserAccessList(){} // 调用 COleDispatchDriver 默认构造函数
	CUserAccessList(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CUserAccessList(const CUserAccessList& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IUserAccessList 方法
public:
	STDMETHOD(get_Count)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Item)(VARIANT Index, UserAccess * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xaa, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &Index, RHS);
		return result;
	}
	STDMETHOD(Add)(LPCTSTR Name, BOOL AllowEdit, UserAccess * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_BOOL VTS_PDISPATCH ;
		InvokeHelper(0xb5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Name, AllowEdit, RHS);
		return result;
	}
	STDMETHOD(DeleteAll)()
	{
		HRESULT result;
		InvokeHelper(0x8bf, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(get__Default)(VARIANT Index, UserAccess * * RHS)
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

	// IUserAccessList 属性
public:

};
