// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CAllowEditRange 包装类

class CAllowEditRange : public COleDispatchDriver
{
public:
	CAllowEditRange(){} // 调用 COleDispatchDriver 默认构造函数
	CAllowEditRange(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAllowEditRange(const CAllowEditRange& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// AllowEditRange 方法
public:
	CString get_Title()
	{
		CString result;
		InvokeHelper(0xc7, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Title(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0xc7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Range()
	{
		LPDISPATCH result;
		InvokeHelper(0xc5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void putref_Range(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xc5, DISPATCH_PROPERTYPUTREF, VT_EMPTY, NULL, parms, newValue);
	}
	void ChangePassword(LPCTSTR Password)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x8bd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Password);
	}
	void Delete()
	{
		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Unprotect(VARIANT Password)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x11d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Password);
	}
	LPDISPATCH get_Users()
	{
		LPDISPATCH result;
		InvokeHelper(0x8be, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// AllowEditRange 属性
public:

};
