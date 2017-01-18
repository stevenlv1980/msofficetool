// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CXSLTransform 包装类

class CXSLTransform : public COleDispatchDriver
{
public:
	CXSLTransform(){} // 调用 COleDispatchDriver 默认构造函数
	CXSLTransform(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CXSLTransform(const CXSLTransform& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// XSLTransform 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_Alias(BOOL AllUsers)
	{
		CString result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms, AllUsers);
		return result;
	}
	void put_Alias(BOOL AllUsers, LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BOOL VTS_BSTR ;
		InvokeHelper(0x2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, AllUsers, newValue);
	}
	CString get_Location(BOOL AllUsers)
	{
		CString result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms, AllUsers);
		return result;
	}
	void put_Location(BOOL AllUsers, LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BOOL VTS_BSTR ;
		InvokeHelper(0x3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, AllUsers, newValue);
	}
	void Delete()
	{
		InvokeHelper(0x65, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// XSLTransform 属性
public:

};
