// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CSmartTagType 包装类

class CSmartTagType : public COleDispatchDriver
{
public:
	CSmartTagType(){} // 调用 COleDispatchDriver 默认构造函数
	CSmartTagType(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSmartTagType(const CSmartTagType& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// SmartTagType 方法
public:
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
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
	LPDISPATCH get_SmartTagActions()
	{
		LPDISPATCH result;
		InvokeHelper(0x3eb, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_SmartTagRecognizers()
	{
		LPDISPATCH result;
		InvokeHelper(0x3ec, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_FriendlyName()
	{
		CString result;
		InvokeHelper(0x3ed, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}

	// SmartTagType 属性
public:

};
