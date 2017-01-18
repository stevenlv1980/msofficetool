// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CPublishObject 包装类

class CPublishObject : public COleDispatchDriver
{
public:
	CPublishObject(){} // 调用 COleDispatchDriver 默认构造函数
	CPublishObject(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CPublishObject(const CPublishObject& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// PublishObject 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void Delete()
	{
		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Publish(VARIANT Create)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x767, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Create);
	}
	CString get_DivID()
	{
		CString result;
		InvokeHelper(0x766, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get_Sheet()
	{
		CString result;
		InvokeHelper(0x2ef, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	long get_SourceType()
	{
		long result;
		InvokeHelper(0x2ad, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	CString get_Source()
	{
		CString result;
		InvokeHelper(0xde, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	long get_HtmlType()
	{
		long result;
		InvokeHelper(0x765, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_HtmlType(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x765, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
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
	CString get_Filename()
	{
		CString result;
		InvokeHelper(0x587, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Filename(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x587, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_AutoRepublish()
	{
		BOOL result;
		InvokeHelper(0x882, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_AutoRepublish(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x882, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// PublishObject 属性
public:

};
