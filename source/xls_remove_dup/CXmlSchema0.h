// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CXmlSchema0 包装类

class CXmlSchema0 : public COleDispatchDriver
{
public:
	CXmlSchema0(){} // 调用 COleDispatchDriver 默认构造函数
	CXmlSchema0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CXmlSchema0(const CXmlSchema0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// XmlSchema 方法
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
	LPDISPATCH get_Namespace()
	{
		LPDISPATCH result;
		InvokeHelper(0x91c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_XML()
	{
		CString result;
		InvokeHelper(0x91d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}

	// XmlSchema 属性
public:

};
