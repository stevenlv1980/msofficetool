// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CBibliography 包装类

class CBibliography : public COleDispatchDriver
{
public:
	CBibliography(){} // 调用 COleDispatchDriver 默认构造函数
	CBibliography(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CBibliography(const CBibliography& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Bibliography 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x65, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x66, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Sources()
	{
		LPDISPATCH result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_BibliographyStyle()
	{
		CString result;
		InvokeHelper(0x69, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_BibliographyStyle(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x69, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString GenerateUniqueTag()
	{
		CString result;
		InvokeHelper(0x68, DISPATCH_METHOD, VT_BSTR, (void*)&result, NULL);
		return result;
	}

	// Bibliography 属性
public:

};
