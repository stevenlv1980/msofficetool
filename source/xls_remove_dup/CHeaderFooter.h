// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CHeaderFooter 包装类

class CHeaderFooter : public COleDispatchDriver
{
public:
	CHeaderFooter(){} // 调用 COleDispatchDriver 默认构造函数
	CHeaderFooter(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CHeaderFooter(const CHeaderFooter& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// HeaderFooter 方法
public:
	CString get_Text()
	{
		CString result;
		InvokeHelper(0x8a, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Text(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x8a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Picture()
	{
		LPDISPATCH result;
		InvokeHelper(0x1df, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// HeaderFooter 属性
public:

};
