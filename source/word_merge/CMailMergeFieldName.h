// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CMailMergeFieldName 包装类

class CMailMergeFieldName : public COleDispatchDriver
{
public:
	CMailMergeFieldName(){} // 调用 COleDispatchDriver 默认构造函数
	CMailMergeFieldName(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CMailMergeFieldName(const CMailMergeFieldName& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// MailMergeFieldName 方法
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
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	long get_Index()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}

	// MailMergeFieldName 属性
public:

};
