// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CEditor 包装类

class CEditor : public COleDispatchDriver
{
public:
	CEditor(){} // 调用 COleDispatchDriver 默认构造函数
	CEditor(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CEditor(const CEditor& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Editor 方法
public:
	CString get_ID()
	{
		CString result;
		InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x65, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Range()
	{
		LPDISPATCH result;
		InvokeHelper(0x66, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_NextRange()
	{
		LPDISPATCH result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
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
	void Delete()
	{
		InvokeHelper(0x1f4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void DeleteAll()
	{
		InvokeHelper(0x1f5, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SelectAll()
	{
		InvokeHelper(0x1f6, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// Editor 属性
public:

};
