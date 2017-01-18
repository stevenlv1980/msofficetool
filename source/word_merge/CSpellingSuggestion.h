// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CSpellingSuggestion 包装类

class CSpellingSuggestion : public COleDispatchDriver
{
public:
	CSpellingSuggestion(){} // 调用 COleDispatchDriver 默认构造函数
	CSpellingSuggestion(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSpellingSuggestion(const CSpellingSuggestion& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// SpellingSuggestion 方法
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

	// SpellingSuggestion 属性
public:

};
