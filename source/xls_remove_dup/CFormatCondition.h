// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CFormatCondition 包装类

class CFormatCondition : public COleDispatchDriver
{
public:
	CFormatCondition(){} // 调用 COleDispatchDriver 默认构造函数
	CFormatCondition(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CFormatCondition(const CFormatCondition& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// FormatCondition 方法
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
	void _Modify(long Type, VARIANT Operator, VARIANT Formula1, VARIANT Formula2)
	{
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0xa3f, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type, &Operator, &Formula1, &Formula2);
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Operator()
	{
		long result;
		InvokeHelper(0x31d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	CString get_Formula1()
	{
		CString result;
		InvokeHelper(0x62b, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get_Formula2()
	{
		CString result;
		InvokeHelper(0x62c, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Interior()
	{
		LPDISPATCH result;
		InvokeHelper(0x81, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Borders()
	{
		LPDISPATCH result;
		InvokeHelper(0x1b3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Font()
	{
		LPDISPATCH result;
		InvokeHelper(0x92, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void Delete()
	{
		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Modify(long Type, VARIANT Operator, VARIANT Formula1, VARIANT Formula2, VARIANT String, VARIANT Operator2)
	{
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x62d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type, &Operator, &Formula1, &Formula2, &String, &Operator2);
	}
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
	long get_TextOperator()
	{
		long result;
		InvokeHelper(0xa35, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_TextOperator(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xa35, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_DateOperator()
	{
		long result;
		InvokeHelper(0xa36, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_DateOperator(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xa36, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	VARIANT get_NumberFormat()
	{
		VARIANT result;
		InvokeHelper(0xc1, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_NumberFormat(VARIANT newValue)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xc1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
	}
	long get_Priority()
	{
		long result;
		InvokeHelper(0x3d9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Priority(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x3d9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_StopIfTrue()
	{
		BOOL result;
		InvokeHelper(0xa41, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_StopIfTrue(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa41, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_AppliesTo()
	{
		LPDISPATCH result;
		InvokeHelper(0xa42, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void ModifyAppliesToRange(LPDISPATCH Range)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xa43, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Range);
	}
	void SetFirstPriority()
	{
		InvokeHelper(0xa45, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SetLastPriority()
	{
		InvokeHelper(0xa46, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	BOOL get_PTCondition()
	{
		BOOL result;
		InvokeHelper(0xa47, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	long get_ScopeType()
	{
		long result;
		InvokeHelper(0xa37, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ScopeType(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xa37, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// FormatCondition 属性
public:

};
