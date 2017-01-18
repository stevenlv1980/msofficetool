// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CAutoCorrect 包装类

class CAutoCorrect : public COleDispatchDriver
{
public:
	CAutoCorrect(){} // 调用 COleDispatchDriver 默认构造函数
	CAutoCorrect(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAutoCorrect(const CAutoCorrect& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// AutoCorrect 方法
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
	VARIANT AddReplacement(LPCTSTR What, LPCTSTR Replacement)
	{
		VARIANT result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR ;
		InvokeHelper(0x47a, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, What, Replacement);
		return result;
	}
	BOOL get_CapitalizeNamesOfDays()
	{
		BOOL result;
		InvokeHelper(0x47e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_CapitalizeNamesOfDays(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x47e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	VARIANT DeleteReplacement(LPCTSTR What)
	{
		VARIANT result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x47b, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, What);
		return result;
	}
	VARIANT get_ReplacementList(VARIANT Index)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x47f, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms, &Index);
		return result;
	}
	void put_ReplacementList(VARIANT Index, VARIANT newValue)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x47f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &Index, &newValue);
	}
	BOOL get_ReplaceText()
	{
		BOOL result;
		InvokeHelper(0x47c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ReplaceText(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x47c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_TwoInitialCapitals()
	{
		BOOL result;
		InvokeHelper(0x47d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_TwoInitialCapitals(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x47d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_CorrectSentenceCap()
	{
		BOOL result;
		InvokeHelper(0x653, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_CorrectSentenceCap(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x653, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_CorrectCapsLock()
	{
		BOOL result;
		InvokeHelper(0x654, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_CorrectCapsLock(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x654, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_DisplayAutoCorrectOptions()
	{
		BOOL result;
		InvokeHelper(0x786, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DisplayAutoCorrectOptions(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x786, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_AutoExpandListRange()
	{
		BOOL result;
		InvokeHelper(0x8f6, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_AutoExpandListRange(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x8f6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_AutoFillFormulasInLists()
	{
		BOOL result;
		InvokeHelper(0xa52, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_AutoFillFormulasInLists(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa52, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// AutoCorrect 属性
public:

};
