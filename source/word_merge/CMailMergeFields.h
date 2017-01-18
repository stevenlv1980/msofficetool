// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CMailMergeFields 包装类

class CMailMergeFields : public COleDispatchDriver
{
public:
	CMailMergeFields(){} // 调用 COleDispatchDriver 默认构造函数
	CMailMergeFields(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CMailMergeFields(const CMailMergeFields& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// MailMergeFields 方法
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
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(long Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH Add(LPDISPATCH Range, LPCTSTR Name)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR ;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, Name);
		return result;
	}
	LPDISPATCH AddAsk(LPDISPATCH Range, LPCTSTR Name, VARIANT * Prompt, VARIANT * DefaultAskText, VARIANT * AskOnce)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x66, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, Name, Prompt, DefaultAskText, AskOnce);
		return result;
	}
	LPDISPATCH AddFillIn(LPDISPATCH Range, VARIANT * Prompt, VARIANT * DefaultFillInText, VARIANT * AskOnce)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x67, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, Prompt, DefaultFillInText, AskOnce);
		return result;
	}
	LPDISPATCH AddIf(LPDISPATCH Range, LPCTSTR MergeField, long Comparison, VARIANT * CompareTo, VARIANT * TrueAutoText, VARIANT * TrueText, VARIANT * FalseAutoText, VARIANT * FalseText)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x68, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, MergeField, Comparison, CompareTo, TrueAutoText, TrueText, FalseAutoText, FalseText);
		return result;
	}
	LPDISPATCH AddMergeRec(LPDISPATCH Range)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x69, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range);
		return result;
	}
	LPDISPATCH AddMergeSeq(LPDISPATCH Range)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x6a, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range);
		return result;
	}
	LPDISPATCH AddNext(LPDISPATCH Range)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x6b, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range);
		return result;
	}
	LPDISPATCH AddNextIf(LPDISPATCH Range, LPCTSTR MergeField, long Comparison, VARIANT * CompareTo)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x6c, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, MergeField, Comparison, CompareTo);
		return result;
	}
	LPDISPATCH AddSet(LPDISPATCH Range, LPCTSTR Name, VARIANT * ValueText, VARIANT * ValueAutoText)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x6d, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, Name, ValueText, ValueAutoText);
		return result;
	}
	LPDISPATCH AddSkipIf(LPDISPATCH Range, LPCTSTR MergeField, long Comparison, VARIANT * CompareTo)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x6e, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, MergeField, Comparison, CompareTo);
		return result;
	}

	// MailMergeFields 属性
public:

};
