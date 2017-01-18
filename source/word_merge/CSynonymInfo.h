// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CSynonymInfo 包装类

class CSynonymInfo : public COleDispatchDriver
{
public:
	CSynonymInfo(){} // 调用 COleDispatchDriver 默认构造函数
	CSynonymInfo(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSynonymInfo(const CSynonymInfo& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// SynonymInfo 方法
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
	CString get_Word()
	{
		CString result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	BOOL get_Found()
	{
		BOOL result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	long get_MeaningCount()
	{
		long result;
		InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	VARIANT get_MeaningList()
	{
		VARIANT result;
		InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	VARIANT get_PartOfSpeechList()
	{
		VARIANT result;
		InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	VARIANT get_SynonymList(VARIANT * Meaning)
	{
		VARIANT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x7, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms, Meaning);
		return result;
	}
	VARIANT get_AntonymList()
	{
		VARIANT result;
		InvokeHelper(0x8, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	VARIANT get_RelatedExpressionList()
	{
		VARIANT result;
		InvokeHelper(0x9, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	VARIANT get_RelatedWordList()
	{
		VARIANT result;
		InvokeHelper(0xa, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}

	// SynonymInfo 属性
public:

};
