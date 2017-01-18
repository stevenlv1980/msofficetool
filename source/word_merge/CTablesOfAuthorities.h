// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CTablesOfAuthorities 包装类

class CTablesOfAuthorities : public COleDispatchDriver
{
public:
	CTablesOfAuthorities(){} // 调用 COleDispatchDriver 默认构造函数
	CTablesOfAuthorities(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTablesOfAuthorities(const CTablesOfAuthorities& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// TablesOfAuthorities 方法
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
	long get_Format()
	{
		long result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Format(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH Item(long Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH Add(LPDISPATCH Range, VARIANT * Category, VARIANT * Bookmark, VARIANT * Passim, VARIANT * KeepEntryFormatting, VARIANT * Separator, VARIANT * IncludeSequenceName, VARIANT * EntrySeparator, VARIANT * PageRangeSeparator, VARIANT * IncludeCategoryHeader, VARIANT * PageNumberSeparator)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x64, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, Category, Bookmark, Passim, KeepEntryFormatting, Separator, IncludeSequenceName, EntrySeparator, PageRangeSeparator, IncludeCategoryHeader, PageNumberSeparator);
		return result;
	}
	void NextCitation(LPCTSTR ShortCitation)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x67, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ShortCitation);
	}
	LPDISPATCH MarkCitation(LPDISPATCH Range, LPCTSTR ShortCitation, VARIANT * LongCitation, VARIANT * LongCitationAutoText, VARIANT * Category)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, ShortCitation, LongCitation, LongCitationAutoText, Category);
		return result;
	}
	void MarkAllCitations(LPCTSTR ShortCitation, VARIANT * LongCitation, VARIANT * LongCitationAutoText, VARIANT * Category)
	{
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ShortCitation, LongCitation, LongCitationAutoText, Category);
	}

	// TablesOfAuthorities 属性
public:

};
