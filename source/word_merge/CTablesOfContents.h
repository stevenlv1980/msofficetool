// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CTablesOfContents 包装类

class CTablesOfContents : public COleDispatchDriver
{
public:
	CTablesOfContents(){} // 调用 COleDispatchDriver 默认构造函数
	CTablesOfContents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTablesOfContents(const CTablesOfContents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// TablesOfContents 方法
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
	LPDISPATCH AddOld(LPDISPATCH Range, VARIANT * UseHeadingStyles, VARIANT * UpperHeadingLevel, VARIANT * LowerHeadingLevel, VARIANT * UseFields, VARIANT * TableID, VARIANT * RightAlignPageNumbers, VARIANT * IncludePageNumbers, VARIANT * AddedStyles)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x64, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, UseHeadingStyles, UpperHeadingLevel, LowerHeadingLevel, UseFields, TableID, RightAlignPageNumbers, IncludePageNumbers, AddedStyles);
		return result;
	}
	LPDISPATCH MarkEntry(LPDISPATCH Range, VARIANT * Entry, VARIANT * EntryAutoText, VARIANT * TableID, VARIANT * Level)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, Entry, EntryAutoText, TableID, Level);
		return result;
	}
	LPDISPATCH Add2000(LPDISPATCH Range, VARIANT * UseHeadingStyles, VARIANT * UpperHeadingLevel, VARIANT * LowerHeadingLevel, VARIANT * UseFields, VARIANT * TableID, VARIANT * RightAlignPageNumbers, VARIANT * IncludePageNumbers, VARIANT * AddedStyles, VARIANT * UseHyperlinks, VARIANT * HidePageNumbersInWeb)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x66, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, UseHeadingStyles, UpperHeadingLevel, LowerHeadingLevel, UseFields, TableID, RightAlignPageNumbers, IncludePageNumbers, AddedStyles, UseHyperlinks, HidePageNumbersInWeb);
		return result;
	}
	LPDISPATCH Add(LPDISPATCH Range, VARIANT * UseHeadingStyles, VARIANT * UpperHeadingLevel, VARIANT * LowerHeadingLevel, VARIANT * UseFields, VARIANT * TableID, VARIANT * RightAlignPageNumbers, VARIANT * IncludePageNumbers, VARIANT * AddedStyles, VARIANT * UseHyperlinks, VARIANT * HidePageNumbersInWeb, VARIANT * UseOutlineLevels)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x67, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, UseHeadingStyles, UpperHeadingLevel, LowerHeadingLevel, UseFields, TableID, RightAlignPageNumbers, IncludePageNumbers, AddedStyles, UseHyperlinks, HidePageNumbersInWeb, UseOutlineLevels);
		return result;
	}

	// TablesOfContents 属性
public:

};
