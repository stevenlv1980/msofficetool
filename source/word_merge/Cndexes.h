// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// Cndexes 包装类

class Cndexes : public COleDispatchDriver
{
public:
	Cndexes(){} // 调用 COleDispatchDriver 默认构造函数
	Cndexes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Cndexes(const Cndexes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Indexes 方法
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
	LPDISPATCH AddOld(LPDISPATCH Range, VARIANT * HeadingSeparator, VARIANT * RightAlignPageNumbers, VARIANT * Type, VARIANT * NumberOfColumns, VARIANT * AccentedLetters)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x64, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, HeadingSeparator, RightAlignPageNumbers, Type, NumberOfColumns, AccentedLetters);
		return result;
	}
	LPDISPATCH MarkEntry(LPDISPATCH Range, VARIANT * Entry, VARIANT * EntryAutoText, VARIANT * CrossReference, VARIANT * CrossReferenceAutoText, VARIANT * BookmarkName, VARIANT * Bold, VARIANT * Italic, VARIANT * Reading)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, Entry, EntryAutoText, CrossReference, CrossReferenceAutoText, BookmarkName, Bold, Italic, Reading);
		return result;
	}
	void MarkAllEntries(LPDISPATCH Range, VARIANT * Entry, VARIANT * EntryAutoText, VARIANT * CrossReference, VARIANT * CrossReferenceAutoText, VARIANT * BookmarkName, VARIANT * Bold, VARIANT * Italic)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Range, Entry, EntryAutoText, CrossReference, CrossReferenceAutoText, BookmarkName, Bold, Italic);
	}
	void AutoMarkEntries(LPCTSTR ConcordanceFileName)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x67, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ConcordanceFileName);
	}
	LPDISPATCH Add(LPDISPATCH Range, VARIANT * HeadingSeparator, VARIANT * RightAlignPageNumbers, VARIANT * Type, VARIANT * NumberOfColumns, VARIANT * AccentedLetters, VARIANT * SortBy, VARIANT * IndexLanguage)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x68, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, HeadingSeparator, RightAlignPageNumbers, Type, NumberOfColumns, AccentedLetters, SortBy, IndexLanguage);
		return result;
	}

	// Indexes 属性
public:

};
