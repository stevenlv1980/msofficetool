// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CListObject0 包装类

class CListObject0 : public COleDispatchDriver
{
public:
	CListObject0(){} // 调用 COleDispatchDriver 默认构造函数
	CListObject0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CListObject0(const CListObject0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ListObject 方法
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
	void Delete()
	{
		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	CString Publish(VARIANT Target, BOOL LinkSource)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_BOOL ;
		InvokeHelper(0x767, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Target, LinkSource);
		return result;
	}
	void Refresh()
	{
		InvokeHelper(0x589, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Unlink()
	{
		InvokeHelper(0x904, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Unlist()
	{
		InvokeHelper(0x905, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void UpdateChanges(long iConflictType)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x906, DISPATCH_METHOD, VT_EMPTY, NULL, parms, iConflictType);
	}
	void Resize(LPDISPATCH Range)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x100, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Range);
	}
	CString get__Default()
	{
		CString result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	BOOL get_Active()
	{
		BOOL result;
		InvokeHelper(0x908, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_DataBodyRange()
	{
		LPDISPATCH result;
		InvokeHelper(0x2c1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_DisplayRightToLeft()
	{
		BOOL result;
		InvokeHelper(0x6ee, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_HeaderRowRange()
	{
		LPDISPATCH result;
		InvokeHelper(0x909, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_InsertRowRange()
	{
		LPDISPATCH result;
		InvokeHelper(0x90a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ListColumns()
	{
		LPDISPATCH result;
		InvokeHelper(0x90b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ListRows()
	{
		LPDISPATCH result;
		InvokeHelper(0x90c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Name(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x6e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_QueryTable()
	{
		LPDISPATCH result;
		InvokeHelper(0x56a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Range()
	{
		LPDISPATCH result;
		InvokeHelper(0xc5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_ShowAutoFilter()
	{
		BOOL result;
		InvokeHelper(0x90d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ShowAutoFilter(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x90d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_ShowTotals()
	{
		BOOL result;
		InvokeHelper(0x90e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ShowTotals(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x90e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_SourceType()
	{
		long result;
		InvokeHelper(0x2ad, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_TotalsRowRange()
	{
		LPDISPATCH result;
		InvokeHelper(0x90f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_SharePointURL()
	{
		CString result;
		InvokeHelper(0x910, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_XmlMap()
	{
		LPDISPATCH result;
		InvokeHelper(0x8cd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_DisplayName()
	{
		CString result;
		InvokeHelper(0xa75, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_DisplayName(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0xa75, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_ShowHeaders()
	{
		BOOL result;
		InvokeHelper(0xa76, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ShowHeaders(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa76, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_AutoFilter()
	{
		LPDISPATCH result;
		InvokeHelper(0x319, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	VARIANT get_TableStyle()
	{
		VARIANT result;
		InvokeHelper(0x5e0, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_TableStyle(VARIANT newValue)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x5e0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
	}
	BOOL get_ShowTableStyleFirstColumn()
	{
		BOOL result;
		InvokeHelper(0xa77, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ShowTableStyleFirstColumn(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa77, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_ShowTableStyleLastColumn()
	{
		BOOL result;
		InvokeHelper(0xa03, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ShowTableStyleLastColumn(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa03, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_ShowTableStyleRowStripes()
	{
		BOOL result;
		InvokeHelper(0xa04, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ShowTableStyleRowStripes(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa04, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_ShowTableStyleColumnStripes()
	{
		BOOL result;
		InvokeHelper(0xa05, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ShowTableStyleColumnStripes(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa05, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Sort()
	{
		LPDISPATCH result;
		InvokeHelper(0x370, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_Comment()
	{
		CString result;
		InvokeHelper(0x38e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Comment(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x38e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void ExportToVisio()
	{
		InvokeHelper(0xa78, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// ListObject 属性
public:

};
