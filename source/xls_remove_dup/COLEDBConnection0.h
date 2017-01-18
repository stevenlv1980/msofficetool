// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// COLEDBConnection0 包装类

class COLEDBConnection0 : public COleDispatchDriver
{
public:
	COLEDBConnection0(){} // 调用 COleDispatchDriver 默认构造函数
	COLEDBConnection0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COLEDBConnection0(const COLEDBConnection0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// OLEDBConnection 方法
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
	LPDISPATCH get_ADOConnection()
	{
		LPDISPATCH result;
		InvokeHelper(0x81a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_BackgroundQuery()
	{
		BOOL result;
		InvokeHelper(0x593, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_BackgroundQuery(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x593, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void CancelRefresh()
	{
		InvokeHelper(0x635, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	VARIANT get_CommandText()
	{
		VARIANT result;
		InvokeHelper(0x725, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_CommandText(VARIANT newValue)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x725, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
	}
	long get_CommandType()
	{
		long result;
		InvokeHelper(0x726, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_CommandType(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x726, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	VARIANT get_Connection()
	{
		VARIANT result;
		InvokeHelper(0x598, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_Connection(VARIANT newValue)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x598, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
	}
	BOOL get_EnableRefresh()
	{
		BOOL result;
		InvokeHelper(0x5c5, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_EnableRefresh(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x5c5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	VARIANT get_LocalConnection()
	{
		VARIANT result;
		InvokeHelper(0x72b, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_LocalConnection(VARIANT newValue)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x72b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
	}
	BOOL get_MaintainConnection()
	{
		BOOL result;
		InvokeHelper(0x728, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_MaintainConnection(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x728, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void MakeConnection()
	{
		InvokeHelper(0x81c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Refresh()
	{
		InvokeHelper(0x589, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	DATE get_RefreshDate()
	{
		DATE result;
		InvokeHelper(0x2b8, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, NULL);
		return result;
	}
	BOOL get_Refreshing()
	{
		BOOL result;
		InvokeHelper(0x633, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_RefreshOnFileOpen()
	{
		BOOL result;
		InvokeHelper(0x5c7, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_RefreshOnFileOpen(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x5c7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_RefreshPeriod()
	{
		long result;
		InvokeHelper(0x729, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_RefreshPeriod(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x729, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_RobustConnect()
	{
		long result;
		InvokeHelper(0x821, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_RobustConnect(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x821, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void SaveAsODC(LPCTSTR ODCFileName, VARIANT Description, VARIANT Keywords)
	{
		static BYTE parms[] = VTS_BSTR VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x822, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ODCFileName, &Description, &Keywords);
	}
	BOOL get_SavePassword()
	{
		BOOL result;
		InvokeHelper(0x5c9, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_SavePassword(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x5c9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_SourceConnectionFile()
	{
		CString result;
		InvokeHelper(0x81f, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_SourceConnectionFile(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x81f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_SourceDataFile()
	{
		CString result;
		InvokeHelper(0x820, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_SourceDataFile(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x820, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_OLAP()
	{
		BOOL result;
		InvokeHelper(0x81d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_UseLocalConnection()
	{
		BOOL result;
		InvokeHelper(0x72d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_UseLocalConnection(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x72d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_MaxDrillthroughRecords()
	{
		long result;
		InvokeHelper(0xa8f, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_MaxDrillthroughRecords(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xa8f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_IsConnected()
	{
		BOOL result;
		InvokeHelper(0x81b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	long get_ServerCredentialsMethod()
	{
		long result;
		InvokeHelper(0xa90, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ServerCredentialsMethod(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xa90, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_ServerSSOApplicationID()
	{
		CString result;
		InvokeHelper(0xa91, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_ServerSSOApplicationID(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0xa91, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_AlwaysUseConnectionFile()
	{
		BOOL result;
		InvokeHelper(0xa92, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_AlwaysUseConnectionFile(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa92, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_ServerFillColor()
	{
		BOOL result;
		InvokeHelper(0xa93, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ServerFillColor(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa93, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_ServerFontStyle()
	{
		BOOL result;
		InvokeHelper(0xa94, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ServerFontStyle(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa94, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_ServerNumberFormat()
	{
		BOOL result;
		InvokeHelper(0xa95, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ServerNumberFormat(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa95, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_ServerTextColor()
	{
		BOOL result;
		InvokeHelper(0xa96, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ServerTextColor(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa96, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_RetrieveInOfficeUILang()
	{
		BOOL result;
		InvokeHelper(0xa97, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_RetrieveInOfficeUILang(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa97, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// OLEDBConnection 属性
public:

};
