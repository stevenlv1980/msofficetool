// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CXmlMap0 包装类

class CXmlMap0 : public COleDispatchDriver
{
public:
	CXmlMap0(){} // 调用 COleDispatchDriver 默认构造函数
	CXmlMap0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CXmlMap0(const CXmlMap0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// XmlMap 方法
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
	CString get__Default()
	{
		CString result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
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
	BOOL get_IsExportable()
	{
		BOOL result;
		InvokeHelper(0x91e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_ShowImportExportValidationErrors()
	{
		BOOL result;
		InvokeHelper(0x91f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ShowImportExportValidationErrors(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x91f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_SaveDataSourceDefinition()
	{
		BOOL result;
		InvokeHelper(0x920, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_SaveDataSourceDefinition(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x920, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_AdjustColumnWidth()
	{
		BOOL result;
		InvokeHelper(0x74c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_AdjustColumnWidth(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x74c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_PreserveColumnFilter()
	{
		BOOL result;
		InvokeHelper(0x921, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_PreserveColumnFilter(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x921, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_PreserveNumberFormatting()
	{
		BOOL result;
		InvokeHelper(0x922, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_PreserveNumberFormatting(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x922, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_AppendOnImport()
	{
		BOOL result;
		InvokeHelper(0x923, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_AppendOnImport(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x923, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_RootElementName()
	{
		CString result;
		InvokeHelper(0x924, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_RootElementNamespace()
	{
		LPDISPATCH result;
		InvokeHelper(0x925, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Schemas()
	{
		LPDISPATCH result;
		InvokeHelper(0x926, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_DataBinding()
	{
		LPDISPATCH result;
		InvokeHelper(0x927, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void Delete()
	{
		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	long Import(LPCTSTR Url, VARIANT Overwrite)
	{
		long result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT ;
		InvokeHelper(0x395, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Url, &Overwrite);
		return result;
	}
	long ImportXml(LPCTSTR XmlData, VARIANT Overwrite)
	{
		long result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT ;
		InvokeHelper(0x928, DISPATCH_METHOD, VT_I4, (void*)&result, parms, XmlData, &Overwrite);
		return result;
	}
	long Export(LPCTSTR Url, VARIANT Overwrite)
	{
		long result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT ;
		InvokeHelper(0x586, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Url, &Overwrite);
		return result;
	}
	long ExportXml(BSTR * Data)
	{
		long result;
		static BYTE parms[] = VTS_PBSTR ;
		InvokeHelper(0x92a, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Data);
		return result;
	}
	LPDISPATCH get_WorkbookConnection()
	{
		LPDISPATCH result;
		InvokeHelper(0x9f0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// XmlMap 属性
public:

};
