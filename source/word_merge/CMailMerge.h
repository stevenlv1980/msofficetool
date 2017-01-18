// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CMailMerge 包装类

class CMailMerge : public COleDispatchDriver
{
public:
	CMailMerge(){} // 调用 COleDispatchDriver 默认构造函数
	CMailMerge(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CMailMerge(const CMailMerge& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// MailMerge 方法
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
	long get_MainDocumentType()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_MainDocumentType(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_State()
	{
		long result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Destination()
	{
		long result;
		InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Destination(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_DataSource()
	{
		LPDISPATCH result;
		InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Fields()
	{
		LPDISPATCH result;
		InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_ViewMailMergeFieldCodes()
	{
		long result;
		InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ViewMailMergeFieldCodes(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_SuppressBlankLines()
	{
		BOOL result;
		InvokeHelper(0x7, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_SuppressBlankLines(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_MailAsAttachment()
	{
		BOOL result;
		InvokeHelper(0x8, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_MailAsAttachment(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_MailAddressFieldName()
	{
		CString result;
		InvokeHelper(0x9, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_MailAddressFieldName(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_MailSubject()
	{
		CString result;
		InvokeHelper(0xa, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_MailSubject(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0xa, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void CreateDataSource(VARIANT * Name, VARIANT * PasswordDocument, VARIANT * WritePasswordDocument, VARIANT * HeaderRecord, VARIANT * MSQuery, VARIANT * SQLStatement, VARIANT * SQLStatement1, VARIANT * Connection, VARIANT * LinkToSource)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name, PasswordDocument, WritePasswordDocument, HeaderRecord, MSQuery, SQLStatement, SQLStatement1, Connection, LinkToSource);
	}
	void CreateHeaderSource(LPCTSTR Name, VARIANT * PasswordDocument, VARIANT * WritePasswordDocument, VARIANT * HeaderRecord)
	{
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name, PasswordDocument, WritePasswordDocument, HeaderRecord);
	}
	void OpenDataSource2000(LPCTSTR Name, VARIANT * Format, VARIANT * ConfirmConversions, VARIANT * ReadOnly, VARIANT * LinkToSource, VARIANT * AddToRecentFiles, VARIANT * PasswordDocument, VARIANT * PasswordTemplate, VARIANT * Revert, VARIANT * WritePasswordDocument, VARIANT * WritePasswordTemplate, VARIANT * Connection, VARIANT * SQLStatement, VARIANT * SQLStatement1)
	{
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x67, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name, Format, ConfirmConversions, ReadOnly, LinkToSource, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Connection, SQLStatement, SQLStatement1);
	}
	void OpenHeaderSource2000(LPCTSTR Name, VARIANT * Format, VARIANT * ConfirmConversions, VARIANT * ReadOnly, VARIANT * AddToRecentFiles, VARIANT * PasswordDocument, VARIANT * PasswordTemplate, VARIANT * Revert, VARIANT * WritePasswordDocument, VARIANT * WritePasswordTemplate)
	{
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x68, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name, Format, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate);
	}
	void Execute(VARIANT * Pause)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x69, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Pause);
	}
	void Check()
	{
		InvokeHelper(0x6a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void EditDataSource()
	{
		InvokeHelper(0x6b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void EditHeaderSource()
	{
		InvokeHelper(0x6c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void EditMainDocument()
	{
		InvokeHelper(0x6d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void UseAddressBook(LPCTSTR Type)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x6f, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type);
	}
	BOOL get_HighlightMergeFields()
	{
		BOOL result;
		InvokeHelper(0xb, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_HighlightMergeFields(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_MailFormat()
	{
		long result;
		InvokeHelper(0xc, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_MailFormat(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xc, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_ShowSendToCustom()
	{
		CString result;
		InvokeHelper(0xd, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_ShowSendToCustom(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0xd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_WizardState()
	{
		long result;
		InvokeHelper(0xe, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_WizardState(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xe, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void OpenDataSource(LPCTSTR Name, VARIANT * Format, VARIANT * ConfirmConversions, VARIANT * ReadOnly, VARIANT * LinkToSource, VARIANT * AddToRecentFiles, VARIANT * PasswordDocument, VARIANT * PasswordTemplate, VARIANT * Revert, VARIANT * WritePasswordDocument, VARIANT * WritePasswordTemplate, VARIANT * Connection, VARIANT * SQLStatement, VARIANT * SQLStatement1, VARIANT * OpenExclusive, VARIANT * SubType)
	{
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x70, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name, Format, ConfirmConversions, ReadOnly, LinkToSource, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Connection, SQLStatement, SQLStatement1, OpenExclusive, SubType);
	}
	void OpenHeaderSource(LPCTSTR Name, VARIANT * Format, VARIANT * ConfirmConversions, VARIANT * ReadOnly, VARIANT * AddToRecentFiles, VARIANT * PasswordDocument, VARIANT * PasswordTemplate, VARIANT * Revert, VARIANT * WritePasswordDocument, VARIANT * WritePasswordTemplate, VARIANT * OpenExclusive)
	{
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x71, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name, Format, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, OpenExclusive);
	}
	void ShowWizard(VARIANT * InitialState, VARIANT * ShowDocumentStep, VARIANT * ShowTemplateStep, VARIANT * ShowDataStep, VARIANT * ShowWriteStep, VARIANT * ShowPreviewStep, VARIANT * ShowMergeStep)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x72, DISPATCH_METHOD, VT_EMPTY, NULL, parms, InitialState, ShowDocumentStep, ShowTemplateStep, ShowDataStep, ShowWriteStep, ShowPreviewStep, ShowMergeStep);
	}

	// MailMerge 属性
public:

};
