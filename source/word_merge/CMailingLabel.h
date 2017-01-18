// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CMailingLabel 包装类

class CMailingLabel : public COleDispatchDriver
{
public:
	CMailingLabel(){} // 调用 COleDispatchDriver 默认构造函数
	CMailingLabel(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CMailingLabel(const CMailingLabel& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// MailingLabel 方法
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
	BOOL get_DefaultPrintBarCode()
	{
		BOOL result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DefaultPrintBarCode(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_DefaultLaserTray()
	{
		long result;
		InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_DefaultLaserTray(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_CustomLabels()
	{
		LPDISPATCH result;
		InvokeHelper(0x8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_DefaultLabelName()
	{
		CString result;
		InvokeHelper(0x9, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_DefaultLabelName(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH CreateNewDocument2000(VARIANT * Name, VARIANT * Address, VARIANT * AutoText, VARIANT * ExtractAddress, VARIANT * LaserTray)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, Address, AutoText, ExtractAddress, LaserTray);
		return result;
	}
	void PrintOut2000(VARIANT * Name, VARIANT * Address, VARIANT * ExtractAddress, VARIANT * LaserTray, VARIANT * SingleLabel, VARIANT * Row, VARIANT * Column)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name, Address, ExtractAddress, LaserTray, SingleLabel, Row, Column);
	}
	void LabelOptions()
	{
		InvokeHelper(0x67, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH CreateNewDocument(VARIANT * Name, VARIANT * Address, VARIANT * AutoText, VARIANT * ExtractAddress, VARIANT * LaserTray, VARIANT * PrintEPostageLabel, VARIANT * Vertical)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x68, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, Address, AutoText, ExtractAddress, LaserTray, PrintEPostageLabel, Vertical);
		return result;
	}
	void PrintOut(VARIANT * Name, VARIANT * Address, VARIANT * ExtractAddress, VARIANT * LaserTray, VARIANT * SingleLabel, VARIANT * Row, VARIANT * Column, VARIANT * PrintEPostageLabel, VARIANT * Vertical)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x69, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Name, Address, ExtractAddress, LaserTray, SingleLabel, Row, Column, PrintEPostageLabel, Vertical);
	}
	BOOL get_Vertical()
	{
		BOOL result;
		InvokeHelper(0xa, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_Vertical(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH CreateNewDocumentByID(VARIANT * LabelID, VARIANT * Address, VARIANT * AutoText, VARIANT * ExtractAddress, VARIANT * LaserTray, VARIANT * PrintEPostageLabel, VARIANT * Vertical)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x6a, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, LabelID, Address, AutoText, ExtractAddress, LaserTray, PrintEPostageLabel, Vertical);
		return result;
	}
	void PrintOutByID(VARIANT * LabelID, VARIANT * Address, VARIANT * ExtractAddress, VARIANT * LaserTray, VARIANT * SingleLabel, VARIANT * Row, VARIANT * Column, VARIANT * PrintEPostageLabel, VARIANT * Vertical)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x6b, DISPATCH_METHOD, VT_EMPTY, NULL, parms, LabelID, Address, ExtractAddress, LaserTray, SingleLabel, Row, Column, PrintEPostageLabel, Vertical);
	}

	// MailingLabel 属性
public:

};
