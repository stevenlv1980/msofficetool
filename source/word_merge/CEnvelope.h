// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CEnvelope 包装类

class CEnvelope : public COleDispatchDriver
{
public:
	CEnvelope(){} // 调用 COleDispatchDriver 默认构造函数
	CEnvelope(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CEnvelope(const CEnvelope& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Envelope 方法
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
	LPDISPATCH get_Address()
	{
		LPDISPATCH result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ReturnAddress()
	{
		LPDISPATCH result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_DefaultPrintBarCode()
	{
		BOOL result;
		InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DefaultPrintBarCode(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_DefaultPrintFIMA()
	{
		BOOL result;
		InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DefaultPrintFIMA(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_DefaultHeight()
	{
		float result;
		InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_DefaultHeight(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_DefaultWidth()
	{
		float result;
		InvokeHelper(0x7, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_DefaultWidth(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_DefaultSize()
	{
		CString result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_DefaultSize(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_DefaultOmitReturnAddress()
	{
		BOOL result;
		InvokeHelper(0x9, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DefaultOmitReturnAddress(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_FeedSource()
	{
		long result;
		InvokeHelper(0xc, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_FeedSource(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xc, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_AddressFromLeft()
	{
		float result;
		InvokeHelper(0xd, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_AddressFromLeft(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0xd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_AddressFromTop()
	{
		float result;
		InvokeHelper(0xe, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_AddressFromTop(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0xe, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_ReturnAddressFromLeft()
	{
		float result;
		InvokeHelper(0xf, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_ReturnAddressFromLeft(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0xf, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_ReturnAddressFromTop()
	{
		float result;
		InvokeHelper(0x10, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_ReturnAddressFromTop(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x10, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_AddressStyle()
	{
		LPDISPATCH result;
		InvokeHelper(0x11, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ReturnAddressStyle()
	{
		LPDISPATCH result;
		InvokeHelper(0x12, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_DefaultOrientation()
	{
		long result;
		InvokeHelper(0x13, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_DefaultOrientation(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x13, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_DefaultFaceUp()
	{
		BOOL result;
		InvokeHelper(0x14, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DefaultFaceUp(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x14, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void Insert2000(VARIANT * ExtractAddress, VARIANT * Address, VARIANT * AutoText, VARIANT * OmitReturnAddress, VARIANT * ReturnAddress, VARIANT * ReturnAutoText, VARIANT * PrintBarCode, VARIANT * PrintFIMA, VARIANT * Size, VARIANT * Height, VARIANT * Width, VARIANT * FeedSource, VARIANT * AddressFromLeft, VARIANT * AddressFromTop, VARIANT * ReturnAddressFromLeft, VARIANT * ReturnAddressFromTop, VARIANT * DefaultFaceUp, VARIANT * DefaultOrientation)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ExtractAddress, Address, AutoText, OmitReturnAddress, ReturnAddress, ReturnAutoText, PrintBarCode, PrintFIMA, Size, Height, Width, FeedSource, AddressFromLeft, AddressFromTop, ReturnAddressFromLeft, ReturnAddressFromTop, DefaultFaceUp, DefaultOrientation);
	}
	void PrintOut2000(VARIANT * ExtractAddress, VARIANT * Address, VARIANT * AutoText, VARIANT * OmitReturnAddress, VARIANT * ReturnAddress, VARIANT * ReturnAutoText, VARIANT * PrintBarCode, VARIANT * PrintFIMA, VARIANT * Size, VARIANT * Height, VARIANT * Width, VARIANT * FeedSource, VARIANT * AddressFromLeft, VARIANT * AddressFromTop, VARIANT * ReturnAddressFromLeft, VARIANT * ReturnAddressFromTop, VARIANT * DefaultFaceUp, VARIANT * DefaultOrientation)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ExtractAddress, Address, AutoText, OmitReturnAddress, ReturnAddress, ReturnAutoText, PrintBarCode, PrintFIMA, Size, Height, Width, FeedSource, AddressFromLeft, AddressFromTop, ReturnAddressFromLeft, ReturnAddressFromTop, DefaultFaceUp, DefaultOrientation);
	}
	void UpdateDocument()
	{
		InvokeHelper(0x67, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Options()
	{
		InvokeHelper(0x68, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	BOOL get_Vertical()
	{
		BOOL result;
		InvokeHelper(0x16, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_Vertical(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x16, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_RecipientNamefromLeft()
	{
		float result;
		InvokeHelper(0x17, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_RecipientNamefromLeft(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x17, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_RecipientNamefromTop()
	{
		float result;
		InvokeHelper(0x18, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_RecipientNamefromTop(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x18, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_RecipientPostalfromLeft()
	{
		float result;
		InvokeHelper(0x19, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_RecipientPostalfromLeft(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x19, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_RecipientPostalfromTop()
	{
		float result;
		InvokeHelper(0x1a, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_RecipientPostalfromTop(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x1a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_SenderNamefromLeft()
	{
		float result;
		InvokeHelper(0x1b, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_SenderNamefromLeft(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x1b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_SenderNamefromTop()
	{
		float result;
		InvokeHelper(0x1c, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_SenderNamefromTop(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x1c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_SenderPostalfromLeft()
	{
		float result;
		InvokeHelper(0x1d, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_SenderPostalfromLeft(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x1d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_SenderPostalfromTop()
	{
		float result;
		InvokeHelper(0x1e, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_SenderPostalfromTop(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x1e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void Insert(VARIANT * ExtractAddress, VARIANT * Address, VARIANT * AutoText, VARIANT * OmitReturnAddress, VARIANT * ReturnAddress, VARIANT * ReturnAutoText, VARIANT * PrintBarCode, VARIANT * PrintFIMA, VARIANT * Size, VARIANT * Height, VARIANT * Width, VARIANT * FeedSource, VARIANT * AddressFromLeft, VARIANT * AddressFromTop, VARIANT * ReturnAddressFromLeft, VARIANT * ReturnAddressFromTop, VARIANT * DefaultFaceUp, VARIANT * DefaultOrientation, VARIANT * PrintEPostage, VARIANT * Vertical, VARIANT * RecipientNamefromLeft, VARIANT * RecipientNamefromTop, VARIANT * RecipientPostalfromLeft, VARIANT * RecipientPostalfromTop, VARIANT * SenderNamefromLeft, VARIANT * SenderNamefromTop, VARIANT * SenderPostalfromLeft, VARIANT * SenderPostalfromTop)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x69, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ExtractAddress, Address, AutoText, OmitReturnAddress, ReturnAddress, ReturnAutoText, PrintBarCode, PrintFIMA, Size, Height, Width, FeedSource, AddressFromLeft, AddressFromTop, ReturnAddressFromLeft, ReturnAddressFromTop, DefaultFaceUp, DefaultOrientation, PrintEPostage, Vertical, RecipientNamefromLeft, RecipientNamefromTop, RecipientPostalfromLeft, RecipientPostalfromTop, SenderNamefromLeft, SenderNamefromTop, SenderPostalfromLeft, SenderPostalfromTop);
	}
	void PrintOut(VARIANT * ExtractAddress, VARIANT * Address, VARIANT * AutoText, VARIANT * OmitReturnAddress, VARIANT * ReturnAddress, VARIANT * ReturnAutoText, VARIANT * PrintBarCode, VARIANT * PrintFIMA, VARIANT * Size, VARIANT * Height, VARIANT * Width, VARIANT * FeedSource, VARIANT * AddressFromLeft, VARIANT * AddressFromTop, VARIANT * ReturnAddressFromLeft, VARIANT * ReturnAddressFromTop, VARIANT * DefaultFaceUp, VARIANT * DefaultOrientation, VARIANT * PrintEPostage, VARIANT * Vertical, VARIANT * RecipientNamefromLeft, VARIANT * RecipientNamefromTop, VARIANT * RecipientPostalfromLeft, VARIANT * RecipientPostalfromTop, VARIANT * SenderNamefromLeft, VARIANT * SenderNamefromTop, VARIANT * SenderPostalfromLeft, VARIANT * SenderPostalfromTop)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x6a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ExtractAddress, Address, AutoText, OmitReturnAddress, ReturnAddress, ReturnAutoText, PrintBarCode, PrintFIMA, Size, Height, Width, FeedSource, AddressFromLeft, AddressFromTop, ReturnAddressFromLeft, ReturnAddressFromTop, DefaultFaceUp, DefaultOrientation, PrintEPostage, Vertical, RecipientNamefromLeft, RecipientNamefromTop, RecipientPostalfromLeft, RecipientPostalfromTop, SenderNamefromLeft, SenderNamefromTop, SenderPostalfromLeft, SenderPostalfromTop);
	}

	// Envelope 属性
public:

};
