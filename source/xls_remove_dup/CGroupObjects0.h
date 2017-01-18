// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CGroupObjects0 包装类

class CGroupObjects0 : public COleDispatchDriver
{
public:
	CGroupObjects0(){} // 调用 COleDispatchDriver 默认构造函数
	CGroupObjects0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CGroupObjects0(const CGroupObjects0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IGroupObjects 方法
public:
	STDMETHOD(get_Application)(Application * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Creator)(XlCreator * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Parent)(LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	void _Dummy3()
	{
		InvokeHelper(0x10003, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(BringToFront)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x25a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Copy)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x227, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(CopyPicture)(XlPictureAppearance Appearance, XlCopyPictureFormat Format, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0xd5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Appearance, Format, RHS);
		return result;
	}
	STDMETHOD(Cut)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x235, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Delete)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x75, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Duplicate)(LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x40f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Enabled)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x258, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Enabled)(BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x258, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Height)(double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PR8 ;
		InvokeHelper(0x7b, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Height)(double RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x7b, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	void _Dummy12()
	{
		InvokeHelper(0x1000c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_Left)(double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PR8 ;
		InvokeHelper(0x7f, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Left)(double RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x7f, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Locked)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x10d, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Locked)(BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x10d, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	void _Dummy15()
	{
		InvokeHelper(0x1000f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_OnAction)(BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBSTR ;
		InvokeHelper(0x254, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_OnAction)(LPCTSTR RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x254, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Placement)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x269, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Placement)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x269, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_PrintObject)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x26a, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_PrintObject)(BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x26a, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(Select)(VARIANT Replace, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0xeb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Replace, RHS);
		return result;
	}
	STDMETHOD(SendToBack)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x25d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Top)(double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PR8 ;
		InvokeHelper(0x7e, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Top)(double RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x7e, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	void _Dummy22()
	{
		InvokeHelper(0x10016, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_Visible)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x22e, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Visible)(BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x22e, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Width)(double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PR8 ;
		InvokeHelper(0x7a, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Width)(double RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x7a, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_ZOrder)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x26e, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_ShapeRange)(ShapeRange * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x5f8, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	void _Dummy27()
	{
		InvokeHelper(0x1001b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy28()
	{
		InvokeHelper(0x1001c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_AddIndent)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x427, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_AddIndent)(BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x427, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	void _Dummy30()
	{
		InvokeHelper(0x1001e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_ArrowHeadLength)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x263, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_ArrowHeadLength)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x263, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_ArrowHeadStyle)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x264, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_ArrowHeadStyle)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x264, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_ArrowHeadWidth)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x265, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_ArrowHeadWidth)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x265, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_AutoSize)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x266, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_AutoSize)(BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x266, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Border)(Border * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x80, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	void _Dummy36()
	{
		InvokeHelper(0x10024, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy37()
	{
		InvokeHelper(0x10025, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy38()
	{
		InvokeHelper(0x10026, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(CheckSpelling)(VARIANT CustomDictionary, VARIANT IgnoreUppercase, VARIANT AlwaysSuggest, VARIANT SpellLang, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x1f9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &CustomDictionary, &IgnoreUppercase, &AlwaysSuggest, &SpellLang, RHS);
		return result;
	}
	STDMETHOD(get__Default)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put__Default)(long RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	void _Dummy41()
	{
		InvokeHelper(0x10029, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy42()
	{
		InvokeHelper(0x1002a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy43()
	{
		InvokeHelper(0x1002b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy44()
	{
		InvokeHelper(0x1002c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy45()
	{
		InvokeHelper(0x1002d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_Font)(Font * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x92, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	void _Dummy47()
	{
		InvokeHelper(0x1002f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy48()
	{
		InvokeHelper(0x10030, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_HorizontalAlignment)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x88, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_HorizontalAlignment)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x88, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	void _Dummy50()
	{
		InvokeHelper(0x10032, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_Interior)(Interior * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x81, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	void _Dummy52()
	{
		InvokeHelper(0x10034, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy53()
	{
		InvokeHelper(0x10035, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy54()
	{
		InvokeHelper(0x10036, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy55()
	{
		InvokeHelper(0x10037, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy56()
	{
		InvokeHelper(0x10038, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy57()
	{
		InvokeHelper(0x10039, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy58()
	{
		InvokeHelper(0x1003a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy59()
	{
		InvokeHelper(0x1003b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy60()
	{
		InvokeHelper(0x1003c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy61()
	{
		InvokeHelper(0x1003d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy62()
	{
		InvokeHelper(0x1003e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy63()
	{
		InvokeHelper(0x1003f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_Orientation)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x86, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Orientation)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x86, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	void _Dummy65()
	{
		InvokeHelper(0x10041, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy66()
	{
		InvokeHelper(0x10042, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy67()
	{
		InvokeHelper(0x10043, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy68()
	{
		InvokeHelper(0x10044, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_RoundedCorners)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x26b, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_RoundedCorners)(BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x26b, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	void _Dummy70()
	{
		InvokeHelper(0x10046, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_Shadow)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Shadow)(BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x67, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	void _Dummy72()
	{
		InvokeHelper(0x10048, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy73()
	{
		InvokeHelper(0x10049, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(Ungroup)(LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0xf4, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	void _Dummy75()
	{
		InvokeHelper(0x1004b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_VerticalAlignment)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x89, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_VerticalAlignment)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x89, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	void _Dummy77()
	{
		InvokeHelper(0x1004d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	STDMETHOD(get_ReadingOrder)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x3cf, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_ReadingOrder)(long RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x3cf, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Count)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Group)(GroupObject * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x2e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Item)(VARIANT Index, LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xaa, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, RHS);
		return result;
	}
	STDMETHOD(_NewEnum)(LPUNKNOWN * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PUNKNOWN ;
		InvokeHelper(0xfffffffc, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}

	// IGroupObjects 属性
public:

};
