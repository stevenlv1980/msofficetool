// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CGraphic0 包装类

class CGraphic0 : public COleDispatchDriver
{
public:
	CGraphic0(){} // 调用 COleDispatchDriver 默认构造函数
	CGraphic0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CGraphic0(const CGraphic0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IGraphic 方法
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
	STDMETHOD(get_Brightness)(float * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PR4 ;
		InvokeHelper(0x892, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Brightness)(float RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x892, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_ColorType)(MsoPictureColorType * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x893, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_ColorType)(MsoPictureColorType RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x893, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Contrast)(float * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PR4 ;
		InvokeHelper(0x894, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Contrast)(float RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x894, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_CropBottom)(float * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PR4 ;
		InvokeHelper(0x895, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_CropBottom)(float RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x895, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_CropLeft)(float * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PR4 ;
		InvokeHelper(0x896, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_CropLeft)(float RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x896, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_CropRight)(float * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PR4 ;
		InvokeHelper(0x897, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_CropRight)(float RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x897, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_CropTop)(float * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PR4 ;
		InvokeHelper(0x898, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_CropTop)(float RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x898, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Filename)(BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBSTR ;
		InvokeHelper(0x587, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Filename)(LPCTSTR RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x587, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Height)(float * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PR4 ;
		InvokeHelper(0x7b, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Height)(float RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x7b, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_LockAspectRatio)(MsoTriState * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x6a4, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_LockAspectRatio)(MsoTriState RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x6a4, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Width)(float * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PR4 ;
		InvokeHelper(0x7a, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Width)(float RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x7a, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}

	// IGraphic 属性
public:

};
