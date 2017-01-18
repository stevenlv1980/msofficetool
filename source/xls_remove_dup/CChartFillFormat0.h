// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CChartFillFormat0 包装类

class CChartFillFormat0 : public COleDispatchDriver
{
public:
	CChartFillFormat0(){} // 调用 COleDispatchDriver 默认构造函数
	CChartFillFormat0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CChartFillFormat0(const CChartFillFormat0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IChartFillFormat 方法
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
	STDMETHOD(OneColorGradient)(MsoGradientStyle Style, long Variant, float Degree)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_R4 ;
		InvokeHelper(0x655, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Style, Variant, Degree);
		return result;
	}
	STDMETHOD(TwoColorGradient)(MsoGradientStyle Style, long Variant)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0x658, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Style, Variant);
		return result;
	}
	STDMETHOD(PresetTextured)(MsoPresetTexture PresetTexture)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x659, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, PresetTexture);
		return result;
	}
	STDMETHOD(Solid)()
	{
		HRESULT result;
		InvokeHelper(0x65b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(Patterned)(MsoPatternType Pattern)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x65c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Pattern);
		return result;
	}
	STDMETHOD(UserPicture)(VARIANT PictureFile, VARIANT PictureFormat, VARIANT PictureStackUnit, VARIANT PicturePlacement)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x65d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &PictureFile, &PictureFormat, &PictureStackUnit, &PicturePlacement);
		return result;
	}
	STDMETHOD(UserTextured)(LPCTSTR TextureFile)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x662, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, TextureFile);
		return result;
	}
	STDMETHOD(PresetGradient)(MsoGradientStyle Style, long Variant, MsoPresetGradientType PresetGradientType)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 ;
		InvokeHelper(0x664, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Style, Variant, PresetGradientType);
		return result;
	}
	STDMETHOD(get_BackColor)(ChartColorFormat * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x666, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_ForeColor)(ChartColorFormat * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x667, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_GradientColorType)(MsoGradientColorType * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x668, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_GradientDegree)(float * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PR4 ;
		InvokeHelper(0x669, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_GradientStyle)(MsoGradientStyle * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x66a, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_GradientVariant)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x66b, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Pattern)(MsoPatternType * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x5f, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_PresetGradientType)(MsoPresetGradientType * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x665, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_PresetTexture)(MsoPresetTexture * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x65a, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_TextureName)(BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBSTR ;
		InvokeHelper(0x66c, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_TextureType)(MsoTextureType * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x66d, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Type)(MsoFillType * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Visible)(MsoTriState * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x22e, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Visible)(MsoTriState RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x22e, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}

	// IChartFillFormat 属性
public:

};
