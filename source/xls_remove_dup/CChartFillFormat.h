// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CChartFillFormat 包装类

class CChartFillFormat : public COleDispatchDriver
{
public:
	CChartFillFormat(){} // 调用 COleDispatchDriver 默认构造函数
	CChartFillFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CChartFillFormat(const CChartFillFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ChartFillFormat 方法
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
	void OneColorGradient(long Style, long Variant, float Degree)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_R4 ;
		InvokeHelper(0x655, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Style, Variant, Degree);
	}
	void TwoColorGradient(long Style, long Variant)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0x658, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Style, Variant);
	}
	void PresetTextured(long PresetTexture)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x659, DISPATCH_METHOD, VT_EMPTY, NULL, parms, PresetTexture);
	}
	void Solid()
	{
		InvokeHelper(0x65b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Patterned(long Pattern)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x65c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Pattern);
	}
	void UserPicture(VARIANT PictureFile, VARIANT PictureFormat, VARIANT PictureStackUnit, VARIANT PicturePlacement)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x65d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &PictureFile, &PictureFormat, &PictureStackUnit, &PicturePlacement);
	}
	void UserTextured(LPCTSTR TextureFile)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x662, DISPATCH_METHOD, VT_EMPTY, NULL, parms, TextureFile);
	}
	void PresetGradient(long Style, long Variant, long PresetGradientType)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 ;
		InvokeHelper(0x664, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Style, Variant, PresetGradientType);
	}
	LPDISPATCH get_BackColor()
	{
		LPDISPATCH result;
		InvokeHelper(0x666, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ForeColor()
	{
		LPDISPATCH result;
		InvokeHelper(0x667, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_GradientColorType()
	{
		long result;
		InvokeHelper(0x668, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	float get_GradientDegree()
	{
		float result;
		InvokeHelper(0x669, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	long get_GradientStyle()
	{
		long result;
		InvokeHelper(0x66a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_GradientVariant()
	{
		long result;
		InvokeHelper(0x66b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Pattern()
	{
		long result;
		InvokeHelper(0x5f, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_PresetGradientType()
	{
		long result;
		InvokeHelper(0x665, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_PresetTexture()
	{
		long result;
		InvokeHelper(0x65a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	CString get_TextureName()
	{
		CString result;
		InvokeHelper(0x66c, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	long get_TextureType()
	{
		long result;
		InvokeHelper(0x66d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Visible()
	{
		long result;
		InvokeHelper(0x22e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Visible(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x22e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// ChartFillFormat 属性
public:

};
