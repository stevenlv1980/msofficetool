// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CGraphic 包装类

class CGraphic : public COleDispatchDriver
{
public:
	CGraphic(){} // 调用 COleDispatchDriver 默认构造函数
	CGraphic(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CGraphic(const CGraphic& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Graphic 方法
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
	float get_Brightness()
	{
		float result;
		InvokeHelper(0x892, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Brightness(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x892, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_ColorType()
	{
		long result;
		InvokeHelper(0x893, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ColorType(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x893, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_Contrast()
	{
		float result;
		InvokeHelper(0x894, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Contrast(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x894, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_CropBottom()
	{
		float result;
		InvokeHelper(0x895, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_CropBottom(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x895, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_CropLeft()
	{
		float result;
		InvokeHelper(0x896, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_CropLeft(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x896, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_CropRight()
	{
		float result;
		InvokeHelper(0x897, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_CropRight(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x897, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_CropTop()
	{
		float result;
		InvokeHelper(0x898, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_CropTop(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x898, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_Filename()
	{
		CString result;
		InvokeHelper(0x587, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Filename(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x587, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_Height()
	{
		float result;
		InvokeHelper(0x7b, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Height(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x7b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_LockAspectRatio()
	{
		long result;
		InvokeHelper(0x6a4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_LockAspectRatio(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x6a4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_Width()
	{
		float result;
		InvokeHelper(0x7a, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Width(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x7a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// Graphic 属性
public:

};
