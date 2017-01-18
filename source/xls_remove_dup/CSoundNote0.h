// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CSoundNote0 包装类

class CSoundNote0 : public COleDispatchDriver
{
public:
	CSoundNote0(){} // 调用 COleDispatchDriver 默认构造函数
	CSoundNote0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSoundNote0(const CSoundNote0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// SoundNote 方法
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
	VARIANT Delete()
	{
		VARIANT result;
		InvokeHelper(0x75, DISPATCH_METHOD, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	VARIANT Import(LPCTSTR Filename)
	{
		VARIANT result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x395, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, Filename);
		return result;
	}
	VARIANT Play()
	{
		VARIANT result;
		InvokeHelper(0x396, DISPATCH_METHOD, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	VARIANT Record()
	{
		VARIANT result;
		InvokeHelper(0x397, DISPATCH_METHOD, VT_VARIANT, (void*)&result, NULL);
		return result;
	}

	// SoundNote 属性
public:

};
