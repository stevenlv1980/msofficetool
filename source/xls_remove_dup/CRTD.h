// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CRTD 包装类

class CRTD : public COleDispatchDriver
{
public:
	CRTD(){} // 调用 COleDispatchDriver 默认构造函数
	CRTD(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CRTD(const CRTD& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IRTD 方法
public:
	STDMETHOD(get_ThrottleInterval)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x8c0, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_ThrottleInterval)(long RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x8c0, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(RefreshData)()
	{
		HRESULT result;
		InvokeHelper(0x8c1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(RestartServers)()
	{
		HRESULT result;
		InvokeHelper(0x8c2, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}

	// IRTD 属性
public:

};
