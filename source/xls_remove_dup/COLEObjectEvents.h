// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// COLEObjectEvents 包装类

class COLEObjectEvents : public COleDispatchDriver
{
public:
	COLEObjectEvents(){} // 调用 COleDispatchDriver 默认构造函数
	COLEObjectEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COLEObjectEvents(const COLEObjectEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IOLEObjectEvents 方法
public:
	STDMETHOD(GotFocus)()
	{
		HRESULT result;
		InvokeHelper(0x605, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(LostFocus)()
	{
		HRESULT result;
		InvokeHelper(0x606, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}

	// IOLEObjectEvents 属性
public:

};
