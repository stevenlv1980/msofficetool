// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// COLEObjectEvents0 包装类

class COLEObjectEvents0 : public COleDispatchDriver
{
public:
	COLEObjectEvents0(){} // 调用 COleDispatchDriver 默认构造函数
	COLEObjectEvents0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COLEObjectEvents0(const COLEObjectEvents0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// OLEObjectEvents 方法
public:
	void GotFocus()
	{
		InvokeHelper(0x605, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void LostFocus()
	{
		InvokeHelper(0x606, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// OLEObjectEvents 属性
public:

};
