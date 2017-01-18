// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CDocEvents0 包装类

class CDocEvents0 : public COleDispatchDriver
{
public:
	CDocEvents0(){} // 调用 COleDispatchDriver 默认构造函数
	CDocEvents0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDocEvents0(const CDocEvents0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IDocEvents 方法
public:
	STDMETHOD(SelectionChange)(Range * Target)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x607, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Target);
		return result;
	}
	STDMETHOD(BeforeDoubleClick)(Range * Target, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x601, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Target, Cancel);
		return result;
	}
	STDMETHOD(BeforeRightClick)(Range * Target, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x5fe, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Target, Cancel);
		return result;
	}
	STDMETHOD(Activate)()
	{
		HRESULT result;
		InvokeHelper(0x130, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(Deactivate)()
	{
		HRESULT result;
		InvokeHelper(0x5fa, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(Calculate)()
	{
		HRESULT result;
		InvokeHelper(0x117, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(Change)(Range * Target)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x609, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Target);
		return result;
	}
	STDMETHOD(FollowHyperlink)(Hyperlink * Target)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x5be, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Target);
		return result;
	}
	STDMETHOD(PivotTableUpdate)(PivotTable * Target)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x86c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Target);
		return result;
	}

	// IDocEvents 属性
public:

};
