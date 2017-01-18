// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CWorksheetFunction 包装类

class CWorksheetFunction : public COleDispatchDriver
{
public:
	CWorksheetFunction(){} // 调用 COleDispatchDriver 默认构造函数
	CWorksheetFunction(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWorksheetFunction(const CWorksheetFunction& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IWorksheetFunction 方法
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
	STDMETHOD(_WSFunction)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0xa9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(Count)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4000, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(IsNA)(VARIANT Arg1, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBOOL ;
		InvokeHelper(0x4002, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(IsError)(VARIANT Arg1, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBOOL ;
		InvokeHelper(0x4003, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(Sum)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4004, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(Average)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4005, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(Min)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4006, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(Max)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4007, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(Npv)(double Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x400b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(StDev)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x400c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(Dollar)(double Arg1, VARIANT Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x400d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Fixed)(double Arg1, VARIANT Arg2, VARIANT Arg3, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x400e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(Pi)(double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PR8 ;
		InvokeHelper(0x4013, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Ln)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4016, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Log10)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4017, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Round)(double Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x401b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(Lookup)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x401c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(Index)(VARIANT Arg1, double Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x401d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, Arg2, &Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(Rept)(LPCTSTR Arg1, double Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_R8 VTS_PBSTR ;
		InvokeHelper(0x401e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(And)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PBOOL ;
		InvokeHelper(0x4024, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(Or)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PBOOL ;
		InvokeHelper(0x4025, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(DCount)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4028, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(DSum)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4029, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(DAverage)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x402a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(DMin)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x402b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(DMax)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x402c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(DStDev)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x402d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(Var)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x402e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(DVar)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x402f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(Text)(VARIANT Arg1, LPCTSTR Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_BSTR VTS_PBSTR ;
		InvokeHelper(0x4030, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(LinEst)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x4031, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(Trend)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x4032, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(LogEst)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x4033, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(Growth)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x4034, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(Pv)(double Arg1, double Arg2, double Arg3, VARIANT Arg4, VARIANT Arg5, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4038, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, &Arg5, RHS);
		return result;
	}
	STDMETHOD(Fv)(double Arg1, double Arg2, double Arg3, VARIANT Arg4, VARIANT Arg5, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4039, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, &Arg5, RHS);
		return result;
	}
	STDMETHOD(NPer)(double Arg1, double Arg2, double Arg3, VARIANT Arg4, VARIANT Arg5, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x403a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, &Arg5, RHS);
		return result;
	}
	STDMETHOD(Pmt)(double Arg1, double Arg2, double Arg3, VARIANT Arg4, VARIANT Arg5, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x403b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, &Arg5, RHS);
		return result;
	}
	STDMETHOD(Rate)(double Arg1, double Arg2, double Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x403c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, &Arg5, &Arg6, RHS);
		return result;
	}
	STDMETHOD(MIrr)(VARIANT Arg1, double Arg2, double Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x403d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(Irr)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x403e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Match)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4040, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(Weekday)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4046, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Search)(LPCTSTR Arg1, LPCTSTR Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4052, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(Transpose)(VARIANT Arg1, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x4053, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(Atan2)(double Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4061, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(Asin)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4062, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Acos)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4063, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Choose)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x4064, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(HLookup)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x4065, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(VLookup)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x4066, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(Log)(double Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x406d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Proper)(LPCTSTR Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_PBSTR ;
		InvokeHelper(0x4072, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Trim)(LPCTSTR Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_PBSTR ;
		InvokeHelper(0x4076, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Replace)(LPCTSTR Arg1, double Arg2, double Arg3, LPCTSTR Arg4, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_R8 VTS_R8 VTS_BSTR VTS_PBSTR ;
		InvokeHelper(0x4077, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, RHS);
		return result;
	}
	STDMETHOD(Substitute)(LPCTSTR Arg1, LPCTSTR Arg2, LPCTSTR Arg3, VARIANT Arg4, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BSTR VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4078, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(Find)(LPCTSTR Arg1, LPCTSTR Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x407c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(IsErr)(VARIANT Arg1, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBOOL ;
		InvokeHelper(0x407e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(IsText)(VARIANT Arg1, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBOOL ;
		InvokeHelper(0x407f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(IsNumber)(VARIANT Arg1, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBOOL ;
		InvokeHelper(0x4080, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(Sln)(double Arg1, double Arg2, double Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x408e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(Syd)(double Arg1, double Arg2, double Arg3, double Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x408f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, RHS);
		return result;
	}
	STDMETHOD(Ddb)(double Arg1, double Arg2, double Arg3, double Arg4, VARIANT Arg5, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4090, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, &Arg5, RHS);
		return result;
	}
	STDMETHOD(Clean)(LPCTSTR Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_PBSTR ;
		InvokeHelper(0x40a2, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(MDeterm)(VARIANT Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40a3, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(MInverse)(VARIANT Arg1, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x40a4, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(MMult)(VARIANT Arg1, VARIANT Arg2, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x40a5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Ipmt)(double Arg1, double Arg2, double Arg3, double Arg4, VARIANT Arg5, VARIANT Arg6, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40a7, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, &Arg5, &Arg6, RHS);
		return result;
	}
	STDMETHOD(Ppmt)(double Arg1, double Arg2, double Arg3, double Arg4, VARIANT Arg5, VARIANT Arg6, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40a8, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, &Arg5, &Arg6, RHS);
		return result;
	}
	STDMETHOD(CountA)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40a9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(Product)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40b7, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(Fact)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x40b8, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(DProduct)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40bd, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(IsNonText)(VARIANT Arg1, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBOOL ;
		InvokeHelper(0x40be, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(StDevP)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40c1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(VarP)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40c2, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(DStDevP)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40c3, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(DVarP)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40c4, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(IsLogical)(VARIANT Arg1, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBOOL ;
		InvokeHelper(0x40c6, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(DCountA)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40c7, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(USDollar)(double Arg1, double Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_PBSTR ;
		InvokeHelper(0x40cc, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(FindB)(LPCTSTR Arg1, LPCTSTR Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40cd, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(SearchB)(LPCTSTR Arg1, LPCTSTR Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40ce, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(ReplaceB)(LPCTSTR Arg1, double Arg2, double Arg3, LPCTSTR Arg4, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_R8 VTS_R8 VTS_BSTR VTS_PBSTR ;
		InvokeHelper(0x40cf, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, RHS);
		return result;
	}
	STDMETHOD(RoundUp)(double Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x40d4, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(RoundDown)(double Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x40d5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(Rank)(double Arg1, Range * Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_DISPATCH VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40d8, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(Days360)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40dc, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(Vdb)(double Arg1, double Arg2, double Arg3, double Arg4, double Arg5, VARIANT Arg6, VARIANT Arg7, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40de, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, Arg5, &Arg6, &Arg7, RHS);
		return result;
	}
	STDMETHOD(Median)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40e3, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(SumProduct)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40e4, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(Sinh)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x40e5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Cosh)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x40e6, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Tanh)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x40e7, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Asinh)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x40e8, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Acosh)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x40e9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Atanh)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x40ea, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(DGet)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x40eb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(Db)(double Arg1, double Arg2, double Arg3, double Arg4, VARIANT Arg5, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x40f7, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, &Arg5, RHS);
		return result;
	}
	STDMETHOD(Frequency)(VARIANT Arg1, VARIANT Arg2, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x40fc, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(AveDev)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x410d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(BetaDist)(double Arg1, double Arg2, double Arg3, VARIANT Arg4, VARIANT Arg5, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x410e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, &Arg5, RHS);
		return result;
	}
	STDMETHOD(GammaLn)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x410f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(BetaInv)(double Arg1, double Arg2, double Arg3, VARIANT Arg4, VARIANT Arg5, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4110, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, &Arg5, RHS);
		return result;
	}
	STDMETHOD(BinomDist)(double Arg1, double Arg2, double Arg3, BOOL Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_BOOL VTS_PR8 ;
		InvokeHelper(0x4111, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, RHS);
		return result;
	}
	STDMETHOD(ChiDist)(double Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4112, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(ChiInv)(double Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4113, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(Combin)(double Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4114, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(Confidence)(double Arg1, double Arg2, double Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4115, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(CritBinom)(double Arg1, double Arg2, double Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4116, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(Even)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4117, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(ExponDist)(double Arg1, double Arg2, BOOL Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_BOOL VTS_PR8 ;
		InvokeHelper(0x4118, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(FDist)(double Arg1, double Arg2, double Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4119, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(FInv)(double Arg1, double Arg2, double Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x411a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(Fisher)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x411b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(FisherInv)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x411c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Floor)(double Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x411d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(GammaDist)(double Arg1, double Arg2, double Arg3, BOOL Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_BOOL VTS_PR8 ;
		InvokeHelper(0x411e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, RHS);
		return result;
	}
	STDMETHOD(GammaInv)(double Arg1, double Arg2, double Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x411f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(Ceiling)(double Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4120, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(HypGeomDist)(double Arg1, double Arg2, double Arg3, double Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4121, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, RHS);
		return result;
	}
	STDMETHOD(LogNormDist)(double Arg1, double Arg2, double Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4122, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(LogInv)(double Arg1, double Arg2, double Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4123, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(NegBinomDist)(double Arg1, double Arg2, double Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4124, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(NormDist)(double Arg1, double Arg2, double Arg3, BOOL Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_BOOL VTS_PR8 ;
		InvokeHelper(0x4125, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, RHS);
		return result;
	}
	STDMETHOD(NormSDist)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4126, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(NormInv)(double Arg1, double Arg2, double Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4127, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(NormSInv)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4128, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Standardize)(double Arg1, double Arg2, double Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4129, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(Odd)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x412a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Permut)(double Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x412b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(Poisson)(double Arg1, double Arg2, BOOL Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_BOOL VTS_PR8 ;
		InvokeHelper(0x412c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(TDist)(double Arg1, double Arg2, double Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x412d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, RHS);
		return result;
	}
	STDMETHOD(Weibull)(double Arg1, double Arg2, double Arg3, BOOL Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_BOOL VTS_PR8 ;
		InvokeHelper(0x412e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, RHS);
		return result;
	}
	STDMETHOD(SumXMY2)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x412f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(SumX2MY2)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4130, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(SumX2PY2)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4131, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(ChiTest)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4132, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Correl)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4133, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Covar)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4134, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Forecast)(double Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4135, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(FTest)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4136, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Intercept)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4137, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Pearson)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4138, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(RSq)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4139, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(StEyx)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x413a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Slope)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x413b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(TTest)(VARIANT Arg1, VARIANT Arg2, double Arg3, double Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x413c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, Arg3, Arg4, RHS);
		return result;
	}
	STDMETHOD(Prob)(VARIANT Arg1, VARIANT Arg2, double Arg3, VARIANT Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_R8 VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x413d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(DevSq)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x413e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(GeoMean)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x413f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(HarMean)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4140, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(SumSq)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4141, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(Kurt)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4142, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(Skew)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4143, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(ZTest)(VARIANT Arg1, double Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4144, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(Large)(VARIANT Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4145, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(Small)(VARIANT Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4146, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(Quartile)(VARIANT Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4147, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(Percentile)(VARIANT Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4148, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(PercentRank)(VARIANT Arg1, double Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4149, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(Mode)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x414a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(TrimMean)(VARIANT Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 VTS_PR8 ;
		InvokeHelper(0x414b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(TInv)(double Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x414c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(Power)(double Arg1, double Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4151, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, RHS);
		return result;
	}
	STDMETHOD(Radians)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4156, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Degrees)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4157, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Subtotal)(double Arg1, Range * Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4158, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(SumIf)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4159, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(CountIf)(Range * Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x415a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(CountBlank)(Range * Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PR8 ;
		InvokeHelper(0x415b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Ispmt)(double Arg1, double Arg2, double Arg3, double Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_PR8 ;
		InvokeHelper(0x415e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, RHS);
		return result;
	}
	STDMETHOD(Roman)(double Arg1, VARIANT Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4162, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Asc)(LPCTSTR Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_PBSTR ;
		InvokeHelper(0x40d6, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Dbcs)(LPCTSTR Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_PBSTR ;
		InvokeHelper(0x40d7, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(Phonetic)(Range * Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBSTR ;
		InvokeHelper(0x4168, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(BahtText)(double Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PBSTR ;
		InvokeHelper(0x4170, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(ThaiDayOfWeek)(double Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PBSTR ;
		InvokeHelper(0x4171, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(ThaiDigit)(LPCTSTR Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_PBSTR ;
		InvokeHelper(0x4172, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(ThaiMonthOfYear)(double Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PBSTR ;
		InvokeHelper(0x4173, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(ThaiNumSound)(double Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PBSTR ;
		InvokeHelper(0x4174, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(ThaiNumString)(double Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PBSTR ;
		InvokeHelper(0x4175, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(ThaiStringLength)(LPCTSTR Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_PR8 ;
		InvokeHelper(0x4176, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(IsThaiDigit)(LPCTSTR Arg1, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_PBOOL ;
		InvokeHelper(0x4177, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(RoundBahtDown)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4178, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(RoundBahtUp)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x4179, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(ThaiYear)(double Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R8 VTS_PR8 ;
		InvokeHelper(0x417a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, RHS);
		return result;
	}
	STDMETHOD(RTD)(VARIANT progID, VARIANT server, VARIANT topic1, VARIANT topic2, VARIANT topic3, VARIANT topic4, VARIANT topic5, VARIANT topic6, VARIANT topic7, VARIANT topic8, VARIANT topic9, VARIANT topic10, VARIANT topic11, VARIANT topic12, VARIANT topic13, VARIANT topic14, VARIANT topic15, VARIANT topic16, VARIANT topic17, VARIANT topic18, VARIANT topic19, VARIANT topic20, VARIANT topic21, VARIANT topic22, VARIANT topic23, VARIANT topic24, VARIANT topic25, VARIANT topic26, VARIANT topic27, VARIANT topic28, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x417b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &progID, &server, &topic1, &topic2, &topic3, &topic4, &topic5, &topic6, &topic7, &topic8, &topic9, &topic10, &topic11, &topic12, &topic13, &topic14, &topic15, &topic16, &topic17, &topic18, &topic19, &topic20, &topic21, &topic22, &topic23, &topic24, &topic25, &topic26, &topic27, &topic28, RHS);
		return result;
	}
	STDMETHOD(Hex2Bin)(VARIANT Arg1, VARIANT Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4180, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Hex2Dec)(VARIANT Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4181, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(Hex2Oct)(VARIANT Arg1, VARIANT Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4182, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Dec2Bin)(VARIANT Arg1, VARIANT Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4183, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Dec2Hex)(VARIANT Arg1, VARIANT Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4184, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Dec2Oct)(VARIANT Arg1, VARIANT Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4185, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Oct2Bin)(VARIANT Arg1, VARIANT Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4186, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Oct2Hex)(VARIANT Arg1, VARIANT Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4187, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Oct2Dec)(VARIANT Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4188, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(Bin2Dec)(VARIANT Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4189, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(Bin2Oct)(VARIANT Arg1, VARIANT Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x418a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Bin2Hex)(VARIANT Arg1, VARIANT Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x418b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(ImSub)(VARIANT Arg1, VARIANT Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x418c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(ImDiv)(VARIANT Arg1, VARIANT Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x418d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(ImPower)(VARIANT Arg1, VARIANT Arg2, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x418e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(ImAbs)(VARIANT Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x418f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(ImSqrt)(VARIANT Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4190, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(ImLn)(VARIANT Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4191, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(ImLog2)(VARIANT Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4192, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(ImLog10)(VARIANT Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4193, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(ImSin)(VARIANT Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4194, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(ImCos)(VARIANT Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4195, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(ImExp)(VARIANT Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4196, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(ImArgument)(VARIANT Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4197, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(ImConjugate)(VARIANT Arg1, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x4198, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(Imaginary)(VARIANT Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x4199, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(ImReal)(VARIANT Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x419a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(Complex)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x419b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(ImSum)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x419c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(ImProduct)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x419d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(SeriesSum)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x419e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(FactDouble)(VARIANT Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x419f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(SqrtPi)(VARIANT Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41a0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(Quotient)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41a1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Delta)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41a2, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(GeStep)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41a3, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(IsEven)(VARIANT Arg1, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBOOL ;
		InvokeHelper(0x41a4, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(IsOdd)(VARIANT Arg1, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBOOL ;
		InvokeHelper(0x41a5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(MRound)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41a6, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Erf)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41a7, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(ErfC)(VARIANT Arg1, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41a8, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, RHS);
		return result;
	}
	STDMETHOD(BesselJ)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41a9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(BesselK)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41aa, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(BesselY)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41ab, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(BesselI)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41ac, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Xirr)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41ad, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(Xnpv)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41ae, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(PriceMat)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41af, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, RHS);
		return result;
	}
	STDMETHOD(YieldMat)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41b0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, RHS);
		return result;
	}
	STDMETHOD(IntRate)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41b1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, RHS);
		return result;
	}
	STDMETHOD(Received)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41b2, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, RHS);
		return result;
	}
	STDMETHOD(Disc)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41b3, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, RHS);
		return result;
	}
	STDMETHOD(PriceDisc)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41b4, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, RHS);
		return result;
	}
	STDMETHOD(YieldDisc)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41b5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, RHS);
		return result;
	}
	STDMETHOD(TBillEq)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41b6, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(TBillPrice)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41b7, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(TBillYield)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41b8, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(Price)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41b9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, RHS);
		return result;
	}
	STDMETHOD(DollarDe)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41bb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(DollarFr)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41bc, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Nominal)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41bd, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(Effect)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41be, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(CumPrinc)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41bf, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, RHS);
		return result;
	}
	STDMETHOD(CumIPmt)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41c0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, RHS);
		return result;
	}
	STDMETHOD(EDate)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41c1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(EoMonth)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41c2, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(YearFrac)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41c3, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(CoupDayBs)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41c4, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(CoupDays)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41c5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(CoupDaysNc)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41c6, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(CoupNcd)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41c7, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(CoupNum)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41c8, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(CoupPcd)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41c9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, RHS);
		return result;
	}
	STDMETHOD(Duration)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41ca, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, RHS);
		return result;
	}
	STDMETHOD(MDuration)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41cb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, RHS);
		return result;
	}
	STDMETHOD(OddLPrice)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41cc, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, RHS);
		return result;
	}
	STDMETHOD(OddLYield)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41cd, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, RHS);
		return result;
	}
	STDMETHOD(OddFPrice)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41ce, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, RHS);
		return result;
	}
	STDMETHOD(OddFYield)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41cf, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, RHS);
		return result;
	}
	STDMETHOD(RandBetween)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41d0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(WeekNum)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41d1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(AmorDegrc)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41d2, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, RHS);
		return result;
	}
	STDMETHOD(AmorLinc)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41d3, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, RHS);
		return result;
	}
	STDMETHOD(Convert)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41d4, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(AccrInt)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41d5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, RHS);
		return result;
	}
	STDMETHOD(AccrIntM)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41d6, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, RHS);
		return result;
	}
	STDMETHOD(WorkDay)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41d7, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(NetworkDays)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41d8, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(Gcd)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41d9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(MultiNomial)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41da, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(Lcm)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41db, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(FVSchedule)(VARIANT Arg1, VARIANT Arg2, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41dc, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}
	STDMETHOD(SumIfs)(Range * Arg1, Range * Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41e2, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, RHS);
		return result;
	}
	STDMETHOD(CountIfs)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41e1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(AverageIf)(Range * Arg1, VARIANT Arg2, VARIANT Arg3, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41e3, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, &Arg2, &Arg3, RHS);
		return result;
	}
	STDMETHOD(AverageIfs)(Range * Arg1, Range * Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, double * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PR8 ;
		InvokeHelper(0x41e4, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Arg1, Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, RHS);
		return result;
	}
	STDMETHOD(IfError)(VARIANT Arg1, VARIANT Arg2, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x41e0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, RHS);
		return result;
	}

	// IWorksheetFunction 属性
public:

};
