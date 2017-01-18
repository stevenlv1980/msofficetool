// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CWorksheetFunction0 包装类

class CWorksheetFunction0 : public COleDispatchDriver
{
public:
	CWorksheetFunction0(){} // 调用 COleDispatchDriver 默认构造函数
	CWorksheetFunction0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWorksheetFunction0(const CWorksheetFunction0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// WorksheetFunction 方法
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
	VARIANT _WSFunction(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0xa9, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double Count(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4000, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	BOOL IsNA(VARIANT Arg1)
	{
		BOOL result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4002, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, &Arg1);
		return result;
	}
	BOOL IsError(VARIANT Arg1)
	{
		BOOL result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4003, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, &Arg1);
		return result;
	}
	double Sum(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4004, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double Average(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4005, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double Min(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4006, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double Max(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4007, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double Npv(double Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x400b, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double StDev(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x400c, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	CString Dollar(double Arg1, VARIANT Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_R8 VTS_VARIANT ;
		InvokeHelper(0x400d, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1, &Arg2);
		return result;
	}
	CString Fixed(double Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		CString result;
		static BYTE parms[] = VTS_R8 VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x400e, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	double Pi()
	{
		double result;
		InvokeHelper(0x4013, DISPATCH_METHOD, VT_R8, (void*)&result, NULL);
		return result;
	}
	double Ln(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4016, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double Log10(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4017, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double Round(double Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 ;
		InvokeHelper(0x401b, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2);
		return result;
	}
	VARIANT Lookup(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x401c, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1, &Arg2, &Arg3);
		return result;
	}
	VARIANT Index(VARIANT Arg1, double Arg2, VARIANT Arg3, VARIANT Arg4)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x401d, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1, Arg2, &Arg3, &Arg4);
		return result;
	}
	CString Rept(LPCTSTR Arg1, double Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_BSTR VTS_R8 ;
		InvokeHelper(0x401e, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1, Arg2);
		return result;
	}
	BOOL And(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		BOOL result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4024, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	BOOL Or(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		BOOL result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4025, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double DCount(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4028, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	double DSum(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4029, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	double DAverage(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x402a, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	double DMin(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x402b, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	double DMax(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x402c, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	double DStDev(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x402d, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	double Var(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x402e, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double DVar(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x402f, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	CString Text(VARIANT Arg1, LPCTSTR Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_BSTR ;
		InvokeHelper(0x4030, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, Arg2);
		return result;
	}
	VARIANT LinEst(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4031, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4);
		return result;
	}
	VARIANT Trend(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4032, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4);
		return result;
	}
	VARIANT LogEst(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4033, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4);
		return result;
	}
	VARIANT Growth(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4034, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4);
		return result;
	}
	double Pv(double Arg1, double Arg2, double Arg3, VARIANT Arg4, VARIANT Arg5)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4038, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, &Arg5);
		return result;
	}
	double Fv(double Arg1, double Arg2, double Arg3, VARIANT Arg4, VARIANT Arg5)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4039, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, &Arg5);
		return result;
	}
	double NPer(double Arg1, double Arg2, double Arg3, VARIANT Arg4, VARIANT Arg5)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x403a, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, &Arg5);
		return result;
	}
	double Pmt(double Arg1, double Arg2, double Arg3, VARIANT Arg4, VARIANT Arg5)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x403b, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, &Arg5);
		return result;
	}
	double Rate(double Arg1, double Arg2, double Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x403c, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, &Arg5, &Arg6);
		return result;
	}
	double MIrr(VARIANT Arg1, double Arg2, double Arg3)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 VTS_R8 ;
		InvokeHelper(0x403d, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, Arg2, Arg3);
		return result;
	}
	double Irr(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x403e, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double Match(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4040, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3);
		return result;
	}
	double Weekday(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4046, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double Search(LPCTSTR Arg1, LPCTSTR Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_VARIANT ;
		InvokeHelper(0x4052, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, &Arg3);
		return result;
	}
	VARIANT Transpose(VARIANT Arg1)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4053, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1);
		return result;
	}
	double Atan2(double Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 ;
		InvokeHelper(0x4061, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2);
		return result;
	}
	double Asin(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4062, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double Acos(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4063, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	VARIANT Choose(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4064, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	VARIANT HLookup(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4065, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4);
		return result;
	}
	VARIANT VLookup(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4066, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4);
		return result;
	}
	double Log(double Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_VARIANT ;
		InvokeHelper(0x406d, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2);
		return result;
	}
	CString Proper(LPCTSTR Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x4072, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1);
		return result;
	}
	CString Trim(LPCTSTR Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x4076, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1);
		return result;
	}
	CString Replace(LPCTSTR Arg1, double Arg2, double Arg3, LPCTSTR Arg4)
	{
		CString result;
		static BYTE parms[] = VTS_BSTR VTS_R8 VTS_R8 VTS_BSTR ;
		InvokeHelper(0x4077, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4);
		return result;
	}
	CString Substitute(LPCTSTR Arg1, LPCTSTR Arg2, LPCTSTR Arg3, VARIANT Arg4)
	{
		CString result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BSTR VTS_VARIANT ;
		InvokeHelper(0x4078, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4);
		return result;
	}
	double Find(LPCTSTR Arg1, LPCTSTR Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_VARIANT ;
		InvokeHelper(0x407c, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, &Arg3);
		return result;
	}
	BOOL IsErr(VARIANT Arg1)
	{
		BOOL result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x407e, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, &Arg1);
		return result;
	}
	BOOL IsText(VARIANT Arg1)
	{
		BOOL result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x407f, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, &Arg1);
		return result;
	}
	BOOL IsNumber(VARIANT Arg1)
	{
		BOOL result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4080, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, &Arg1);
		return result;
	}
	double Sln(double Arg1, double Arg2, double Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x408e, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3);
		return result;
	}
	double Syd(double Arg1, double Arg2, double Arg3, double Arg4)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x408f, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4);
		return result;
	}
	double Ddb(double Arg1, double Arg2, double Arg3, double Arg4, VARIANT Arg5)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT ;
		InvokeHelper(0x4090, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, &Arg5);
		return result;
	}
	CString Clean(LPCTSTR Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x40a2, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1);
		return result;
	}
	double MDeterm(VARIANT Arg1)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x40a3, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1);
		return result;
	}
	VARIANT MInverse(VARIANT Arg1)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x40a4, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1);
		return result;
	}
	VARIANT MMult(VARIANT Arg1, VARIANT Arg2)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40a5, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double Ipmt(double Arg1, double Arg2, double Arg3, double Arg4, VARIANT Arg5, VARIANT Arg6)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40a7, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, &Arg5, &Arg6);
		return result;
	}
	double Ppmt(double Arg1, double Arg2, double Arg3, double Arg4, VARIANT Arg5, VARIANT Arg6)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40a8, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, &Arg5, &Arg6);
		return result;
	}
	double CountA(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40a9, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double Product(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40b7, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double Fact(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x40b8, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double DProduct(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40bd, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	BOOL IsNonText(VARIANT Arg1)
	{
		BOOL result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x40be, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, &Arg1);
		return result;
	}
	double StDevP(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40c1, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double VarP(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40c2, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double DStDevP(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40c3, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	double DVarP(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40c4, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	BOOL IsLogical(VARIANT Arg1)
	{
		BOOL result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x40c6, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, &Arg1);
		return result;
	}
	double DCountA(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40c7, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	CString USDollar(double Arg1, double Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_R8 VTS_R8 ;
		InvokeHelper(0x40cc, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1, Arg2);
		return result;
	}
	double FindB(LPCTSTR Arg1, LPCTSTR Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_VARIANT ;
		InvokeHelper(0x40cd, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, &Arg3);
		return result;
	}
	double SearchB(LPCTSTR Arg1, LPCTSTR Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_VARIANT ;
		InvokeHelper(0x40ce, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, &Arg3);
		return result;
	}
	CString ReplaceB(LPCTSTR Arg1, double Arg2, double Arg3, LPCTSTR Arg4)
	{
		CString result;
		static BYTE parms[] = VTS_BSTR VTS_R8 VTS_R8 VTS_BSTR ;
		InvokeHelper(0x40cf, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4);
		return result;
	}
	double RoundUp(double Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 ;
		InvokeHelper(0x40d4, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2);
		return result;
	}
	double RoundDown(double Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 ;
		InvokeHelper(0x40d5, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2);
		return result;
	}
	double Rank(double Arg1, LPDISPATCH Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_DISPATCH VTS_VARIANT ;
		InvokeHelper(0x40d8, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, &Arg3);
		return result;
	}
	double Days360(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40dc, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3);
		return result;
	}
	double Vdb(double Arg1, double Arg2, double Arg3, double Arg4, double Arg5, VARIANT Arg6, VARIANT Arg7)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40de, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, Arg5, &Arg6, &Arg7);
		return result;
	}
	double Median(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40e3, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double SumProduct(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40e4, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double Sinh(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x40e5, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double Cosh(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x40e6, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double Tanh(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x40e7, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double Asinh(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x40e8, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double Acosh(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x40e9, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double Atanh(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x40ea, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	VARIANT DGet(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		VARIANT result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40eb, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	double Db(double Arg1, double Arg2, double Arg3, double Arg4, VARIANT Arg5)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT ;
		InvokeHelper(0x40f7, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4, &Arg5);
		return result;
	}
	VARIANT Frequency(VARIANT Arg1, VARIANT Arg2)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x40fc, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double AveDev(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x410d, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double BetaDist(double Arg1, double Arg2, double Arg3, VARIANT Arg4, VARIANT Arg5)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x410e, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, &Arg5);
		return result;
	}
	double GammaLn(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x410f, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double BetaInv(double Arg1, double Arg2, double Arg3, VARIANT Arg4, VARIANT Arg5)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4110, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, &Arg4, &Arg5);
		return result;
	}
	double BinomDist(double Arg1, double Arg2, double Arg3, BOOL Arg4)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_BOOL ;
		InvokeHelper(0x4111, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4);
		return result;
	}
	double ChiDist(double Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 ;
		InvokeHelper(0x4112, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2);
		return result;
	}
	double ChiInv(double Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 ;
		InvokeHelper(0x4113, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2);
		return result;
	}
	double Combin(double Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 ;
		InvokeHelper(0x4114, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2);
		return result;
	}
	double Confidence(double Arg1, double Arg2, double Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x4115, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3);
		return result;
	}
	double CritBinom(double Arg1, double Arg2, double Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x4116, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3);
		return result;
	}
	double Even(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4117, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double ExponDist(double Arg1, double Arg2, BOOL Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_BOOL ;
		InvokeHelper(0x4118, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3);
		return result;
	}
	double FDist(double Arg1, double Arg2, double Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x4119, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3);
		return result;
	}
	double FInv(double Arg1, double Arg2, double Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x411a, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3);
		return result;
	}
	double Fisher(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x411b, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double FisherInv(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x411c, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double Floor(double Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 ;
		InvokeHelper(0x411d, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2);
		return result;
	}
	double GammaDist(double Arg1, double Arg2, double Arg3, BOOL Arg4)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_BOOL ;
		InvokeHelper(0x411e, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4);
		return result;
	}
	double GammaInv(double Arg1, double Arg2, double Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x411f, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3);
		return result;
	}
	double Ceiling(double Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 ;
		InvokeHelper(0x4120, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2);
		return result;
	}
	double HypGeomDist(double Arg1, double Arg2, double Arg3, double Arg4)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x4121, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4);
		return result;
	}
	double LogNormDist(double Arg1, double Arg2, double Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x4122, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3);
		return result;
	}
	double LogInv(double Arg1, double Arg2, double Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x4123, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3);
		return result;
	}
	double NegBinomDist(double Arg1, double Arg2, double Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x4124, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3);
		return result;
	}
	double NormDist(double Arg1, double Arg2, double Arg3, BOOL Arg4)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_BOOL ;
		InvokeHelper(0x4125, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4);
		return result;
	}
	double NormSDist(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4126, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double NormInv(double Arg1, double Arg2, double Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x4127, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3);
		return result;
	}
	double NormSInv(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4128, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double Standardize(double Arg1, double Arg2, double Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x4129, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3);
		return result;
	}
	double Odd(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x412a, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double Permut(double Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 ;
		InvokeHelper(0x412b, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2);
		return result;
	}
	double Poisson(double Arg1, double Arg2, BOOL Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_BOOL ;
		InvokeHelper(0x412c, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3);
		return result;
	}
	double TDist(double Arg1, double Arg2, double Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x412d, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3);
		return result;
	}
	double Weibull(double Arg1, double Arg2, double Arg3, BOOL Arg4)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_BOOL ;
		InvokeHelper(0x412e, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4);
		return result;
	}
	double SumXMY2(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x412f, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double SumX2MY2(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4130, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double SumX2PY2(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4131, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double ChiTest(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4132, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double Correl(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4133, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double Covar(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4134, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double Forecast(double Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4135, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	double FTest(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4136, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double Intercept(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4137, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double Pearson(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4138, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double RSq(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4139, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double StEyx(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x413a, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double Slope(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x413b, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double TTest(VARIANT Arg1, VARIANT Arg2, double Arg3, double Arg4)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_R8 VTS_R8 ;
		InvokeHelper(0x413c, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, Arg3, Arg4);
		return result;
	}
	double Prob(VARIANT Arg1, VARIANT Arg2, double Arg3, VARIANT Arg4)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_R8 VTS_VARIANT ;
		InvokeHelper(0x413d, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, Arg3, &Arg4);
		return result;
	}
	double DevSq(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x413e, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double GeoMean(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x413f, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double HarMean(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4140, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double SumSq(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4141, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double Kurt(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4142, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double Skew(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4143, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double ZTest(VARIANT Arg1, double Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 VTS_VARIANT ;
		InvokeHelper(0x4144, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, Arg2, &Arg3);
		return result;
	}
	double Large(VARIANT Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 ;
		InvokeHelper(0x4145, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, Arg2);
		return result;
	}
	double Small(VARIANT Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 ;
		InvokeHelper(0x4146, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, Arg2);
		return result;
	}
	double Quartile(VARIANT Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 ;
		InvokeHelper(0x4147, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, Arg2);
		return result;
	}
	double Percentile(VARIANT Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 ;
		InvokeHelper(0x4148, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, Arg2);
		return result;
	}
	double PercentRank(VARIANT Arg1, double Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 VTS_VARIANT ;
		InvokeHelper(0x4149, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, Arg2, &Arg3);
		return result;
	}
	double Mode(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x414a, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double TrimMean(VARIANT Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_R8 ;
		InvokeHelper(0x414b, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, Arg2);
		return result;
	}
	double TInv(double Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 ;
		InvokeHelper(0x414c, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2);
		return result;
	}
	double Power(double Arg1, double Arg2)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 ;
		InvokeHelper(0x4151, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2);
		return result;
	}
	double Radians(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4156, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double Degrees(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4157, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double Subtotal(double Arg1, LPDISPATCH Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4158, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double SumIf(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4159, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	double CountIf(LPDISPATCH Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT ;
		InvokeHelper(0x415a, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2);
		return result;
	}
	double CountBlank(LPDISPATCH Arg1)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x415b, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double Ispmt(double Arg1, double Arg2, double Arg3, double Arg4)
	{
		double result;
		static BYTE parms[] = VTS_R8 VTS_R8 VTS_R8 VTS_R8 ;
		InvokeHelper(0x415e, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4);
		return result;
	}
	CString Roman(double Arg1, VARIANT Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_R8 VTS_VARIANT ;
		InvokeHelper(0x4162, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1, &Arg2);
		return result;
	}
	CString Asc(LPCTSTR Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x40d6, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1);
		return result;
	}
	CString Dbcs(LPCTSTR Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x40d7, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1);
		return result;
	}
	CString Phonetic(LPDISPATCH Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x4168, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1);
		return result;
	}
	CString BahtText(double Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4170, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1);
		return result;
	}
	CString ThaiDayOfWeek(double Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4171, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1);
		return result;
	}
	CString ThaiDigit(LPCTSTR Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x4172, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1);
		return result;
	}
	CString ThaiMonthOfYear(double Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4173, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1);
		return result;
	}
	CString ThaiNumSound(double Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4174, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1);
		return result;
	}
	CString ThaiNumString(double Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4175, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Arg1);
		return result;
	}
	double ThaiStringLength(LPCTSTR Arg1)
	{
		double result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x4176, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	BOOL IsThaiDigit(LPCTSTR Arg1)
	{
		BOOL result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x4177, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Arg1);
		return result;
	}
	double RoundBahtDown(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4178, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double RoundBahtUp(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x4179, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	double ThaiYear(double Arg1)
	{
		double result;
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0x417a, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1);
		return result;
	}
	VARIANT RTD(VARIANT progID, VARIANT server, VARIANT topic1, VARIANT topic2, VARIANT topic3, VARIANT topic4, VARIANT topic5, VARIANT topic6, VARIANT topic7, VARIANT topic8, VARIANT topic9, VARIANT topic10, VARIANT topic11, VARIANT topic12, VARIANT topic13, VARIANT topic14, VARIANT topic15, VARIANT topic16, VARIANT topic17, VARIANT topic18, VARIANT topic19, VARIANT topic20, VARIANT topic21, VARIANT topic22, VARIANT topic23, VARIANT topic24, VARIANT topic25, VARIANT topic26, VARIANT topic27, VARIANT topic28)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x417b, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &progID, &server, &topic1, &topic2, &topic3, &topic4, &topic5, &topic6, &topic7, &topic8, &topic9, &topic10, &topic11, &topic12, &topic13, &topic14, &topic15, &topic16, &topic17, &topic18, &topic19, &topic20, &topic21, &topic22, &topic23, &topic24, &topic25, &topic26, &topic27, &topic28);
		return result;
	}
	CString Hex2Bin(VARIANT Arg1, VARIANT Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4180, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	CString Hex2Dec(VARIANT Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4181, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1);
		return result;
	}
	CString Hex2Oct(VARIANT Arg1, VARIANT Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4182, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	CString Dec2Bin(VARIANT Arg1, VARIANT Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4183, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	CString Dec2Hex(VARIANT Arg1, VARIANT Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4184, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	CString Dec2Oct(VARIANT Arg1, VARIANT Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4185, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	CString Oct2Bin(VARIANT Arg1, VARIANT Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4186, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	CString Oct2Hex(VARIANT Arg1, VARIANT Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x4187, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	CString Oct2Dec(VARIANT Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4188, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1);
		return result;
	}
	CString Bin2Dec(VARIANT Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4189, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1);
		return result;
	}
	CString Bin2Oct(VARIANT Arg1, VARIANT Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x418a, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	CString Bin2Hex(VARIANT Arg1, VARIANT Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x418b, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	CString ImSub(VARIANT Arg1, VARIANT Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x418c, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	CString ImDiv(VARIANT Arg1, VARIANT Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x418d, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	CString ImPower(VARIANT Arg1, VARIANT Arg2)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x418e, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	CString ImAbs(VARIANT Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x418f, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1);
		return result;
	}
	CString ImSqrt(VARIANT Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4190, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1);
		return result;
	}
	CString ImLn(VARIANT Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4191, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1);
		return result;
	}
	CString ImLog2(VARIANT Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4192, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1);
		return result;
	}
	CString ImLog10(VARIANT Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4193, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1);
		return result;
	}
	CString ImSin(VARIANT Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4194, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1);
		return result;
	}
	CString ImCos(VARIANT Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4195, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1);
		return result;
	}
	CString ImExp(VARIANT Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4196, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1);
		return result;
	}
	CString ImArgument(VARIANT Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4197, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1);
		return result;
	}
	CString ImConjugate(VARIANT Arg1)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4198, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1);
		return result;
	}
	double Imaginary(VARIANT Arg1)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x4199, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1);
		return result;
	}
	double ImReal(VARIANT Arg1)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x419a, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1);
		return result;
	}
	CString Complex(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x419b, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2, &Arg3);
		return result;
	}
	CString ImSum(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x419c, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	CString ImProduct(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		CString result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x419d, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double SeriesSum(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x419e, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4);
		return result;
	}
	double FactDouble(VARIANT Arg1)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x419f, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1);
		return result;
	}
	double SqrtPi(VARIANT Arg1)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x41a0, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1);
		return result;
	}
	double Quotient(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41a1, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double Delta(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41a2, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double GeStep(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41a3, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	BOOL IsEven(VARIANT Arg1)
	{
		BOOL result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x41a4, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, &Arg1);
		return result;
	}
	BOOL IsOdd(VARIANT Arg1)
	{
		BOOL result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x41a5, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, &Arg1);
		return result;
	}
	double MRound(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41a6, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double Erf(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41a7, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double ErfC(VARIANT Arg1)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x41a8, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1);
		return result;
	}
	double BesselJ(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41a9, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double BesselK(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41aa, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double BesselY(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41ab, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double BesselI(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41ac, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double Xirr(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41ad, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3);
		return result;
	}
	double Xnpv(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41ae, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double PriceMat(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41af, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6);
		return result;
	}
	double YieldMat(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41b0, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6);
		return result;
	}
	double IntRate(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41b1, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5);
		return result;
	}
	double Received(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41b2, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5);
		return result;
	}
	double Disc(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41b3, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5);
		return result;
	}
	double PriceDisc(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41b4, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5);
		return result;
	}
	double YieldDisc(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41b5, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5);
		return result;
	}
	double TBillEq(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41b6, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3);
		return result;
	}
	double TBillPrice(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41b7, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3);
		return result;
	}
	double TBillYield(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41b8, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3);
		return result;
	}
	double Price(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41b9, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7);
		return result;
	}
	double DollarDe(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41bb, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double DollarFr(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41bc, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double Nominal(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41bd, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double Effect(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41be, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double CumPrinc(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41bf, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6);
		return result;
	}
	double CumIPmt(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41c0, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6);
		return result;
	}
	double EDate(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41c1, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double EoMonth(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41c2, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double YearFrac(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41c3, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3);
		return result;
	}
	double CoupDayBs(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41c4, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4);
		return result;
	}
	double CoupDays(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41c5, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4);
		return result;
	}
	double CoupDaysNc(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41c6, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4);
		return result;
	}
	double CoupNcd(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41c7, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4);
		return result;
	}
	double CoupNum(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41c8, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4);
		return result;
	}
	double CoupPcd(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41c9, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4);
		return result;
	}
	double Duration(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41ca, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6);
		return result;
	}
	double MDuration(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41cb, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6);
		return result;
	}
	double OddLPrice(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41cc, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8);
		return result;
	}
	double OddLYield(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41cd, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8);
		return result;
	}
	double OddFPrice(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41ce, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9);
		return result;
	}
	double OddFYield(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41cf, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9);
		return result;
	}
	double RandBetween(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41d0, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double WeekNum(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41d1, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double AmorDegrc(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41d2, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7);
		return result;
	}
	double AmorLinc(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41d3, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7);
		return result;
	}
	double Convert(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41d4, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3);
		return result;
	}
	double AccrInt(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41d5, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7);
		return result;
	}
	double AccrIntM(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41d6, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5);
		return result;
	}
	double WorkDay(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41d7, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3);
		return result;
	}
	double NetworkDays(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41d8, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3);
		return result;
	}
	double Gcd(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41d9, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double MultiNomial(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41da, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double Lcm(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41db, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double FVSchedule(VARIANT Arg1, VARIANT Arg2)
	{
		double result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41dc, DISPATCH_METHOD, VT_R8, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}
	double SumIfs(LPDISPATCH Arg1, LPDISPATCH Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41e2, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29);
		return result;
	}
	double CountIfs(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41e1, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
		return result;
	}
	double AverageIf(LPDISPATCH Arg1, VARIANT Arg2, VARIANT Arg3)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41e3, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, &Arg2, &Arg3);
		return result;
	}
	double AverageIfs(LPDISPATCH Arg1, LPDISPATCH Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29)
	{
		double result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41e4, DISPATCH_METHOD, VT_R8, (void*)&result, parms, Arg1, Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29);
		return result;
	}
	VARIANT IfError(VARIANT Arg1, VARIANT Arg2)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x41e0, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Arg1, &Arg2);
		return result;
	}

	// WorksheetFunction 属性
public:

};
