using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace TestCSServer
{
    // The VariantStructGeneric structure is taken from 
    // "Defining a VARIANT Structure in Managed Code Part 2"
    // https://limbioliong.wordpress.com/2011/09/19/defining-a-variant-structure-in-managed-code-part-2/
    //
    [StructLayout(LayoutKind.Sequential)]
    internal struct CY_internal_struct_01
    {
        public UInt32 Lo;
        public Int32 Hi;
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct CY
    {
        [FieldOffset(0)]
        public CY_internal_struct_01 cy_internal_struct_01;
        [FieldOffset(0)]
        public long int64;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct __tagBRECORD
    {
        public IntPtr pvRecord;
        public IntPtr pRecInfo;
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct VariantUnion
    {
        [FieldOffset(0)]
        public byte ubyte_data;  // This is the BYTE in C++.
        [FieldOffset(0)]
        public sbyte sbyte_data;  // This is the CHAR in C++.
        [FieldOffset(0)]
        public UInt16 uint16_data; // This is the USHORT in C++.
        [FieldOffset(0)]
        public Int16 int16_data;  // This is the SHORT in C++. Also used for the VARIANT_BOOL.
        [FieldOffset(0)]
        public UInt32 uint32_data; // This is the ULONG in C++. Also for the UINT.
        [FieldOffset(0)]
        public Int32 int32_data;  // This is the LONG in C++. Also used for the SCODE, the INT.
        [FieldOffset(0)]
        public ulong ulong_data;  // This is the ULONGLONG in C++.
        [FieldOffset(0)]
        public long long_data;   // This is the LONGLONG in C++.
        [FieldOffset(0)]
        public float float_data;  // This is the FLOAT in C++.
        [FieldOffset(0)]
        public double double_data; // This is the DOUBLE in C++. Also used for the DATE.
        [FieldOffset(0)]
        public IntPtr pointer_data;// Used for BSTR and all other pointer types.
        [FieldOffset(0)]
        public CY cy_data;  // This is the CY structure in C++.
        [FieldOffset(0)]
        public __tagBRECORD record_data; // This is the __tagBRECORD structure in C++.
    };

    [StructLayout(LayoutKind.Sequential)]
    internal struct VariantStructure
    {
        public ushort vt;
        public ushort wReserved1;
        public ushort wReserved2;
        public ushort wReserved3;
        public VariantUnion variant_union;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct Decimal_internal_struct_01
    {
        public byte scale;
        public byte sign;
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct Decimal_internal_union_01
    {
        [FieldOffset(0)]
        public Decimal_internal_struct_01 decimal_internal_struct_01;
        [FieldOffset(0)]
        public ushort signscale;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct Decimal_internal_struct_02
    {
        public UInt32 Lo32;
        public UInt32 Mid32;
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct Decimal_internal_union_02
    {
        [FieldOffset(0)]
        public Decimal_internal_struct_02 decimal_internal_struct_02;
        [FieldOffset(0)]
        public ulong Lo64;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct DecimalStructure
    {
        public ushort wReserved;
        public Decimal_internal_union_01 decimal_internal_union_01;
        public UInt32 Hi32;
        public Decimal_internal_union_02 decimal_internal_union_02;
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct VariantStructGeneric
    {
        [FieldOffset(0)]
        public VariantStructure variant_part;
        [FieldOffset(0)]
        public DecimalStructure decimal_part;
    }

    // The following definition of IDispatch is based on the code provided in
    // "CustomQueryInterface - IDispatch and Aggregation" (http://clrinterop.codeplex.com/releases/view/32350)
    /// <summary>
    /// This interface is the managed version of the native com IDispatch, it has the same vtable layout
    /// as well as the GUID of the native IDispatch interface.
    /// </summary>
    [ComImport]
    [Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDispatch
    {
        void GetTypeInfoCount(out uint pctinfo);

        void GetTypeInfo(uint iTInfo, int lcid, out IntPtr info);

        [PreserveSig]
        UInt32 GetIDsOfNames
        (
            ref Guid iid,
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 2)]
            string[] names,
            int cNames,
            int lcid,
            [Out][MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4, SizeParamIndex = 2)]
            int[] rgDispId
        );

        void Invoke(
            int dispId,
            ref Guid riid,
            int lcid,
            ushort wFlags,
            ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams,
            out object result,
            IntPtr pExcepInfo,
            IntPtr puArgErr);
    }

    public class Constants
    {
        public const UInt32 DISP_E_NONAMEDARGS = 0x80020007;
    }

    [ComVisible(true)]
    [Guid("2A189B2E-774A-4412-808E-09DF04529F5A")]
    [ProgId("TestCSServer.TestClass")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class TestClass : ICustomQueryInterface, IDispatch
    {
        public TestClass()
        {

        }

        #region ICustomQueryInterface implementation
        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;

            if (iid == typeof(IDispatch).GUID)
            {
                // CustomQueryInterfaceMode.Ignore is used below to notify CLR to bypass the invocation of
                // ICustomQueryInterface.GetInterface to avoid infinite loop during QI.
                //
                // We return a pointer to this object's IDispatch interface to the caller.
                ppv = Marshal.GetComInterfaceForObject(this, typeof(IDispatch), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }

            // Let CLR handle the rest of the QI
            return CustomQueryInterfaceResult.NotHandled;
        }
        #endregion

        #region IDispatch implementation
        public void GetTypeInfoCount(out uint pctinfo)
        {
            //NOT IMPLEMENTED
            pctinfo = 0;
        }

        public void GetTypeInfo(uint iTInfo, int lcid, out IntPtr info)
        {
            //NOT IMPLEMENTED
            info = IntPtr.Zero;
        }

        [PreserveSig]
        public UInt32 GetIDsOfNames
        (
            ref Guid iid,
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 2)]
            string[] names,
            int cNames,
            int lcid,
            [Out][MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4, SizeParamIndex = 2)]
            int[] rgDispId
        )
        {
            UInt32 uiRet = 0;

            // ignore lcid, let's assume it is eng
            // Add strict check for parameters Here
            for (int i = 0; i < cNames; i++)
            {
                string name = names[i];

                switch (names[i])
                {
                    case "TestMethod":
                        {
                            rgDispId[i] = 1;
                            break;
                        }

                    default:
                        {
                            rgDispId[i] = -1;
                            uiRet = Constants.DISP_E_NONAMEDARGS;
                            break;
                        }
                }
            }

            return uiRet;
        }

        public void Invoke
        (
            int dispId,
            ref Guid riid,
            int lcid,
            ushort wFlags,
            ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams,
            out object result,
            IntPtr pExcepInfo,
            IntPtr puArgErr
        )
        {
            result = null;

            switch (dispId)
            {
                // 1 is the dispid of TestMethod()
                case 1:
                    {
                        int i = 0;
                        result = null;

                        // We create an array of pointers to the VARIANTs inside the 
                        // DISPPARAMS.
                        IntPtr[] pVariantArray = new IntPtr[pDispParams.cArgs];

                        IntPtr pVariantTemp = pDispParams.rgvarg;
                        for (i = 0; i < pDispParams.cArgs; i++)
                        {
                            pVariantArray[i] = pVariantTemp;
                            pVariantTemp += Marshal.SizeOf(typeof(VariantStructGeneric));
                        }

                        // We access the reference to the SAFEARRAY in the first VARIANT and 
                        // create a managed array from it. Note that this is an array of UInt16.
                        UInt16[] uiArray = Marshal.GetObjectForNativeVariant<UInt16[]>(pVariantArray[0]);

                        char[] chArray = new char[uiArray.Length];

                        // We convert the UInt16 array to a Char array.
                        for (i = 0; i < uiArray.Length; i++)
                        {
                            chArray[i] = Convert.ToChar(uiArray[i]);
                        }

                        TestMethod(chArray);

                        break;
                    }
            }

            return;
        }

        #endregion IDispatch implementation

        [DispId(1)]
        public void TestMethod(char[] a)
        {
            if (a != null)
            {
                Console.WriteLine(a.GetType());

                for (int i = 0; i < a.Length; i++)
                {
                    Console.WriteLine(a[i]);
                }
            }
        }
    }
}
