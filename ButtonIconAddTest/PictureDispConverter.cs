﻿using System.Runtime.InteropServices;
using IPictureDisp = stdole.IPictureDisp;

public sealed class PictureDispConverter

{

    [DllImport("OleAut32.dll",

        EntryPoint = "OleCreatePictureIndirect",

        ExactSpelling = true,

        PreserveSig = false)]

    private static extern IPictureDisp

        OleCreatePictureIndirect(

            [MarshalAs(UnmanagedType.AsAny)] object picdesc,

            ref Guid iid,

            [MarshalAs(UnmanagedType.Bool)] bool fOwn);



    static Guid iPictureDispGuid = typeof(IPictureDisp).GUID;



    private static class PICTDESC

    {

        //Picture Types

        public const short PICTYPE_UNINITIALIZED = -1;

        public const short PICTYPE_NONE = 0;

        public const short PICTYPE_BITMAP = 1;

        public const short PICTYPE_METAFILE = 2;

        public const short PICTYPE_ICON = 3;

        public const short PICTYPE_ENHMETAFILE = 4;



        [StructLayout(LayoutKind.Sequential)]

        public class Icon

        {

            internal int cbSizeOfStruct =

                Marshal.SizeOf(typeof(PICTDESC.Icon));

            internal int picType = PICTDESC.PICTYPE_ICON;

            internal IntPtr hicon = IntPtr.Zero;

            internal int unused1;

            internal int unused2;



            internal Icon(System.Drawing.Icon icon)

            {

                this.hicon = icon.ToBitmap().GetHicon();

            }

        }



        [StructLayout(LayoutKind.Sequential)]

        public class Bitmap

        {

            internal int cbSizeOfStruct =

                Marshal.SizeOf(typeof(PICTDESC.Bitmap));

            internal int picType = PICTDESC.PICTYPE_BITMAP;

            internal IntPtr hbitmap = IntPtr.Zero;

            internal IntPtr hpal = IntPtr.Zero;

            internal int unused;



            internal Bitmap(System.Drawing.Bitmap bitmap)

            {

                this.hbitmap = bitmap.GetHbitmap();

            }

        }

    }

    public static IPictureDisp ToIPictureDisp(

        System.Drawing.Bitmap bmp)

    {

        PICTDESC.Bitmap pictBmp = new PICTDESC.Bitmap(bmp);



        return OleCreatePictureIndirect(pictBmp, ref iPictureDispGuid, true);

    }
}