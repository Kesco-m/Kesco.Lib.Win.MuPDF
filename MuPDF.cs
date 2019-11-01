using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml.Serialization;

namespace Kesco.Lib.Win.MuPDFLib
{
    interface IApi
    {
        IntPtr CreateMuPDFClass();
        void DisposeMuPDFClass(IntPtr pTestClassObject);
        int LoadPdf(IntPtr pTestClassObject, string filename, string password, bool share);
        int LoadPdfFromStream(IntPtr pTestClassObject, byte[] buffer, int bufferSize, string password);
        int LoadPage(IntPtr pTestClassObject, int pageNumber);
		int RenderPage(IntPtr pTestClassObject, IntPtr hDC, Rectangle screenRect, int pageNo, float zoom, int rotation, int target);
        float GetWidth(IntPtr pTestClassObject);
        float GetHeight(IntPtr pTestClassObject);
        void SetAlphaBits(IntPtr pTestClassObject, int alphaBits);
        IntPtr GetBitmapData(IntPtr pTestClassObject, out int width, out int height, float dpix, float dpiy, int rotation, int colorspace, bool rotateLandscapePages, bool convertToLetter, out int nLength, int maxSize);
        void FreeRenderedPage(IntPtr pTestClassObject);

        void ShowAnnots(IntPtr pTestClassObject, bool showAnnots);
        //fz_text_char_Ex_t[] TextToClient(IntPtr pTestClassObject);//, IntPtr txtPt);
        List<mChar> TextToClient(IntPtr pTestClassObject);//, IntPtr txtPt);
        //int[] TextFromRectangle(IntPtr pTestClassObject, Rectangle rect);

        int TextLen(IntPtr pTestClassObject);

        int GetRotation(IntPtr pTestClassObject);


        void PrintDoc(IntPtr pTestClassObject, string printer_name, uint nFromPage, uint nToPage, uint copies, short orientation, bool isOrigin);
    }

    internal class ThirtyTwoBitApi : IApi
    {
        private const string MuDLL = "MuPDFLib-x86.dll";

        [DllImport(MuDLL, EntryPoint = "CreateMuPDFClass", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        static private extern IntPtr CreateMuPDFClass_EXT();

        public IntPtr CreateMuPDFClass()
        {
            return CreateMuPDFClass_EXT();
        }

        [DllImport(MuDLL, EntryPoint = "DisposeMuPDFClass", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        static private extern void DisposeMuPDFClass_EXT(IntPtr pTestClassObject);

        public void DisposeMuPDFClass(IntPtr pTestClassObject)
        {
            DisposeMuPDFClass_EXT(pTestClassObject);
        }

        [DllImport(MuDLL, EntryPoint = "CallGetBitmap", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern IntPtr CallGetBitmap_EXT(IntPtr pTestClassObject, out int width, out int height, float dpix, float dpiy, int rotation, int colorspace, bool rotateLandscapePages, bool convertToLetter, out int nLength, int maxSize);

        public IntPtr GetBitmapData(IntPtr pTestClassObject, out int width, out int height, float dpix, float dpiy, int rotation, int colorspace, bool rotateLandscapePages, bool convertToLetter, out int nLength, int maxSize)
        {
            return CallGetBitmap_EXT(pTestClassObject, out width, out height, dpix, dpiy, rotation, colorspace, rotateLandscapePages, convertToLetter, out nLength, maxSize);
        }

        [DllImport(MuDLL, EntryPoint = "CallLoadPdf", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Auto)]
        private static extern int CallLoadPdf_EXT(IntPtr pTestClassObject, string filename, string password, bool share);

        public int LoadPdf(IntPtr pTestClassObject, string filename, string password, bool share)
        {
            return CallLoadPdf_EXT(pTestClassObject, filename, password, share);
        }

		[DllImport(MuDLL, EntryPoint = "CallRenderPage", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Auto)]
		private static extern int CallRenderPage_EXT(IntPtr pTestClassObject, IntPtr hDC, Rectangle screenRect, int pageNo, float zoom, int rotation, int target);

		public int RenderPage(IntPtr pTestClassObject, IntPtr hDC, Rectangle screenRect, int pageNo, float zoom, int rotation, int target) 
		{
			return CallRenderPage_EXT(pTestClassObject, hDC, screenRect, pageNo, zoom, rotation, target);
		}

        [DllImport(MuDLL, EntryPoint = "CallLoadPdfFromStream", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Auto)]
        private static extern int CallLoadPdfFromStream_EXT(IntPtr pTestClassObject, byte[] buffer, int bufferSize, string password);

        public int LoadPdfFromStream(IntPtr pTestClassObject, byte[] buffer, int bufferSize, string password)
        {
            return CallLoadPdfFromStream_EXT(pTestClassObject, buffer, bufferSize, password);
        }

        [DllImport(MuDLL, EntryPoint = "CallLoadPage", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int CallLoadPage_EXT(IntPtr pTestClassObject, int pageNumber);

        public int LoadPage(IntPtr pTestClassObject, int pageNumber)
        {
            return CallLoadPage_EXT(pTestClassObject, pageNumber);
        }

        [DllImport(MuDLL, EntryPoint = "CallGetWidth", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern float CallGetWidth_EXT(IntPtr pTestClassObject);

        public float GetWidth(IntPtr pTestClassObject)
        {
            return CallGetWidth_EXT(pTestClassObject);
        }

        [DllImport(MuDLL, EntryPoint = "CallGetHeight", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern float CallGetHeight_EXT(IntPtr pTestClassObject);

        public float GetHeight(IntPtr pTestClassObject)
        {
            return CallGetHeight_EXT(pTestClassObject);
        }

        [DllImport(MuDLL, EntryPoint = "CallSetAlphaBits", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern void CallSetAlphaBits_EXT(IntPtr pTestClassObject, int alphaBits);

        public void SetAlphaBits(IntPtr pTestClassObject, int alphaBits)
        {
            CallSetAlphaBits_EXT(pTestClassObject, alphaBits);
        }

        [DllImport(MuDLL, EntryPoint = "CallFreeRenderedPage", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern void CallFreeRenderedPage_EXT(IntPtr pTestClassObject);

        public void FreeRenderedPage(IntPtr pTestClassObject)
        {
            CallFreeRenderedPage_EXT(pTestClassObject);
        }

        [DllImport(MuDLL, EntryPoint = "CallSetShowAnnots", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern void CallSetShowAnnots_EXT(IntPtr pTestClassObject, bool showAnnots);

        public void ShowAnnots(IntPtr pTestClassObject, bool showAnnots)
        {
            CallSetShowAnnots_EXT(pTestClassObject, showAnnots); 
        }

        [DllImport(MuDLL, EntryPoint = "CallTextToClient", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        //private static extern int CallTextToClient_EXT(IntPtr pTestClassObject, /*[In, Out]*/ ref IntPtr txtPt);
        //private static extern int CallTextToClient_EXT(IntPtr pTestClassObject, /*[In, Out]*/ FileStream txtPt);
        //private static extern FileStream CallTextToClient_EXT(IntPtr pTestClassObject, string txtPt);
        //private static extern IntPtr CallTextToClient_EXT(IntPtr pTestClassObject, IntPtr txtPt, ref uint rl);
        private static extern void CallTextToClient_EXT(IntPtr pTestClassObject, StringBuilder str, int sz);

        public List<mChar> TextToClient(IntPtr pTestClassObject)
        {
            List<mChar> ret = new List<mChar>();
            
            int sz = CallTextLen_EXT(pTestClassObject) * 128;
            if (sz > 0)
            {
                StringBuilder sb = new StringBuilder(sz);
                CallTextToClient_EXT(pTestClassObject, sb, sb.Capacity);

                //System.Windows.Forms.MessageBox.Show("[" + sb.ToString() + "]");

                ret = Xml2Object(sb.ToString()).chars;
            }
#region
            //using (FileStream file_stream = new FileStream(dstPath, FileMode.Create, FileAccess.Write, FileShare.None))
            //{
            //    file_stream.Op
            //    //IntPtr pPoint = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(FileStream)));
            //    //Marshal.StructureToPtr(file_stream, pPoint, false);
            //CallTextToClient_EXT(pTestClassObject, dstPath);
            //}
            //fz_text_char_Ex_t point = new fz_text_char_Ex_t();

            ////Получаем маршализованный указатель на структуру:
            //IntPtr pPoint = Marshal.AllocHGlobal(Marshal.SizeOf(point));
            //Marshal.StructureToPtr(point, pPoint, false);//true);

            ////… а затем указатель на указатель:
            //IntPtr ppPoint = Marshal.AllocHGlobal(Marshal.SizeOf(pPoint));
            //Marshal.StructureToPtr(pPoint, ppPoint, false);

            //int resultLength = CallTextToClient_EXT(pTestClassObject, ppPoint);

            //int elementSize = Marshal.SizeOf(typeof(IntPtr));
            //List<fz_text_char_Ex_t> ret = new List<fz_text_char_Ex_t>();

            //point = (fz_text_char_Ex_t)Marshal.PtrToStructure(pPoint, typeof(fz_text_char_Ex_t));

            //Marshal.FreeHGlobal(pPoint);
            //Marshal.FreeHGlobal(ppPoint);

            #region
            //int resultLength1 = CallTextLen_EXT(pTestClassObject);
            ////IntPtr txtPtr = Marshal.AllocHGlobal(resultLength1 * Marshal.SizeOf(typeof(fz_text_char_Ex_t)));
            //IntPtr txtPtr = Marshal.AllocCoTaskMem(resultLength1 * Marshal.SizeOf(typeof(IntPtr)));
            #endregion

            //////int structSize = Marshal.SizeOf(typeof(fz_text_char_t));
            #region
            //int elementSize = Marshal.SizeOf(typeof(IntPtr));
            #endregion

            //for (int i = 0; i < resultLength1; i++)
            //{
            //    fz_text_char_Ex_t point = new fz_text_char_Ex_t();
            //    IntPtr pPoint = Marshal.AllocCoTaskMem(Marshal.SizeOf(point));

            //    Marshal.StructureToPtr(point, pPoint, true);
            //    Marshal.WriteIntPtr(txtPtr, elementSize * i, pPoint);

            //    //Marshal.FreeCoTaskMem(pPoint);
            //}

            ////IntPtr ppPoint = Marshal.AllocHGlobal(Marshal.SizeOf(txtPtr));
            ////Marshal.StructureToPtr(txtPtr, ppPoint, false);
            #region
            //uint resultLength = 0;
            //txtPtr = CallTextToClient_EXT(pTestClassObject, txtPtr, ref resultLength);



            #endregion
            //uint resultLength = 0;
            ////IntPtr val = CallTextToClient_EXT(pTestClassObject, txtPt, ref resultLength);
            //IntPtr val = CallTextToClient_EXT(pTestClassObject, ref resultLength);

            ////int structSize = Marshal.SizeOf(typeof(fz_text_char_t));
            //int elementSize = Marshal.SizeOf(typeof(IntPtr));

            ////fz_text_char_Ex_t[] ret = new fz_text_char_Ex_t[resultLength];

            ////for (int i = 0; i < resultLength; i++)
            ////{
            ////    IntPtr data = Marshal.ReadIntPtr(val, elementSize * i);
            ////    fz_text_char_Ex_t point = (fz_text_char_Ex_t)Marshal.PtrToStructure(data, typeof(fz_text_char_Ex_t));

            ////    ret[i] = point;
            ////}


            //IntPtr data = Marshal.ReadIntPtr(txtPtr, 0);
            //while (ret.Count < resultLength)
            //{
            //    ret.Add((fz_text_char_Ex_t)Marshal.PtrToStructure(data, typeof(fz_text_char_Ex_t)));

            //    //Marshal.Release(data);
            //    Marshal.DestroyStructure(data, typeof(fz_text_char_Ex_t));

            //    data = Marshal.ReadIntPtr(txtPtr, elementSize * ret.Count);
            //}
            ////GC.KeepAlive(txtPtr);

            ////Marshal.Release(data);
            ////Marshal.Release(txtPtr);
            //Marshal.DestroyStructure(data, typeof(fz_text_char_Ex_t));
            //Marshal.DestroyStructure(txtPtr, typeof(fz_text_char_Ex_t));

            //Marshal.FreeCoTaskMem(txtPtr);


            //Marshal.FinalReleaseComObject(txtPtr);

#endregion
            GC.Collect();

            return ret;
        }
        private static Page Xml2Object(string xml)
        {
            // создаем reader
            StringReader reader = new StringReader(xml);

            // создаем XmlSerializer
            XmlSerializer dsr = new XmlSerializer(typeof(Page));

            // десериализуем
            return dsr.Deserialize(reader) as Page;
        }


        [DllImport(MuDLL, EntryPoint = "CallTextFromRectangle", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern IntPtr CallTextToClient_EXT(IntPtr pTestClassObject, int x, int y, int w, int h);

        public int[] TextFromRectangle(IntPtr pTestClassObject, Rectangle rect)
        {
            int[] ret = new int[10];
            IntPtr val = CallTextToClient_EXT(pTestClassObject, rect.X, rect.Y, rect.Width, rect.Height);
            int elementSize = Marshal.SizeOf(typeof(IntPtr));
            int i = 0;
            while (i < 10)
            {
                ret[i] = Marshal.ReadInt32(val, elementSize * i);
                i++;
            }

            Marshal.Release(val);

            return ret;
        }

        
        [DllImport(MuDLL, EntryPoint = "CallTextLen", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int CallTextLen_EXT(IntPtr pTestClassObject);

        public int TextLen(IntPtr pTestClassObject)
        {
            return CallTextLen_EXT(pTestClassObject); 
        }

        [DllImport(MuDLL, EntryPoint = "CallGetRotation", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int CallGetRotation_EXT(IntPtr pTestClassObject);

        public int GetRotation(IntPtr pTestClassObject)
        {
            return CallGetRotation_EXT(pTestClassObject); 
        }

        [DllImport(MuDLL, EntryPoint = "CallPrintDoc", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Auto)]
        private static extern void CallPrintDoc_EXT(IntPtr pTestClassObject, string printer_name, uint nFromPage, uint nToPage, uint copies, short orientation, bool isOrigin);

        public void PrintDoc(IntPtr pTestClassObject, string printer_name, uint nFromPage, uint nToPage, uint copies, short orientation, bool isOrigin)
        {
            CallPrintDoc_EXT(pTestClassObject, printer_name, nFromPage, nToPage, copies, orientation, isOrigin); 
        }
    }

    internal class SixtyFourBitApi : IApi
    {
        private const string MuDLL = "MuPDFLib-x64.dll";

        [DllImport(MuDLL, EntryPoint = "CreateMuPDFClass", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        static private extern IntPtr CreateMuPDFClass_EXT();

        public IntPtr CreateMuPDFClass()
        {
            return CreateMuPDFClass_EXT();
        }

        [DllImport(MuDLL, EntryPoint = "DisposeMuPDFClass", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        static private extern void DisposeMuPDFClass_EXT(IntPtr pTestClassObject);

        public void DisposeMuPDFClass(IntPtr pTestClassObject)
        {
            DisposeMuPDFClass_EXT(pTestClassObject);
        }

        [DllImport(MuDLL, EntryPoint = "CallGetBitmap", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern IntPtr CallGetBitmap_EXT(IntPtr pTestClassObject, out int width, out int height, float dpix, float dpiy, int rotation, int colorspace, bool rotateLandscapePages, bool convertToLetter, out int nLength, int maxSize);

        public IntPtr GetBitmapData(IntPtr pTestClassObject, out int width, out int height, float dpix, float dpiy, int rotation, int colorspace, bool rotateLandscapePages, bool convertToLetter, out int nLength, int maxSize)
        {
            return CallGetBitmap_EXT(pTestClassObject, out width, out height, dpix, dpiy, rotation, colorspace, rotateLandscapePages, convertToLetter, out nLength, maxSize);
        }

		[DllImport(MuDLL, EntryPoint = "CallRenderPage", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Auto)]
		private static extern int CallRenderPage_EXT(IntPtr pTestClassObject, IntPtr hDC, Rectangle screenRect, int pageNo, float zoom, int rotation, int target);

		public int RenderPage(IntPtr pTestClassObject, IntPtr hDC, Rectangle screenRect, int pageNo, float zoom, int rotation, int target) 
		{
			return CallRenderPage_EXT(pTestClassObject, hDC, screenRect, pageNo, zoom, rotation, target);
		}

        [DllImport(MuDLL, EntryPoint = "CallLoadPdf", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Auto)]
        private static extern int CallLoadPdf_EXT(IntPtr pTestClassObject, string filename, string password, bool share);
        
        public int LoadPdf(IntPtr pTestClassObject, string filename, string password, bool share)
        {
            return CallLoadPdf_EXT(pTestClassObject, filename, password, share);
        }

        [DllImport(MuDLL, EntryPoint = "CallLoadPdfFromStream", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Auto)]
        private static extern int CallLoadPdfFromStream_EXT(IntPtr pTestClassObject, byte[] buffer, int bufferSize, string password);

        public int LoadPdfFromStream(IntPtr pTestClassObject, byte[] buffer, int bufferSize, string password)
        {
            return CallLoadPdfFromStream_EXT(pTestClassObject, buffer, bufferSize, password);
        }

        [DllImport(MuDLL, EntryPoint = "CallLoadPage", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int CallLoadPage_EXT(IntPtr pTestClassObject, int pageNumber);

        public int LoadPage(IntPtr pTestClassObject, int pageNumber)
        {
            return CallLoadPage_EXT(pTestClassObject, pageNumber);
        }

        [DllImport(MuDLL, EntryPoint = "CallGetWidth", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern float CallGetWidth_EXT(IntPtr pTestClassObject);

        public float GetWidth(IntPtr pTestClassObject)
        {
            return CallGetWidth_EXT(pTestClassObject);
        }

        [DllImport(MuDLL, EntryPoint = "CallGetHeight", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern float CallGetHeight_EXT(IntPtr pTestClassObject);

        public float GetHeight(IntPtr pTestClassObject)
        {
            return CallGetHeight_EXT(pTestClassObject);
        }

        [DllImport(MuDLL, EntryPoint = "CallSetAlphaBits", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern void CallSetAlphaBits_EXT(IntPtr pTestClassObject, int alphaBits);

        public void SetAlphaBits(IntPtr pTestClassObject, int alphaBits)
        {
            CallSetAlphaBits_EXT(pTestClassObject, alphaBits);
        }

        [DllImport(MuDLL, EntryPoint = "CallFreeRenderedPage", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern void CallFreeRenderedPage_EXT(IntPtr pTestClassObject);

        public void FreeRenderedPage(IntPtr pTestClassObject)
        {
            CallFreeRenderedPage_EXT(pTestClassObject);
        }

        [DllImport(MuDLL, EntryPoint = "CallSetShowAnnots", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern void CallSetShowAnnots_EXT(IntPtr pTestClassObject, bool showAnnots);

        public void ShowAnnots(IntPtr pTestClassObject, bool showAnnots)
        {
            CallSetShowAnnots_EXT(pTestClassObject, showAnnots);
        }

        [DllImport(MuDLL, EntryPoint = "CallTextToClient", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern void CallTextToClient_EXT(IntPtr pTestClassObject, StringBuilder str, int sz);

        public List<mChar> TextToClient(IntPtr pTestClassObject)
        {
            //uint resultLength = 0;
            ////IntPtr val = CallTextToClient_EXT(pTestClassObject, txtPt, ref resultLength);
            //IntPtr val = CallTextToClient_EXT(pTestClassObject, ref resultLength);

            ////int structSize = Marshal.SizeOf(typeof(fz_text_char_t));
            //int elementSize = Marshal.SizeOf(typeof(IntPtr));

            //fz_text_char_Ex_t[] ret = new fz_text_char_Ex_t[resultLength];

            //for (int i = 0; i < resultLength; i++)
            //{
            //    IntPtr data = Marshal.ReadIntPtr(val, elementSize * i);
            //    fz_text_char_Ex_t point = (fz_text_char_Ex_t)Marshal.PtrToStructure(data, typeof(fz_text_char_Ex_t));

            //    ret[i] = point;
            //}

            ////Marshal.FreeHGlobal(val);

            ////return ret;
            return new List<mChar>();
        }

        [DllImport(MuDLL, EntryPoint = "CallTextLen", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int CallTextLen_EXT(IntPtr pTestClassObject);

        public int TextLen(IntPtr pTestClassObject)
        {
            return CallTextLen_EXT(pTestClassObject);
        }


        [DllImport(MuDLL, EntryPoint = "CallTextFromRectangle", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int[] CallTextToClient_EXT(IntPtr pTestClassObject, Rectangle rect);

        public int[] TextFromRectangle(IntPtr pTestClassObject, Rectangle rect)
        {
            return CallTextToClient_EXT(pTestClassObject, rect);
        }

        [DllImport(MuDLL, EntryPoint = "CallGetRotation", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int CallGetRotation_EXT(IntPtr pTestClassObject);

        public int GetRotation(IntPtr pTestClassObject)
        {
            return CallGetRotation_EXT(pTestClassObject);
        }

        [DllImport(MuDLL, EntryPoint = "CallPrintDoc", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Auto)]
        private static extern void CallPrintDoc_EXT(IntPtr pTestClassObject, string printer_name, uint nFromPage, uint nToPage, uint copies, short orientation, bool isOrigin);

        public void PrintDoc(IntPtr pTestClassObject, string printer_name, uint nFromPage, uint nToPage, uint copies, short orientation, bool isOrigin)
        {
            CallPrintDoc_EXT(pTestClassObject, printer_name, nFromPage, nToPage, copies, orientation, isOrigin);
        }
    }

    public class MuPDF : IDisposable
    {
        private IApi _Api;
        private IntPtr m_pNativeObject;
        private string _FileName;
        private byte[] _Image;
        private GCHandle _ImagePin;
        private string _PdfPassword;
        private bool _Share;
        private int _CurrentPage;
        private int _LoadType;
        private int _AliasBits;

        //private bool _showAnnots = true; 

        public int PageCount { get; set; }
        public int Page
        {
            get { return _CurrentPage; }
            set
            {
                _CurrentPage = _Api.LoadPage(this.m_pNativeObject, value);
            }
        }

        public double Width
        {
            get
            {
                if (_CurrentPage > 0)
                    return _Api.GetWidth(this.m_pNativeObject);
                else
                    return 0;
            }
        }

        public double Height
        {
            get
            {
                if (_CurrentPage > 0)
                    return _Api.GetHeight(this.m_pNativeObject);
                else
                    return 0;
            }
        }

        public int Rotation
        {
            get
            {
                if (_CurrentPage > 0)
                    return _Api.GetRotation(this.m_pNativeObject);
                else
                    return 0;
            }
        }

        public bool AntiAlias
        {
            get
            {
                if (_AliasBits > 0)
                    return true;
                else
                    return false;
            }
            set
            {
                if (value)
                {
                    if (_AliasBits != 8)
                    {
                        _AliasBits = 8;
                        _Api.SetAlphaBits(this.m_pNativeObject, 8);
                    }
                }
                else
                {
                    if (_AliasBits != 0)
                    {
                        _AliasBits = 0;
                        _Api.SetAlphaBits(this.m_pNativeObject, 0);
                    }
                }
            }
        }

        public short AlphaLevel
        {
            set
            {
                if (value > 6)
                {
                    _AliasBits = 8;
                }
                else if (value > 4)
                {
                    _AliasBits = 6;
                }
                else if (value > 2)
                {
                    _AliasBits = 4;
                }
                else if (value > 0)
                {
                    _AliasBits = 2;
                }
                else
                {
                    _AliasBits = 0;
                }

                _Api.SetAlphaBits(this.m_pNativeObject, _AliasBits);
            }
        }

        public bool ShowAnnots
        {
            set { _Api.ShowAnnots(this.m_pNativeObject, value); }
        }

        public MuPDF(byte[] image, string pdfPassword)
        {
            _LoadType = 1;
            _Image = image;
            _PdfPassword = pdfPassword;

            if (image == null)
                throw new ArgumentNullException();
            Initialize();
        }

        public MuPDF(string fileName, string pdfPassword, bool share)
        {
            _LoadType = 0;
            _FileName = fileName;
            _PdfPassword = pdfPassword;
            _Share = share;

            if (!File.Exists(_FileName))
                throw new FileNotFoundException("Cannot find file to open!", _FileName);
            Initialize();
        }

       // [System.Runtime.ExceptionServices.HandleProcessCorruptedStateExceptions]
        public void Initialize()
        {
            if (IntPtr.Size == 8)
                _Api = new SixtyFourBitApi();
            else
                _Api = new ThirtyTwoBitApi();


            this.m_pNativeObject = _Api.CreateMuPDFClass();
            if (_LoadType == 0)
                PageCount = _Api.LoadPdf(this.m_pNativeObject, _FileName, _PdfPassword, _Share);
            else if (_LoadType == 1)
            {
                _ImagePin = GCHandle.Alloc(_Image, GCHandleType.Pinned);
                PageCount = _Api.LoadPdfFromStream(this.m_pNativeObject, _Image, (int)_Image.Length, _PdfPassword);
            }

            if (PageCount == -5)
                throw new Exception("PDF password needed!");
            else if (PageCount == -6)
                throw new Exception("Invalid PDF password supplied!");
            else if (PageCount < 1)
                throw new Exception("Unable to open pdf document!");
            _CurrentPage = 1;

            //_AliasBits = 0;
            //_Api.SetAlphaBits(this.m_pNativeObject, 0);
        }

        public void Dispose()
        {
            Dispose(true);
        }

        //[System.Runtime.ExceptionServices.HandleProcessCorruptedStateExceptions]
        protected virtual void Dispose(bool bDisposing)
        {
            if (this.m_pNativeObject != IntPtr.Zero)
            {
                // Call the DLL Export to dispose this class
                _Api.DisposeMuPDFClass(this.m_pNativeObject);
                this.m_pNativeObject = IntPtr.Zero;
                if (_ImagePin.IsAllocated)
                    _ImagePin.Free();
            }

            if (bDisposing)
            {
                // No need to call the finalizer since we've now cleaned
                // up the unmanaged memory
                GC.SuppressFinalize(this);
            }
        }

        ~MuPDF()
        {
            Dispose(false);
        }


		public void RenderPage(IntPtr hDC, Rectangle screenRect, int pageNo, float zoom, int rotation, int target)
		{
			_Api.RenderPage(this.m_pNativeObject, hDC, screenRect, pageNo, zoom, rotation, target);
		}

        //[System.Runtime.ExceptionServices.HandleProcessCorruptedStateExceptions]
        public unsafe Bitmap GetBitmap(int width, int height, float dpix, float dpiy, int rotation, RenderType type, bool rotateLandscapePages, bool convertToLetter, int maxSize)
        {
            Bitmap bitmap2 = null;
            int nLength = 0;
            IntPtr data = _Api.GetBitmapData(this.m_pNativeObject, out width, out height, dpix, dpiy, rotation, (int)type, rotateLandscapePages, convertToLetter, out nLength, maxSize);
            if (data == null || data == IntPtr.Zero)
                throw new Exception("Unable to render pdf page to bitmap!");

            if (type == RenderType.RGB)
            {
                bitmap2 = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                BitmapData imageData = bitmap2.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.ReadWrite, bitmap2.PixelFormat);
                byte* ptrSrc = (byte*)data;
                byte* ptrDest = (byte*)imageData.Scan0;
                for (int y = 0; y < height; y++)
                {
                    byte* pl = ptrDest;
                    byte* sl = ptrSrc;
                    for (int x = 0; x < width; x++)
                    {
                        //Swap these here instead of in MuPDF because most pdf images will be rgb or cmyk.
                        //Since we are going through the pixels one by one anyway swap here to save a conversion from rgb to bgr.
                        try
                        {
                            pl[2] = sl[0]; //b-r
                            pl[1] = sl[1]; //g-g
                            pl[0] = sl[2]; //r-b
                            //pl[3] = sl[3]; //alpha

                            pl += 3;
                            sl += 4;
                        }
                        catch { }
                    }
                    ptrDest += imageData.Stride;
                    ptrSrc += width * 4;
                }
                bitmap2.UnlockBits(imageData);
            }
            else if (type == RenderType.Grayscale)
            {
                bitmap2 = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format8bppIndexed);
                ColorPalette palette = bitmap2.Palette;
                for (int i = 0; i < 256; ++i)
                    palette.Entries[i] = System.Drawing.Color.FromArgb(i, i, i);
                bitmap2.Palette = palette;
                BitmapData imageData = bitmap2.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.ReadWrite, bitmap2.PixelFormat);

                byte* ptrSrc = (byte*)data;
                byte* ptrDest = (byte*)imageData.Scan0;
                for (int y = 0; y < height; y++)
                {
                    byte* pl = ptrDest;
                    byte* sl = ptrSrc;
                    for (int x = 0; x < width; x++)
                    {
                        pl[0] = sl[0];
                        //pl[1] = sl[1]; //alpha
                        pl += 1;
                        sl += 2;
                    }
                    ptrDest += imageData.Stride;
                    ptrSrc += width * 2;
                }
                bitmap2.UnlockBits(imageData);
            }
            else//RenderType.Monochrome
            {
                //bitmap2 = new Bitmap(width, height, bmpstride, PixelFormat.Format1bppIndexed, data);//Doesn't free memory
                bitmap2 = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format1bppIndexed);
                ColorPalette palette = bitmap2.Palette;
                palette.Entries[0] = System.Drawing.Color.FromArgb(0, 0, 0);
                palette.Entries[1] = System.Drawing.Color.FromArgb(255, 255, 255);
                bitmap2.Palette = palette;
                BitmapData imageData = bitmap2.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.ReadWrite, bitmap2.PixelFormat);

                byte* ptrSrc = (byte*)data;
                byte* ptrDest = (byte*)imageData.Scan0;
                for (int i = 0; i < nLength; i++)
                {
                    ptrDest[i] = ptrSrc[i];
                }
                bitmap2.UnlockBits(imageData);
            }
            bitmap2.SetResolution(dpix, dpiy);
            _Api.FreeRenderedPage(this.m_pNativeObject);//Free unmanaged array

            //Bitmap bitmap1 = new Bitmap(bitmap2.Width, bitmap2.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            //Graphics gr = Graphics.FromImage(bitmap1);
            //gr.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            //gr.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            ////gr.TextContrast = ;
            //gr.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            //gr.DrawImage(bitmap2, 0, 0);
            //return bitmap1;


            return bitmap2;
        }

        public unsafe Bitmap[] GetBitmapMatrix(int width, int height, int stepx, int stepy, float dpix, float dpiy, int rotation, RenderType type, bool rotateLandscapePages, bool convertToLetter, int maxSize)
        {
            
            int nLength = 0;
            IntPtr data = _Api.GetBitmapData(this.m_pNativeObject, out width, out height, dpix, dpiy, rotation, (int)type, rotateLandscapePages, convertToLetter, out nLength, maxSize);
            if (data == null || data == IntPtr.Zero)
                throw new Exception("Unable to render pdf page to bitmap!");
            int count = (height + stepy - 1) / stepy;
            Bitmap[] bitmap2 = new Bitmap[count];
            if (type == RenderType.RGB)
            {
                int k = 0; 
                bitmap2[0] = new Bitmap(width, stepy, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                BitmapData imageData = bitmap2[0].LockBits(new Rectangle(0, 0, width, stepy), ImageLockMode.ReadWrite, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                byte* ptrSrc = (byte*)data;
                byte* ptrDest = (byte*)imageData.Scan0;
                for (int y = 0; y < height; y++)
                {
                    byte* pl = ptrDest;
                    byte* sl = ptrSrc;
                    for (int x = 0; x < width; x++)
                    {
                        //Swap these here instead of in MuPDF because most pdf images will be rgb or cmyk.
                        //Since we are going through the pixels one by one anyway swap here to save a conversion from rgb to bgr.
                        try
                        {
                            pl[2] = sl[0]; //b-r
                            pl[1] = sl[1]; //g-g
                            pl[0] = sl[2]; //r-b
                            //pl[3] = sl[3]; //alpha

                            pl += 3;
                            sl += 4;
                        }
                        catch { }
                    }
                    ptrDest += imageData.Stride;
                    ptrSrc += width * 4;
                    if (y > stepy * (k+1) - 2)
                    {
                        bitmap2[k].UnlockBits(imageData);
                        k++;
                        if(k == count - 1 && height % stepy > 0)
                        {
                            bitmap2[k] = new Bitmap(width, height % stepy, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                            imageData = bitmap2[k].LockBits(new Rectangle(0, 0, width, height % stepy), ImageLockMode.ReadWrite, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                        }
                        else
                        {
							if (k < count)
							{
								bitmap2[k] = new Bitmap(width, stepy, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
								imageData = bitmap2[k].LockBits(new Rectangle(0, 0, width, stepy), ImageLockMode.ReadWrite, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
							}
                        }
						if (k < count)
							ptrDest = (byte*)imageData.Scan0;
                    }

                }
				if(k < count)
                bitmap2[k].UnlockBits(imageData);
            }
            else
                throw (new NotImplementedException("No"));
            return bitmap2;
        }

        public unsafe Bitmap GetBitmapPart(int width, int height, Rectangle drawRect, float dpix, float dpiy, int rotation, RenderType type, bool rotateLandscapePages, bool convertToLetter, int maxSize)
        {
            Bitmap bitmap2 = null;
            int nLength = 0;
            IntPtr data = _Api.GetBitmapData(this.m_pNativeObject, out width, out height, dpix, dpiy, rotation, (int)type, rotateLandscapePages, convertToLetter, out nLength, maxSize);
            if (data == null || data == IntPtr.Zero)
                throw new Exception("Unable to render pdf page to bitmap!");
            if (type == RenderType.RGB)
            {
                bitmap2 = new Bitmap(drawRect.Width, drawRect.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                BitmapData imageData = bitmap2.LockBits(new Rectangle(0, 0, drawRect.Width, drawRect.Height), ImageLockMode.ReadWrite, bitmap2.PixelFormat);
                byte* ptrSrc = (byte*)data;
                byte* ptrDest = (byte*)imageData.Scan0;
                for (int y = 0; y < height; y++)
                {
                    byte* pl = ptrDest;
                    byte* sl = ptrSrc;
                    for (int x = 0; x < width; x++)
                    {
                        if( y >= drawRect.Top && y <= drawRect.Bottom && x>= drawRect.Left && x<= drawRect.Right)
                        //Swap these here instead of in MuPDF because most pdf images will be rgb or cmyk.
                        //Since we are going through the pixels one by one anyway swap here to save a conversion from rgb to bgr.
                        try
                        {
                            pl[2] = sl[0]; //b-r
                            pl[1] = sl[1]; //g-g
                            pl[0] = sl[2]; //r-b
                            //pl[3] = sl[3]; //alpha

                            pl += 3;
                            sl += 4;
                        }
                        catch { }
                        else
                            sl += 4;
                        if (x > drawRect.Right)
                            break;
                    }
                    ptrDest += imageData.Stride;
                    ptrSrc += width * 4;
                    if (y > drawRect.Bottom)
                        break;
                }
                bitmap2.UnlockBits(imageData);
            }
            else
                throw (new NotImplementedException("No"));
            return bitmap2;
        }

        //[System.Runtime.ExceptionServices.HandleProcessCorruptedStateExceptions]
        public unsafe BitmapSource GetBitmapSource(int width, int height, float dpix, float dpiy, int rotation, RenderType type, bool rotateLandscapePages, bool convertToLetter, int maxSize)
        {
            WriteableBitmap write = null;
            int nLength = 0;
            IntPtr data = _Api.GetBitmapData(this.m_pNativeObject, out width, out height, dpix, dpiy, rotation, (int)type, rotateLandscapePages, convertToLetter, out nLength, maxSize);
            if (data == null || data == IntPtr.Zero)
                throw new Exception("Unable to render pdf page to bitmap!");

            if (type == RenderType.RGB)
            {
                const int depth = 24;
                int bmpstride = ((width * depth + 31) & ~31) >> 3;

                write = new WriteableBitmap(width, height, (double)dpix, (double)dpiy, PixelFormats.Bgr24, null);
                Int32Rect rect = new Int32Rect(0, 0, width, height);
                write.Lock();

                byte* ptrSrc = (byte*)data;
                byte* ptrDest = (byte*)write.BackBuffer;
                for (int y = 0; y < height; y++)
                {
                    byte* pl = ptrDest;
                    byte* sl = ptrSrc;
                    for (int x = 0; x < width; x++)
                    {
                        //Swap these here instead of in MuPDF because most pdf images will be rgb or cmyk.
                        //Since we are going through the pixels one by one anyway swap here to save a conversion from rgb to bgr.
                        pl[2] = sl[0]; //b-r
                        pl[1] = sl[1]; //g-g
                        pl[0] = sl[2]; //r-b
                        //pl[3] = sl[3]; //alpha
                        pl += 3;
                        sl += 4;
                    }
                    ptrDest += bmpstride;
                    ptrSrc += width * 4;
                }
                write.AddDirtyRect(rect);
                write.Unlock();
            }
            else if (type == RenderType.Grayscale)
            {
                const int depth = 8;//(n * 8)
                int bmpstride = ((width * depth + 31) & ~31) >> 3;

                write = new WriteableBitmap(width, height, (double)dpix, (double)dpiy, PixelFormats.Gray8, BitmapPalettes.Gray256);
                Int32Rect rect = new Int32Rect(0, 0, width, height);
                write.Lock();

                byte* ptrSrc = (byte*)data;
                byte* ptrDest = (byte*)write.BackBuffer;
                for (int y = 0; y < height; y++)
                {
                    byte* pl = ptrDest;
                    byte* sl = ptrSrc;
                    for (int x = 0; x < width; x++)
                    {
                        pl[0] = sl[0]; //g
                        //pl[1] = sl[1]; //alpha
                        pl += 1;
                        sl += 2;
                    }
                    ptrDest += bmpstride;
                    ptrSrc += width * 2;
                }
                write.AddDirtyRect(rect);
                write.Unlock();
            }
            else//RenderType.Monochrome
            {
                write = new WriteableBitmap(width, height, (double)dpix, (double)dpiy, PixelFormats.BlackWhite, BitmapPalettes.BlackAndWhite);
                Int32Rect rect = new Int32Rect(0, 0, width, height);
                write.Lock();
                byte* ptrSrc = (byte*)data;
                byte* ptrDest = (byte*)write.BackBuffer;
                for (int i = 0; i < nLength; i++)
                {
                    ptrDest[i] = ptrSrc[i];
                }
                write.AddDirtyRect(rect);
                write.Unlock();
            }
            _Api.FreeRenderedPage(this.m_pNativeObject);//Free unmanaged array
            if (write.CanFreeze)
                write.Freeze();
            return write;
        }

        //[System.Runtime.ExceptionServices.HandleProcessCorruptedStateExceptions]
        public unsafe WriteableBitmap GetWriteableBitmap(int width, int height, float dpix, float dpiy, int rotation, RenderType type, bool rotateLandscapePages, bool convertToLetter, int maxSize)
        {
            WriteableBitmap write = null;
            int nLength = 0;
            IntPtr data = _Api.GetBitmapData(this.m_pNativeObject, out width, out height, dpix, dpiy, rotation, (int)type, rotateLandscapePages, convertToLetter, out nLength, maxSize);
            if (data == null || data == IntPtr.Zero)
                throw new Exception("Unable to render pdf page to bitmap!");

            if (type == RenderType.RGB)
            {
                const int depth = 24;
                int bmpstride = ((width * depth + 31) & ~31) >> 3;

                write = new WriteableBitmap(width, height, (double)dpix, (double)dpiy, PixelFormats.Bgr24, null);
                Int32Rect rect = new Int32Rect(0, 0, width, height);
                write.Lock();

                byte* ptrSrc = (byte*)data;
                byte* ptrDest = (byte*)write.BackBuffer;
                for (int y = 0; y < height; y++)
                {
                    byte* pl = ptrDest;
                    byte* sl = ptrSrc;
                    for (int x = 0; x < width; x++)
                    {
                        //Swap these here instead of in MuPDF because most pdf images will be rgb or cmyk.
                        //Since we are going through the pixels one by one anyway swap here to save a conversion from rgb to bgr.
                        pl[2] = sl[0]; //b-r
                        pl[1] = sl[1]; //g-g
                        pl[0] = sl[2]; //r-b
                        //pl[3] = sl[3]; //alpha
                        pl += 3;
                        sl += 4;
                    }
                    ptrDest += bmpstride;
                    ptrSrc += width * 4;
                }
                write.AddDirtyRect(rect);
                write.Unlock();
            }
            else if (type == RenderType.Grayscale)
            {
                const int depth = 8;//(n * 8)
                int bmpstride = ((width * depth + 31) & ~31) >> 3;

                write = new WriteableBitmap(width, height, (double)dpix, (double)dpiy, PixelFormats.Gray8, BitmapPalettes.Gray256);
                Int32Rect rect = new Int32Rect(0, 0, width, height);
                write.Lock();

                byte* ptrSrc = (byte*)data;
                byte* ptrDest = (byte*)write.BackBuffer;
                for (int y = 0; y < height; y++)
                {
                    byte* pl = ptrDest;
                    byte* sl = ptrSrc;
                    for (int x = 0; x < width; x++)
                    {
                        pl[0] = sl[0]; //g
                        //pl[1] = sl[1]; //alpha
                        pl += 1;
                        sl += 2;
                    }
                    ptrDest += bmpstride;
                    ptrSrc += width * 2;
                }
                write.AddDirtyRect(rect);
                write.Unlock();
            }
            else
            {
                write = new WriteableBitmap(width, height, (double)dpix, (double)dpiy, PixelFormats.BlackWhite, BitmapPalettes.BlackAndWhite);
                Int32Rect rect = new Int32Rect(0, 0, width, height);
                write.Lock();
                byte* ptrSrc = (byte*)data;
                byte* ptrDest = (byte*)write.BackBuffer;
                for (int i = 0; i < nLength; i++)
                {
                    ptrDest[i] = ptrSrc[i];
                }
                write.AddDirtyRect(rect);
                write.Unlock();
            }
            _Api.FreeRenderedPage(this.m_pNativeObject);//Free unmanaged array
            if (write.CanFreeze)
                write.Freeze();
            return write;
        }

        //[System.Runtime.ExceptionServices.HandleProcessCorruptedStateExceptions]
        public unsafe byte[] GetPixels(ref int width, ref int height, float dpix, float dpiy, int rotation, RenderType type, bool rotateLandscapePages, bool convertToLetter, out uint cbStride, int maxSize)
        {
            byte[] output = null;
            int nLength = 0;
            IntPtr data = _Api.GetBitmapData(this.m_pNativeObject, out width, out height, dpix, dpiy, rotation, (int)type, rotateLandscapePages, convertToLetter, out nLength, maxSize);
            if (data == null || data == IntPtr.Zero)
                throw new Exception("Unable to render pdf page to bitmap!");

            if (type == RenderType.RGB)
            {
                const int depth = 24;
                int bmpstride = ((width * depth + 31) & ~31) >> 3;
                int newSize = bmpstride * height;

                output = new byte[newSize];
                cbStride = (uint)bmpstride;
                IntPtr DestPointer = Marshal.UnsafeAddrOfPinnedArrayElement(output, 0);

                byte* ptrSrc = (byte*)data;
                byte* ptrDest = (byte*)DestPointer;
                for (int y = 0; y < height; y++)
                {
                    byte* pl = ptrDest;
                    byte* sl = ptrSrc;
                    for (int x = 0; x < width; x++)
                    {
                        //Swap these here instead of in MuPDF because most pdf images will be rgb or cmyk.
                        //Since we are going through the pixels one by one anyway swap here to save a conversion from rgb to bgr.
                        pl[2] = sl[0]; //b-r
                        pl[1] = sl[1]; //g-g
                        pl[0] = sl[2]; //r-b
                        //pl[3] = sl[3]; //alpha
                        pl += 3;
                        sl += 4;
                    }
                    ptrDest += cbStride;
                    ptrSrc += width * 4;
                }
            }
            else if (type == RenderType.Grayscale)
            {
                const int depth = 8;//(n * 8)
                int bmpstride = ((width * depth + 31) & ~31) >> 3;
                int newSize = bmpstride * height;

                output = new byte[newSize];
                cbStride = (uint)bmpstride;
                IntPtr DestPointer = Marshal.UnsafeAddrOfPinnedArrayElement(output, 0);

                byte* ptrSrc = (byte*)data;
                byte* ptrDest = (byte*)DestPointer;
                for (int y = 0; y < height; y++)
                {
                    byte* pl = ptrDest;
                    byte* sl = ptrSrc;
                    for (int x = 0; x < width; x++)
                    {
                        pl[0] = sl[0]; //g
                        //pl[1] = sl[1]; //alpha
                        pl += 1;
                        sl += 2;
                    }
                    ptrDest += cbStride;
                    ptrSrc += width * 2;
                }
            }
            else//RenderType.Monochrome
            {
                const int depth = 1;
                int bmpstride = ((width * depth + 31) & ~31) >> 3;

                cbStride = (uint)bmpstride;
                output = new byte[nLength];
                Marshal.Copy(data, output, 0, nLength);
            }
            _Api.FreeRenderedPage(this.m_pNativeObject);//Free unmanaged array
            return output;
        }

        public PdfText GetText()
        {
            List<mChar> buff = _Api.TextToClient(this.m_pNativeObject);

            try
            {
                if (buff == null)// || buff == IntPtr.Zero)
                    throw new Exception("Unable to get pdf page text!");

                return new PdfText(buff);//, txtPtr);
            }
            finally
            {
                if (buff != null)
                {
                    buff.Clear();
                    buff = null;
                }
            }
        }

        public void PrintDoc(string printer_name, uint nFromPage, uint nToPage, uint copies, short orientation, bool isOrigin)
        {
            _Api.PrintDoc(this.m_pNativeObject, printer_name, nFromPage, nToPage, copies, orientation, isOrigin);
        }
    }

    public enum RenderType
    {
        /// <summary>24-bit Color RGB</summary>
        RGB = 0,
        /// <summary>8-bit Grayscale</summary>
        Grayscale = 1,
        /// <summary>1-bit Monochrome</summary>
        Monochrome = 2
    }

    //[StructLayout(LayoutKind.Sequential, Pack = 4)]
    //internal struct fz_text_char_Ex_t
    //{
    //    public fz_text_char_t text_char;
    //    public int rNum;
    //};

    //[StructLayout(LayoutKind.Sequential, Pack = 4)]
    //internal struct fz_text_char_t
    //{
    //    public fz_rect_t bbox;
    //    public int c;
    //};

    //[StructLayout(LayoutKind.Sequential, Pack = 4)]
    //internal struct fz_rect_t
    //{
    //    public float x0, y0;
    //    public float x1, y1;
    //};

    internal class Char_Point
    {
        public Rectangle Location { get; private set; }
        public String Symbol { get; private set; }
        public int LineNumber { get; private set; }

        private int codeOrigin;

        internal Char_Point(int ch, float x0, float y0,  float x1, float y1, int ln)
        {
            Location = new Rectangle((int)x0, (int)y0, (int)(x1 - x0), (int)(y1 - y0));
            codeOrigin = ch;
            Symbol = Encoding.Unicode.GetString(BitConverter.GetBytes((short)ch));
            LineNumber = ln;
        }
    }

    public class PdfText : IDisposable
    {
        private List<Char_Point> fz_text;// = new List<Char_Point>();

        internal PdfText(List<mChar> src)
        {
            fz_text = src.Select(x => new Char_Point(x.c, x.x0, x.y0, x.x1, x.y1, x.line)).ToList();
        }

        public string ExtractText(Rectangle roundRect)
        {
            //string ret = "";

            //List<fz_text_char_Ex_t> tmp = fz_text.Where(x => (((x.text_char.bbox.x0 >= roundRect.X && x.text_char.bbox.x0 <= roundRect.Right) || (x.text_char.bbox.x1 >= roundRect.X && x.text_char.bbox.x1 <= roundRect.Right))
            //    && ((x.text_char.bbox.y0 >= roundRect.Y && x.text_char.bbox.y0 <= roundRect.Bottom ) || (x.text_char.bbox.y1 >= roundRect.Y && x.text_char.bbox.y1 <= roundRect.Bottom)))
            ////    || (x.text_char.bbox.x0 >= roundRect.X && x.text_char.bbox.x1 >= roundRect.Right)
            //).ToList();

            List<Char_Point> tmp = fz_text.Where(x => roundRect.IntersectsWith(x.Location)).ToList();

            for (int i = 0; i < tmp.Count; i++)
            {
                if (i + 1 < tmp.Count && tmp[i].LineNumber != tmp[i + 1].LineNumber)
                {
                    tmp.Insert(i + 1, new Char_Point(0x0d, 0, 0, 0, 0, 0));
                    tmp.Insert(i + 2, new Char_Point(0x0a, 0, 0, 0, 0, 0));
                    i += 2;
                }
            }
            
            return string.Concat(tmp.Select(x => x.Symbol).ToArray());
        }

        //public void FreeText()
        //{
        //    Marshal.FreeHGlobal(textPtr);
        //}

         public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool bDisposing)
        {
            if (fz_text != null)
            {
                fz_text.Clear();
                fz_text = null;
            }

            if (bDisposing)
            {
                GC.SuppressFinalize(this);
            }
        }

        ~PdfText()
        {
            Dispose(false);
        }
    }

    //<page>
    //  <chars> 
    //      <char line="1" x0="106.319" y0="127.735" x1="111.637" y1="139.488" c="1"/>
    //  </chars> 
    //</page>

    [XmlRoot("page")]
    public class Page
    {
        [XmlArray("chars")]
        [XmlArrayItem("char", typeof(mChar))]
        public List<mChar> chars = new List<mChar>();
    }

    public class mChar
    {
        [XmlAttribute] public int line;
        [XmlAttribute] public float x0;
        [XmlAttribute] public float x1;
        [XmlAttribute] public float y0;
        [XmlAttribute] public float y1;
        [XmlAttribute] public int c;
    }

}
