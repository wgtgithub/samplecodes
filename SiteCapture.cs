using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices.ComTypes;

namespace Sampleresources
{
    public class SiteCapture
    {
        [DllImport("ole32.dll")]
        private static extern int OleDraw(IntPtr pUnk, int dwAspect, IntPtr hdcDraw, ref Rectangle lprcBounds);

        private static bool _isCaptured;
        private static WebBrowser _browser;
        private static Bitmap _bitmap;
        private static Graphics _graphic;
        private static IntPtr _ptrObj;
        private static IntPtr _ptrHdc;

        public static Bitmap getBitmapFromUrl(string url)
        {
            ObjectInitialize();
            Navigate(url);
            WaitUntilCaptured(20);

            // return got bitmap information
            return _bitmap;
        }

        private static void ObjectInitialize()
        {
            _isCaptured = false;
            _graphic = null;
            _ptrObj = IntPtr.Zero;
            _ptrHdc = IntPtr.Zero;
            _bitmap = null;

            _browser = new WebBrowser();
            _browser.ScrollBarsEnabled = false;
            _browser.ScriptErrorsSuppressed = true;
            _browser.DocumentCompleted += Browser_DocumentCompleted;
        }

        private static void Navigate(string url)
        {
            _browser.Navigate(url);
        }

        private static void WaitUntilCaptured(int times)
        {
            while (!_isCaptured)
            {
                // wait untile complete load and capture for the specified
                Application.DoEvents();
                System.Threading.Thread.Sleep(times);
            }

        }

        private static void Browser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            _bitmap = null;
            if (e.Url.Equals("about:blank"))
            {
                _isCaptured = true;
                return;
            }

            try
            {
                SetSizes();
                CreateBitmap();
            }
            finally
            {
                _isCaptured = true;
            }
        }

        private static void SetSizes()
        {
            // set capture size
            _browser.Width = _browser.Document.Body.ScrollRectangle.Width;
            _browser.Height = _browser.Document.Body.ScrollRectangle.Height;

            // create bitmap to save captured picture
            _bitmap = new Bitmap(_browser.Width, _browser.Height);
        }

        private static void CreateBitmap()
        {
            try
            {
                _graphic = Graphics.FromImage(_bitmap);
                _ptrHdc = _graphic.GetHdc();
                _ptrObj = Marshal.GetIUnknownForObject(_browser.ActiveXInstance);
                Rectangle rect = new Rectangle(0, 0, _browser.Width, _browser.Height);

                // ptrObj Paste the area specified by rect in the image to the HDC area
                OleDraw(_ptrObj, (int)DVASPECT.DVASPECT_CONTENT, _ptrHdc, ref rect);
            }
            finally
            {
                ClearObjects();
            }
        }

        private static void ClearObjects()
        {
            if (_ptrObj != IntPtr.Zero)
            {
                Marshal.Release(_ptrObj);
            }
            if (_ptrHdc != IntPtr.Zero)
            {
                _graphic.ReleaseHdc(_ptrHdc);
            }
            if (_graphic != null)
            {
                _graphic.Dispose();
            }
        }
    }
}

