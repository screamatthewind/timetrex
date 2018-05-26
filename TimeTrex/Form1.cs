using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Net;
using System.Net.Mail;
using System.Drawing.Imaging;
using System.Net.Mime;
using System.IO;
using System.Runtime.InteropServices;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        int navNum;
        String expectInOut;
        Boolean applicationError = false;
        String errorFunction;

        String imagePath = @"C:\Users\Bob\Downloads\TimeTrex_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".jpg";

        public Form1()
        {
            InitializeComponent();
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            timer1.Stop();
            timer1.Start();
        }

        private void operationFailed(String funcName)
        {
            errorFunction = funcName;

            timer1.Stop();
  
            label1.Text = "Error trying to get: " + funcName;

            if (navNum >= 9)
            {
                sendMail("Error logging " + expectInOut, funcName + " Unable to log out", imagePath);
                Application.Exit();
                Environment.Exit(1);
                return;
            }
            else
            {
                applicationError = true;

                captureScreenShot(imagePath);

                // attempt to log out
                navNum = 9;
                timer1.Start();
            }
        }

        private void login()
        {
            HtmlElement el;
            
            el = webBrowser1.Document.GetElementById("user_name");
            if (el != null)
                el.InnerText = "username";
            else
                operationFailed("user_name");

            el = webBrowser1.Document.GetElementById("password");
            if (el != null)
                el.InnerText = "password";
            else
                operationFailed("password");

            el = webBrowser1.Document.GetElementById("login_btn");
            if (el != null)
                el.InvokeMember("click");
            else
                operationFailed("login_btn");

            navNum++;
        }

        private String validateInOut()
        {
            HtmlElementCollection els;

            els = webBrowser1.Document.GetElementsByTagName("select");

            foreach (HtmlElement el in els)
            {
                if (el.InnerText == "InOut")
                {
                    object objElement = el.DomElement;

                    object objSelectedIndex = objElement.GetType().InvokeMember("selectedIndex", BindingFlags.GetProperty, null, objElement, null);
                    int selectedIndex = (int)objSelectedIndex;
                    if (selectedIndex != -1)
                        return (el.Children[selectedIndex].InnerText.ToLower());
                }
            }

            operationFailed("validateInOut");

            return "error";
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            HtmlElement el;

            label1.Text += ".";
            if (webBrowser1.ReadyState != WebBrowserReadyState.Complete) return;

            timer1.Stop();

            switch (navNum)
            {
                case 1:
                    Application.DoEvents();

                    navNum++;
                    timer1.Start();

                    break;

                case 2:
                    label1.Text = "Logging In";

                    el = webBrowser1.Document.GetElementById("user_name");
                    if (el == null)
                    {
                        timer1.Start();
                        return;
                    }

                    login();

                    break;

                case 3: 
                    Application.DoEvents();

                    navNum++;
                    timer1.Start();

                    break;


                case 4:
                    label1.Text = "Goto In/Out Page";

                    el = webBrowser1.Document.GetElementById("inOutIcon");
                    if (el == null)
                    {
                        timer1.Start();
                        return;
                    }

                    el.InvokeMember("click");

                    navNum++;

                    break;

                case 5:
                    Application.DoEvents();

                    navNum++;
                    timer1.Start();

                    break;

                case 6:
                    label1.Text = "Click Save";

                    el = webBrowser1.Document.GetElementById("saveIcon");
                    if (el == null)
                    {
                        timer1.Start();
                        return;
                    }

                    HtmlElementCollection els;

                    els = webBrowser1.Document.GetElementsByTagName("select");

                    foreach (HtmlElement sel in els)
                    {
                        if (sel.InnerText == "InOut")
                        {

                            if (expectInOut.ToLower() == validateInOut())
                            {
                                el = webBrowser1.Document.GetElementById("saveIcon");
                                if (el == null) return;

                                navNum++;

                                el.InvokeMember("click");

                                break;
                            }
                            else
                                operationFailed("Already logged " + expectInOut);

                            break;
                        }
                    }

                    break;

                case 7:
                    Application.DoEvents();

                    navNum++;
                    timer1.Start();

                    break;

                case 8:
                    el = webBrowser1.Document.GetElementById("inOutIcon");
                    if (el == null)
                    {
                        timer1.Start();
                        return;
                    }

                    label1.Text = "Updated In/Out";

                    navNum++;
                    timer1.Start();

                    break;

                case 9: // begin logging out by clicking "My Account" link

                    label1.Text = "Attempting to log out";

                    foreach (HtmlElement elm in webBrowser1.Document.GetElementsByTagName("A"))
                    {
                        if (elm.OuterText == "My Account")
                        {
                            navNum++;
                            timer1.Start();

                            elm.InvokeMember("click");

                            return;
                        }
                    }

                    operationFailed("My Account link was not found");

                    break;

                case 10:
                    Application.DoEvents();

                    navNum++;
                    timer1.Start();

                    break;

                case 11:
                    el = webBrowser1.Document.GetElementById("Logout");
                    if (el == null)
                    {
                        timer1.Start();
                        return;
                    }

                    label1.Text = "Logging off";

                    navNum++;
                    el.InvokeMember("click");

                    break;

                case 12:
                    Application.DoEvents();

                    navNum++;
                    timer1.Start();

                    break;

                case 13:
                    el = webBrowser1.Document.GetElementById("user_name");
                    if (el == null)
                    {
                        timer1.Start();
                        return;
                    }

                    if (applicationError)
                        sendMail("Error logging " + expectInOut, "Error: " + errorFunction + ": Successfully logged off", imagePath);
                    else
                        sendMail("Logged " + expectInOut, "Success", null);

                    Application.Exit();
                    Environment.Exit(1);
                    return;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string[] args = Environment.GetCommandLineArgs();

            if (args.Count() != 2)
            {
                sendMail("Error: wrong number of parameters passed", "Cannot start application", null);
                Application.Exit();
                Environment.Exit(1);
                return;
            }

            expectInOut = args[1];

            if (!(expectInOut.ToLower().Equals("in") || expectInOut.ToLower().Equals("out")))
            {
                sendMail("Error: wrong parameters passed", "Received parameter: " + expectInOut, null);
                Application.Exit();
                Environment.Exit(1);
                return;

            }


            navNum = 1;
            webBrowser1.Navigate("http://timetrex.screamatthewind.com/timetrex/interface/html5/");
        }

        private void sendMail(String subject, String body, String theImagePath)
        {
            var fromAddress = new MailAddress("username@gmail.com", "Bob");
            var toAddress = new MailAddress("username@gmail.com", "Bob");
            const string fromPassword = "password";

            NetworkCredential creds = new NetworkCredential(fromAddress.Address, fromPassword);

            var smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = creds
            };

            Attachment inline = null;

            if (theImagePath != null)
            {

                String attachmentPath = theImagePath;
                inline = new Attachment(attachmentPath);

                // string attachmentPath = Environment.CurrentDirectory + imagePath;
                inline.ContentDisposition.Inline = true;
                inline.ContentDisposition.DispositionType = DispositionTypeNames.Inline;
                inline.ContentId = "TimeTrex Screenshot";
                inline.ContentType.MediaType = "image/jpg";
                inline.ContentType.Name = Path.GetFileName(attachmentPath);
            }

            using (var message = new MailMessage(fromAddress, toAddress)
            {
                Subject = "TimeTrex: " + subject,
                Body = body
            })
            {
                if (theImagePath != null)
                    message.Attachments.Add(inline);

                smtp.Send(message);
            }
        }

        private void captureScreenShot(String imagePath)
        {
            if (webBrowser1.ActiveXInstance == null) { 
                imagePath = null;
                return;
            }

            NativeMethods nm = new NativeMethods();

            Bitmap screenshot = new Bitmap(1024, 768);
            NativeMethods.GetImage(webBrowser1.ActiveXInstance, screenshot, Color.White);

            screenshot.Save(imagePath, ImageFormat.Jpeg);
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            timer2.Stop();

            captureScreenShot(imagePath);

            sendMail("Operation Timed Out", "Took more than 15 minutes to complete.  NavNum = " + navNum.ToString() + " Status: "+  label1.Text, imagePath);
            Application.Exit();
            Environment.Exit(1);
            return;
        }
    }
}

class NativeMethods
{
    [ComImport]
    [Guid("0000010D-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IViewObject
    {
        void Draw([MarshalAs(UnmanagedType.U4)] uint dwAspect, int lindex, IntPtr pvAspect, [In] IntPtr ptd, IntPtr hdcTargetDev, IntPtr hdcDraw, [MarshalAs(UnmanagedType.Struct)] ref RECT lprcBounds, [In] IntPtr lprcWBounds, IntPtr pfnContinue, [MarshalAs(UnmanagedType.U4)] uint dwContinue);
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public static void GetImage(object obj, Image destination, Color backgroundColor)
    {
        using (Graphics graphics = Graphics.FromImage(destination))
        {
            IntPtr deviceContextHandle = IntPtr.Zero;
            RECT rectangle = new RECT();

            rectangle.Right = destination.Width;
            rectangle.Bottom = destination.Height;

            graphics.Clear(backgroundColor);

            try
            {
                deviceContextHandle = graphics.GetHdc();

                IViewObject viewObject = obj as IViewObject;
                viewObject.Draw(1, -1, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, deviceContextHandle, ref rectangle, IntPtr.Zero, IntPtr.Zero, 0);
            }
            finally
            {
                if (deviceContextHandle != IntPtr.Zero)
                {
                    graphics.ReleaseHdc(deviceContextHandle);
                }
            }
        }
    }
}