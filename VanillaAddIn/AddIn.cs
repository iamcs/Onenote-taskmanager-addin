/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Extensibility;
using Microsoft.Office.Core;
using MyApplication.VanillaAddIn.Utilities;
using Application = Microsoft.Office.Interop.OneNote.Application;  // Conflicts with System.Windows.Forms
using Microsoft.Office.Interop.OneNote;
using System.Xml;

#pragma warning disable CS3003 // Type is not CLS-compliant

namespace MyApplication.VanillaAddIn
{
	[ComVisible(true)]
	[Guid("D5ECCD00-CF2D-409B-B65A-BDBACB9F21DB"), ProgId("MyApplication.VanillaAddIn")]

	public class AddIn : IDTExtensibility2, IRibbonExtensibility
	{
		protected Application OneNoteApplication
		{ get; set; }

		private chosetime mainForm;
        private MainForm form;
        private string newtimeline;
        public const string  VER = "1.0.0.0";
        
        public AddIn()
		{
		}


        /// <summary>
        /// Returns the XML in Ribbon.xml so OneNote knows how to render our ribbon
        /// </summary>
        /// <param name="RibbonID"></param>
        /// <returns></returns>
        public string GetCustomUI(string RibbonID)
		{
			return Properties.Resources.ribbon;
		}

		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>
		/// Cleanup
		/// </summary>
		/// <param name="custom"></param>
		public void OnBeginShutdown(ref Array custom)
		{
			this.mainForm?.Invoke(new Action(() =>
			{
				this.mainForm?.Close();
				this.mainForm = null;
			}));
		}

		/// <summary>
		/// Called upon startup.
		/// Keeps a reference to the current OneNote application object.
		/// </summary>
		/// <param name="application"></param>
		/// <param name="connectMode"></param>
		/// <param name="addInInst"></param>
		/// <param name="custom"></param>
		public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
		{
			SetOneNoteApplication((Application)Application);
		}

		public void SetOneNoteApplication(Application application)
		{
			OneNoteApplication = application;            
        }

		/// <summary>
		/// Cleanup
		/// </summary>
		/// <param name="RemoveMode"></param>
		/// <param name="custom"></param>
		[SuppressMessage("Microsoft.Reliability", "CA2001:AvoidCallingProblematicMethods", MessageId = "System.GC.Collect")]
		public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
		{
			OneNoteApplication = null;
			GC.Collect();
			GC.WaitForPendingFinalizers();
		}

        public void OnStartupComplete(ref Array custom)
        {
        }
              
		public async Task refreshoverview(IRibbonControl control)
		{
            //var qfil = this.OneNoteApplication.QuickFiling();
            //qfil.Run(new Callback());
            //qfil = null;
            //var thread = new Thread(() =>
            //{
            //var qfil = this.OneNoteApplication.QuickFiling();
            //qfil.Run(new Callback());
            //qfil = null;
            //wh.WaitOne();
            //});
            //thread.Start();
            OperateOnenote op = new OperateOnenote(OneNoteApplication);
            op.createtodopage();
            return;
            //SetOneNoteApplication.
                    }

        public async Task newtask(IRibbonControl control)
        {
            //var qfil = this.OneNoteApplication.QuickFiling();
            //qfil.Run(new Callback());
            //qfil = null;
            //OperateOnenote op = new OperateOnenote(OneNoteApplication);
            //op.test();
            mainForm = new chosetime(OneNoteApplication);
            var thread = new Thread(() =>
            {
                IntPtr myWindowHandle = new IntPtr((long)this.OneNoteApplication.Windows.CurrentWindow.WindowHandle);
                NativeWindow nativeWindow = new NativeWindow();
                nativeWindow.AssignHandle(myWindowHandle);
                mainForm.ShowDialog(nativeWindow);                
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return;
        }
        public async Task sortcurrentpage(IRibbonControl control)
        {

            OperateOnenote op = new OperateOnenote(OneNoteApplication);
            op.sortcurrentpage();     
            return;
        }

        public async Task about(IRibbonControl control)
        {
            string aboutline = "               任务管理插件\r\n\r\n          by STAN CHENS\r\n                VER:" + VER+ "\r\n         all rights reserved\r\n             保留所有权利           ";
            IntPtr myWindowHandle = new IntPtr((long)this.OneNoteApplication.Windows.CurrentWindow.WindowHandle);
            NativeWindow nativeWindow = new NativeWindow();
            nativeWindow.AssignHandle(myWindowHandle);
            MessageBox.Show(nativeWindow,aboutline);
            return;
        }
        public async Task navito(IRibbonControl control)
        {
            OperateOnenote op = new OperateOnenote(OneNoteApplication);
            OneNoteApplication.NavigateTo(op.getpageid("OVERVIEW"),"",false);
            return;
        }
        public async Task getxml(IRibbonControl control)
        {
            ShowForm();
            return;
        }
        class Callback : IQuickFilingDialogCallback
        {
            public Callback() { }
            public void OnDialogClosed(IQuickFilingDialog qfDialog)
            {
                //Console.WriteLine(qfDialog.SelectedItem);
                //Console.WriteLine(qfDialog.PressedButton);
                //Console.WriteLine(qfDialog.CheckboxState);
            }
        }


        private void ShowForm()
		{
            var thread = new Thread(() =>
            {
                var pageId = this.OneNoteApplication.Windows.CurrentWindow.CurrentPageId;                
                string pagexml;
                this.OneNoteApplication.GetPageContent(this.OneNoteApplication.Windows.CurrentWindow.CurrentPageId, out pagexml);               
                form = new MainForm(OneNoteApplication);
                IntPtr myWindowHandle = new IntPtr((long)this.OneNoteApplication.Windows.CurrentWindow.WindowHandle);
                NativeWindow nativeWindow = new NativeWindow();
                nativeWindow.AssignHandle(myWindowHandle);
                form.ShowDialog(nativeWindow);
            });
            thread.Start();
		}


        /// <summary>
        /// Specified in Ribbon.xml, this method returns the image to display on the ribbon button
        /// </summary>
        /// <param name="imageName"></param>
        /// <returns></returns>
        public IStream GetImage(string imageName)
		{
            MemoryStream imageStream = new MemoryStream();
            switch (imageName)
            {
                case "Logo.png":                    
                    Properties.Resources.Logo.Save(imageStream, ImageFormat.Png);
                    break;
                case "Logo1.png":
                    Properties.Resources.Logo1.Save(imageStream, ImageFormat.Png);
                    break;
                case "Logo2.png":
                    Properties.Resources.Logo2.Save(imageStream, ImageFormat.Png);
                    break;
                case "about.png":
                    Properties.Resources.about.Save(imageStream, ImageFormat.Png);
                    break;
                case "getxml.png":
                    Properties.Resources.getxml.Save(imageStream, ImageFormat.Png);
                    break;
                case "navito.png":
                    Properties.Resources.navito.Save(imageStream, ImageFormat.Png);
                    break;
                default: break;
            }
            return new CCOMStreamWrapper(imageStream);
		}


	}
}
