using DeckRobotAddin.Properties;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;


namespace DeckRobotAddin
{
	[ComVisible(true)]
	public class DeckRobotRibbon : Office.IRibbonExtensibility
	{
		private Office.IRibbonUI ribbon;

		public DeckRobotRibbon()
		{
		}

		public Bitmap GetImage(IRibbonControl control)
		{
			return Resources.dr_icon;
		}

		#region IRibbonExtensibility Members

		public string GetCustomUI(string ribbonID)
		{
			return GetResourceText("DeckRobotAddin.DeckRobotRibbon.xml");
		}

		#endregion

		#region Ribbon Callbacks
		public void OnButtonClick(IRibbonControl control)
		{
			switch (control.Id)
			{
				case "btn_hello_world":
					MessageBox.Show(Resources.HelloMessage, Resources.AddinTitle);
					break;
				case "btn_open_new":
					OpenNew();
					break;
			}
		}

		private void OpenNew()
		{
			var ppt = Globals.ThisAddIn.Application.Presentations.Add();
			var slide = ppt.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

			var shape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 300, 200);
			var textRange = shape.TextFrame.TextRange;
			textRange.Text = Resources.HelloMessage;
			textRange.Font.Color.RGB = 0x00CC3333;
			textRange.Font.Bold = Office.MsoTriState.msoTrue;
			textRange.Font.Size = 20;
		}

		public void Ribbon_Load(Office.IRibbonUI ribbonUI)
		{
			this.ribbon = ribbonUI;
		}

		#endregion

		#region Helpers

		private static string GetResourceText(string resourceName)
		{
			Assembly asm = Assembly.GetExecutingAssembly();
			string[] resourceNames = asm.GetManifestResourceNames();
			for (int i = 0; i < resourceNames.Length; ++i)
			{
				if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
				{
					using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
					{
						if (resourceReader != null)
						{
							return resourceReader.ReadToEnd();
						}
					}
				}
			}
			return null;
		}

		#endregion
	}
}
