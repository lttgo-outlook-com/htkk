using System;
using System.Windows.Forms;
using TAX.HTKK.Utilities.Core;

namespace TAX.HTKK.Business;

public class HelpUtil
{
	public static string HelpPath { get; set; }

	public static string HtmFolder { get; set; }

	public static void Show(Form frmParent, string helpId)
	{
		try
		{
			if (HTKKFormater.isNullString(helpId) || helpId.Length <= 3)
			{
				helpId = "outline_0";
			}
			string text = (string.IsNullOrEmpty(HtmFolder) ? (helpId + ".htm") : $"{HtmFolder}\\{helpId}.htm");
			Help.ShowHelp((Control)(object)frmParent, $"file://{HelpPath}", text);
		}
		catch (Exception)
		{
		}
	}
}
