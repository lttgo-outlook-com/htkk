#define TRACE
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace HTKK;

public class Functions
{
	public int Quarter(DateTime dateTime)
	{
		try
		{
			return Convert.ToInt16((dateTime.Month - 1) / 3) + 1;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1;
		}
	}

	public void GetQuyHienTai(int dNgayTaiChinh, int dThangTaiChinh, ref int q, ref int Y)
	{
		try
		{
			DateTime dateTime = new DateTime(DateTime.Now.Year, 1, 1);
			DateTime dateTime2 = new DateTime(DateTime.Now.Year, dThangTaiChinh, dNgayTaiChinh);
			int days = (dateTime2 - dateTime).Days;
			DateTime dateTime3 = DateTime.Now.AddDays(days * -1);
			q = Quarter(dateTime3);
			Y = dateTime3.Year;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public int GetNamHienTai(int dNgayTaiChinh, int dThangTaiChinh)
	{
		try
		{
			DateTime dateTime = new DateTime(DateTime.Now.Year, dThangTaiChinh, dNgayTaiChinh);
			int days = (dateTime - DateTime.Now).Days;
			if (days <= 0)
			{
				return DateTime.Now.Year;
			}
			return DateTime.Now.Year - 1;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1;
		}
	}

	public static bool CheckServiceExists(string url, bool throwExceptions, out string errorMessage)
	{
		try
		{
			errorMessage = string.Empty;
			HttpWebRequest httpWebRequest = WebRequest.Create(url) as HttpWebRequest;
			httpWebRequest.Timeout = 20000;
			using HttpWebResponse httpWebResponse = httpWebRequest.GetResponse() as HttpWebResponse;
			if (httpWebResponse.StatusCode != HttpStatusCode.OK)
			{
				throw new Exception("Error locating web service");
			}
		}
		catch (WebException arg)
		{
			errorMessage = $"Error testing connection to web service at \"{url}\":\r\n{arg}";
			Trace.TraceError(errorMessage);
			return false;
		}
		return true;
	}

	public static bool IsEmail(string email)
	{
		try
		{
			string pattern = "^(([\\w-]+\\.)+[\\w-]+|([a-zA-Z]{1}|[\\w-]{2,}))@((([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\\.([0-1]?\r\n\t\t\t\t[0-9]{1,2}|25[0-5]|2[0-4][0-9])\\.([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\\.([0-1]?\r\n\t\t\t\t[0-9]{1,2}|25[0-5]|2[0-4][0-9])){1}|([a-zA-Z0-9]+[\\w-]+\\.)+[a-zA-Z]{1}[a-zA-Z0-9-]{1,23})$";
			if (email != null)
			{
				return Regex.IsMatch(email, pattern);
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static void ResetInfo(Dictionary<string, int> dicSheet, HTKKPeriodInfo periodInfo, HTKKCMCToKhaiInfo toKhaiInfo)
	{
		dicSheet.Clear();
		dicSheet.Add("TTC-NNT", 1);
		periodInfo = new HTKKPeriodInfo();
		toKhaiInfo = new HTKKCMCToKhaiInfo();
	}

	public static Dictionary<string, string> getDictionary(string nameXML)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_001a: Expected O, but got Unknown
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		try
		{
			string text = "../InterfaceIni/" + nameXML;
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[1];
				string value2 = array[0];
				dictionary.Add(key, value2);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return dictionary;
	}

	public static string FomartKey(string str1, string str2)
	{
		return str1 + "___" + str2;
	}

	public static string[] DeserializeKey(string key)
	{
		string[] separator = new string[1] { "___" };
		return key.Split(separator, StringSplitOptions.RemoveEmptyEntries);
	}

	public static string FormatKey(List<string> listStr, string fomartKey)
	{
		string text = "";
		foreach (string item in listStr)
		{
			text = text + item + fomartKey;
		}
		return text;
	}

	public static string[] DeserializeKey(string text, string fomartKey)
	{
		string[] separator = new string[1] { fomartKey };
		return text.Split(separator, StringSplitOptions.RemoveEmptyEntries);
	}

	public static void InsertCombobox(SheetView sView, int row, int col, Dictionary<string, string> dic, bool deleteText = true)
	{
		try
		{
			ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
			comboBoxCellType.Items = dic.Values.ToArray();
			comboBoxCellType.ItemData = dic.Keys.ToArray();
			comboBoxCellType.AutoSearch = AutoSearch.MultipleCharacter;
			comboBoxCellType.Editable = true;
			comboBoxCellType.MaxDrop = 10;
			comboBoxCellType.ListWidth = 500;
			sView.Cells[row - 1, col - 1].CellType = comboBoxCellType;
			if (deleteText)
			{
				sView.Cells[row - 1, col - 1].Text = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string RemoveUnikey(string text)
	{
		return Regex.Replace(text, "[^\\u0000-\\u007F]+", string.Empty);
	}
}
