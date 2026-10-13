using System;
using System.Globalization;
using System.Reflection;
using TAX.HTKK.Utilities.Core;

namespace HTKK;

public class DateUtils
{
	public DateTime IsDate(int Y, int m, int d)
	{
		try
		{
			return new DateTime(Y, m, d);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return new DateTime(1900, 1, 1);
		}
	}

	public int GetSoNgay(string fromDate, string toDate, string format)
	{
		try
		{
			DateTime dateTime = DateTime.ParseExact(fromDate, format, CultureInfo.CurrentCulture);
			DateTime dateTime2 = DateTime.ParseExact(toDate, format, CultureInfo.CurrentCulture);
			return (dateTime2 - dateTime).Days;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1;
		}
	}

	public static int GetSoThang(string fromMonth, string toMonth)
	{
		int result = 0;
		try
		{
			if (fromMonth.Length == 6 && toMonth.Length == 6)
			{
				int num = ConvertHelpers.ConvertToInt(fromMonth.Substring(0, 2));
				int num2 = ConvertHelpers.ConvertToInt(fromMonth.Substring(2));
				int num3 = ConvertHelpers.ConvertToInt(toMonth.Substring(0, 2));
				int num4 = ConvertHelpers.ConvertToInt(toMonth.Substring(2));
				result = ((num4 != num2) ? ((num4 - num2) * 12 + (num3 - num)) : (num3 - num));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = -999999999;
		}
		return result;
	}

	public DateTime? ToDate(string strDate, string strFormat)
	{
		try
		{
			DateTime? result = null;
			string[] array = strDate.Split(new char[1] { '/' });
			for (int i = 0; i < array.Length; i++)
			{
				array[i] = array[i].Trim();
			}
			int num = 0;
			int num2 = 0;
			int num3 = 0;
			switch (strFormat.Replace("/", "").ToUpper())
			{
			case "DDMMYYYY":
				if (array.Length == 3)
				{
					num = int.Parse(array[0]);
					num2 = int.Parse(array[1]);
					num3 = int.Parse(array[2]);
				}
				else if (array.Length == 1 && (strDate.Length == 6 || strDate.Length == 8))
				{
					num = int.Parse(strDate.Substring(1, 2));
					num2 = int.Parse(strDate.Substring(3, 2));
					num3 = int.Parse(strDate.Substring(5));
				}
				break;
			case "DDMM":
				if (array.Length == 2)
				{
					num = int.Parse(array[0]);
					num2 = int.Parse(array[1]);
					num3 = DateTime.Now.Year;
				}
				break;
			case "MMYYYY":
				if (array.Length == 2)
				{
					num = 1;
					num2 = int.Parse(array[0]);
					num3 = int.Parse(array[1]);
					result = IsDate(num3, num2, num);
				}
				else if (array.Length == 1 && strDate.Length == 6)
				{
					num = 1;
					num2 = int.Parse(strDate.Substring(1, 2));
					num3 = int.Parse(strDate.Substring(3));
				}
				break;
			case "YYYY":
				if (array.Length == 1)
				{
					num = 1;
					num2 = 1;
					num3 = int.Parse(array[0]);
				}
				break;
			case "DD":
				if (array.Length == 1)
				{
					num = int.Parse(array[0]);
					num2 = DateTime.Now.Month;
					num3 = DateTime.Now.Year;
				}
				break;
			case "MM":
				if (array.Length == 1)
				{
					num = 1;
					num2 = int.Parse(array[0]);
					num3 = DateTime.Now.Year;
				}
				break;
			}
			if (num != 0 && num2 != 0)
			{
				if (num3 < 100)
				{
					num3 += 2000;
				}
				result = IsDate(num3, num2, num);
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public string ToString(DateTime dDate, string strFormat)
	{
		try
		{
			string result = null;
			string text = dDate.Day.ToString();
			string text2 = dDate.Month.ToString();
			string text3 = dDate.Year.ToString();
			if (text.Length == 1)
			{
				text = "0" + text;
			}
			if (text2.Length == 1)
			{
				text2 = "0" + text2;
			}
			switch (strFormat.Replace("/", "").ToUpper())
			{
			case "DDMMYYYY":
				result = text + "/" + text2 + "/" + text3;
				break;
			case "DDMM":
				result = text + "/" + text2;
				break;
			case "MMYYYY":
				result = text2 + "/" + text3;
				break;
			case "YYYY":
				result = text3;
				break;
			case "DD":
				result = text;
				break;
			case "MM":
				result = text2;
				break;
			case "DD-MM-YYYY":
				result = text + "-" + text2 + "-" + text3;
				break;
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static int GetSoNgay(DateTime fromDate, DateTime toDate)
	{
		try
		{
			return (toDate - fromDate).Days;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1;
		}
	}

	public static int GetNamHienTai(int dNgayTaiChinh, int dThangTaiChinh)
	{
		try
		{
			DateTime toDate = new DateTime(DateTime.Now.Year, dThangTaiChinh, dNgayTaiChinh);
			int soNgay = GetSoNgay(DateTime.Now.Date, toDate);
			if (soNgay <= 0)
			{
				return DateTime.Now.Year;
			}
			return DateTime.Now.Year - 1;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
