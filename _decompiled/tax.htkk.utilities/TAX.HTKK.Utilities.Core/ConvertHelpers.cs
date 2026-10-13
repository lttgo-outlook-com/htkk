using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using ConvertDB;
using FarPoint.Win.Spread;

namespace TAX.HTKK.Utilities.Core;

public static class ConvertHelpers
{
	public static Dictionary<string, CellTypeKHBS> _dicCellType_01_1_KHBS;

	public static DateTime ConvertToDateTime(object obj, string format)
	{
		DateTime result = default(DateTime);
		string empty = string.Empty;
		if (obj != null && obj != DBNull.Value)
		{
			empty = obj.ToString();
			DateTime.TryParseExact(empty, format, CultureInfo.CurrentCulture, DateTimeStyles.None, out result);
			return result;
		}
		return result;
	}

	public static DateTime ConvertToDateTime2(object obj, string format)
	{
		DateTime result = default(DateTime);
		string empty = string.Empty;
		if (obj != null && obj != DBNull.Value)
		{
			empty = obj.ToString();
			DateTime.TryParseExact(empty, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
			return result;
		}
		return result;
	}

	public static string ConvertDateTimeToStr(DateTime obj, string format)
	{
		string text = "yyyyMMdd";
		if (!HTKKFormater.isNullString(format))
		{
			text = format;
		}
		bool flag = 0 == 0;
		return obj.ToString(text);
	}

	public static double ConvertToDouble(object obj)
	{
		double result = 0.0;
		string empty = string.Empty;
		if (obj != null && obj != DBNull.Value)
		{
			empty = obj.ToString();
			double.TryParse(empty, out result);
			return result;
		}
		return result;
	}

	public static bool ConvertToBool(object obj)
	{
		bool result = false;
		string empty = string.Empty;
		if (obj != null && obj != DBNull.Value)
		{
			empty = obj.ToString();
			bool.TryParse(empty, out result);
			return result;
		}
		return result;
	}

	public static int ConvertToInt(object obj, int def = 0)
	{
		int result = def;
		string empty = string.Empty;
		if (obj != null && obj != DBNull.Value)
		{
			empty = obj.ToString().Replace(".", "").Replace(",", "");
			int.TryParse(empty, out result);
			return result;
		}
		return result;
	}

	public static string ConvertToString(object obj)
	{
		string empty = string.Empty;
		if (obj != null && obj != DBNull.Value)
		{
			return obj.ToString();
		}
		return "";
	}

	public static string ConvertCellValueToText(SheetView sView, int row, int col)
	{
		string result = string.Empty;
		try
		{
			Cell cell = sView.Cells[row, col];
			if (cell == null)
			{
				return string.Empty;
			}
			if (cell.CellType == null)
			{
				return ConvertToString(cell.Text);
			}
			result = cell.CellType.GetType().Name switch
			{
				"NumberCellType" => ConvertToString(cell.Value), 
				"CurrencyCellType" => ConvertToString(cell.Value), 
				"DateTimeCellType" => ConvertToString(cell.Text), 
				_ => ConvertToString(cell.Text), 
			};
		}
		catch (Exception)
		{
		}
		return result;
	}

	public static string ConvertCellValueToTextWithControlType(SheetView sView, int row, int col, string controlType)
	{
		Cell cell = sView.Cells[row, col];
		if (cell == null)
		{
			return string.Empty;
		}
		string empty = string.Empty;
		return controlType switch
		{
			"16" => ConvertToString(cell.Value), 
			"14" => ConvertToString(cell.Text), 
			_ => ConvertToString(cell.Text), 
		};
	}

	public static string GetValueImportExcel(SheetView sView, int row, int col, string controlType)
	{
		Cell cell = sView.Cells[row, col];
		if (cell == null)
		{
			return string.Empty;
		}
		string empty = string.Empty;
		return controlType switch
		{
			"16" => ConvertToString(cell.Value).Replace(".", "").Replace(" ", ""), 
			"14" => ConvertToString(cell.Text), 
			"2" => (!(cell.Text.ToLower() == "x") && !(cell.Text.ToLower() == "1") && !(cell.Text.ToLower() == "true")) ? "False" : "True", 
			_ => ConvertToString(cell.Text), 
		};
	}

	public static double ConvertToDouble1(object obj)
	{
		double result = 0.0;
		string empty = string.Empty;
		if (obj != null && obj != DBNull.Value)
		{
			empty = obj.ToString().Replace(",", "").Replace(".", "");
			double.TryParse(empty, out result);
			return result;
		}
		return result;
	}

	public static decimal ToNumber(this object obj, string decimalT = ".")
	{
		decimal num = 0m;
		string s = string.Empty;
		if (obj == null || obj == DBNull.Value || HTKKFormater.isNullString(obj.ToString()))
		{
			return 0m;
		}
		if ((object)obj.GetType() == typeof(double) || (object)obj.GetType() == typeof(double) || (object)obj.GetType() == typeof(int) || (object)obj.GetType() == typeof(int) || (object)obj.GetType() == typeof(short) || (object)obj.GetType() == typeof(long) || (object)obj.GetType() == typeof(long) || (object)obj.GetType() == typeof(decimal))
		{
			num = decimal.Parse(obj.ToString());
		}
		else if (obj.ToString().IndexOf("E") > -1)
		{
			num = decimal.Parse(obj.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture);
		}
		else
		{
			if (".".Equals(decimalT))
			{
				s = obj.ToString().Replace(",", "");
			}
			else if (",".Equals(decimalT))
			{
				s = obj.ToString().Replace(".", "").Replace(",", ".");
			}
			num = decimal.Parse(s);
		}
		return num;
	}

	public static string ToStr(this object obj)
	{
		string text = "";
		if (obj == null || HTKKFormater.isNullString(obj.ToString()))
		{
			return "";
		}
		return obj.ToString();
	}

	public static string UNICODETOTCVN(string xINPUT)
	{
		if (string.IsNullOrEmpty(xINPUT))
		{
			return xINPUT;
		}
		string strConv = xINPUT;
		ConvertFont convertFont = new ConvertFont();
		convertFont.Convert(ref strConv, FontIndex.iNotKnown, FontIndex.iTCV);
		return strConv;
	}

	public static string VN2UN(string xINPUT)
	{
		if (string.IsNullOrEmpty(xINPUT))
		{
			return xINPUT;
		}
		string strConv = xINPUT;
		ConvertFont convertFont = new ConvertFont();
		convertFont.Convert(ref strConv, FontIndex.iNotKnown, FontIndex.iUNI);
		return strConv;
	}

	public static string T2U(string xINPUT)
	{
		if (string.IsNullOrEmpty(xINPUT))
		{
			return xINPUT;
		}
		string strConv = xINPUT;
		ConvertFont convertFont = new ConvertFont();
		convertFont.Convert(ref strConv, FontIndex.iNotKnown, FontIndex.iUNI);
		return strConv;
	}

	public static DateTime GetNgayDauQuy(int iQuy, int iNam, int iNgayBdNamTc, int iThangBdNamTc)
	{
		DateTime dateTime = default(DateTime);
		try
		{
			int num = (iQuy - 1) * 3 + iThangBdNamTc;
			int year = iNam;
			if (num > 12)
			{
				num -= 12;
				year = iNam + 1;
			}
			return new DateTime(year, num, iNgayBdNamTc);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static DateTime GetNgayCuoiQuy(int iQuy, int iNam, int iNgayBdNamTc, int iThangBdNamTc)
	{
		DateTime dateTime = default(DateTime);
		try
		{
			int num = (iQuy - 1) * 3 + iThangBdNamTc + 2;
			int year = iNam;
			if (iNgayBdNamTc == 1)
			{
				int num2 = num + 1;
				year = iNam;
				int num3 = iNam;
				if (num > 12)
				{
					num -= 12;
					year = iNam + 1;
				}
				if (num2 > 12)
				{
					num2 -= 12;
					num3 = iNam + 1;
				}
				if (num3 >= 10000)
				{
					num3 = 9999;
				}
				int num4 = GetSoNgay(new DateTime(year, num, 1), new DateTime(num3, num2, 1)) - 1;
				return new DateTime(year, num, 1).AddDays(num4);
			}
			return new DateTime(year, num, 1);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
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

	public static string ToStringOrDefault(this object obj, string defaultValue = "")
	{
		try
		{
			if (obj == null || obj == DBNull.Value)
			{
				return defaultValue;
			}
			return obj.ToString();
		}
		catch
		{
			return defaultValue;
		}
	}

	public static bool IsNullOrEmpty(this object obj, string ignore = "")
	{
		try
		{
			if (string.IsNullOrEmpty(obj.ToString()))
			{
				return true;
			}
			return string.IsNullOrEmpty(obj.ToString().Replace(ignore, ""));
		}
		catch
		{
			return false;
		}
	}

	public static int ToIntOrDefault(this object obj, int defaultValue = 0)
	{
		try
		{
			if (obj == null || obj == DBNull.Value || obj.IsNullOrEmpty())
			{
				return defaultValue;
			}
			return ConvertToInt(obj.ToString().Replace(".", "").Replace(",", ""), defaultValue);
		}
		catch
		{
			return defaultValue;
		}
	}

	public static bool ToBoolOrDefault(this object obj, bool defaultValue = false)
	{
		try
		{
			if (obj == null || obj == DBNull.Value)
			{
				return defaultValue;
			}
			return ConvertToBool(obj.ToString());
		}
		catch
		{
			return defaultValue;
		}
	}

	public static bool IsNulOrEmpty(this string str)
	{
		return string.IsNullOrEmpty(str);
	}
}
