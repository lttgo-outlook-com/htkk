using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using TAX.HTKK.Utilities.Control;

namespace TAX.HTKK.Utilities.Core;

public class HTKKFormater
{
	public static bool IsNumber(string strNumber, out long intNumber)
	{
		bool result = false;
		long result2 = 0L;
		long.TryParse(strNumber, out result2);
		intNumber = result2;
		return result;
	}

	public static bool IsNumber(string strNumber)
	{
		try
		{
			long result = 0L;
			return long.TryParse(strNumber, out result);
		}
		catch (Exception)
		{
			return false;
		}
	}

	public static bool IsMonth(string strMonth)
	{
		bool result = true;
		try
		{
			int result2 = 0;
			int.TryParse(strMonth, out result2);
			if (result2 < 1 || result2 > 12)
			{
				result = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public static bool IsDay(string strDay)
	{
		bool result = true;
		try
		{
			int num = int.Parse(strDay);
			if (num < 1 || num > 31)
			{
				result = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public static bool IsYear(string strYear)
	{
		bool result = true;
		try
		{
			int num = int.Parse(strYear);
			if (num < 0 || num > 9999)
			{
				result = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public static bool IsHour(string strHour)
	{
		bool result = true;
		try
		{
			int num = int.Parse(strHour);
			if (num < 0 || num > 23)
			{
				result = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public static bool IsMinute(string strMinute)
	{
		bool result = true;
		try
		{
			int num = int.Parse(strMinute);
			if (num < 0 || num > 59)
			{
				result = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public static bool IsSecond(string strSecond)
	{
		bool result = true;
		try
		{
			int num = int.Parse(strSecond);
			if (num < 0 || num > 59)
			{
				result = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public static string FormatTimeVn(DateTime dt, string defaultText)
	{
		return (ToDateTime(dt) != new DateTime(1900, 1, 1)) ? dt.ToString("dd-mm-yy") : defaultText;
	}

	public static DateTime ToDateTime(object obj)
	{
		DateTime dateTime;
		try
		{
			dateTime = Convert.ToDateTime(obj);
		}
		catch
		{
			dateTime = DateTime.Now;
		}
		if (dateTime == new DateTime(1, 1, 1))
		{
			return DateTime.Now;
		}
		return dateTime;
	}

	public static DateTime ConvertToDateTime(string input, out bool success)
	{
		try
		{
			success = true;
			string text = input.Replace(" ", "").Replace("_", "");
			if (text.Length != 10)
			{
				success = false;
			}
			DateTimeFormatInfo dateTimeFormatInfo = new DateTimeFormatInfo();
			dateTimeFormatInfo.ShortDatePattern = "dd/MM/yyyy";
			dateTimeFormatInfo.DateSeparator = "/";
			DateTimeFormatInfo provider = dateTimeFormatInfo;
			return Convert.ToDateTime(text, provider);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			success = false;
			return DateTime.Parse("1/1/1900");
		}
	}

	public static DateTime ToDateTime(object obj, IFormatProvider format)
	{
		try
		{
			return Convert.ToDateTime(obj, format);
		}
		catch
		{
			return DateTime.Now;
		}
	}

	public static DateTime ToDateTime(object obj, DateTime defaultValue)
	{
		DateTime dateTime;
		try
		{
			dateTime = Convert.ToDateTime(obj);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			dateTime = DateTime.Now;
		}
		if (dateTime == new DateTime(1, 1, 1))
		{
			return defaultValue.AddDays(1.0);
		}
		return dateTime;
	}

	public static DateTime ToDateTime(object obj, string format, DateTime defaultValue)
	{
		try
		{
			DateTimeFormatInfo dateTimeFormatInfo = new DateTimeFormatInfo();
			dateTimeFormatInfo.ShortDatePattern = format;
			DateTimeFormatInfo provider = dateTimeFormatInfo;
			return Convert.ToDateTime(obj, provider);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return defaultValue;
		}
	}

	public static string ConvertMdytoDMY(string date)
	{
		string[] array = date.Split(new char[1] { ' ' });
		string empty = string.Empty;
		string[] array2 = null;
		string empty2 = string.Empty;
		try
		{
			empty = array[0];
			array2 = empty.Split(new char[1] { '/' });
			return array2[1] + "/" + array2[0] + "/" + array2[2];
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string NgayThang()
	{
		try
		{
			string text = "";
			string[] array = new string[7] { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday" };
			string[] array2 = new string[7] { "Thứ hai", "Thứ ba", "Thứ tư", "Thứ năm", "Thứ sáu", "Thứ bảy", "Chủ Nhật" };
			for (int i = 0; i < array.Length; i++)
			{
				if (DateTime.Now.DayOfWeek.ToString().ToUpper() == array[i].ToUpper())
				{
					text = array2[i];
					break;
				}
			}
			return "Hôm nay: " + text + ", " + DateTime.Now.ToString("dd - MM - yyyy");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string FormatStringToDate(string value)
	{
		try
		{
			value = value.Replace("/", "");
			string result = string.Empty;
			if (value.Length == 4)
			{
				result = value;
			}
			if (value.Length == 6)
			{
				result = value.Substring(0, 2) + "/" + value.Substring(2, 4);
			}
			if (value.Length == 8)
			{
				result = value.Substring(0, 2) + "/" + value.Substring(2, 2) + "/" + value.Substring(4, 4);
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string FormatSubStringToDate(string strValue)
	{
		try
		{
			string text = long.Parse(strValue).ToString();
			string result = string.Empty;
			if (text.Length == 4)
			{
				result = text;
			}
			else if (text.Length <= 6)
			{
				text = text.PadLeft(6, '0');
				result = text.Substring(0, 2) + "/" + text.Substring(2, 4);
			}
			else if (text.Length <= 8)
			{
				text = text.PadLeft(8, '0');
				result = text.Substring(0, 2) + "/" + text.Substring(2, 2) + "/" + text.Substring(4, 4);
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string toDateMMYYYY(string value, bool isMess)
	{
		try
		{
			if (isNullString(value))
			{
				return "";
			}
			string result = "";
			string text = value;
			text = text.Replace("\\", "").Replace("/", "").Replace("-", "")
				.Replace(" ", "");
			if (text.Length == 6)
			{
				if (!IsMonth(text.Substring(0, 2)) || !IsYear(text.Substring(2, 4)))
				{
					if (isMess)
					{
					}
					result = "";
				}
				else
				{
					result = text.Substring(0, 2) + "/" + text.Substring(2, 4);
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string toDateDDMMYYYY(string value, bool isMess)
	{
		try
		{
			if (isNullString(value))
			{
				return "";
			}
			string result = "";
			string text = value;
			text = text.Replace("\\", "").Replace("/", "").Replace("-", "")
				.Replace(" ", "");
			if (text.Length == 4)
			{
				if (!IsYear(text))
				{
					if (isMess)
					{
					}
					result = "";
				}
				else
				{
					result = "01/01/" + text;
				}
			}
			if (text.Length == 6)
			{
				if (!IsMonth(text.Substring(2, 2)) || !IsYear(text.Substring(4, 4)))
				{
					if (isMess)
					{
					}
					result = "";
				}
				else
				{
					result = "01/" + text.Substring(2, 2) + "/" + text.Substring(4, 4);
				}
			}
			if (text.Length == 8)
			{
				if (!IsDay(text.Substring(0, 2)) || !IsMonth(text.Substring(2, 2)) || !IsYear(text.Substring(4, 4)))
				{
					if (isMess)
					{
					}
					result = "";
				}
				else
				{
					result = text.Substring(0, 2) + "/" + text.Substring(2, 2) + "/" + text.Substring(4, 4);
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string toTimeHHMMSS(string value, bool isMess)
	{
		try
		{
			if (isNullString(value))
			{
				return "";
			}
			string result = "";
			string text = value;
			text = text.Replace("\\", "").Replace("/", "").Replace("-", "")
				.Replace(" ", "")
				.Replace(":", "");
			if (text.Length == 2)
			{
				if (!IsHour(text))
				{
					if (isMess)
					{
					}
					result = "";
				}
				else
				{
					result = text + ":00:00";
				}
			}
			if (text.Length == 4)
			{
				if (!IsHour(text.Substring(0, 2)) || !IsMinute(text.Substring(2, 2)))
				{
					if (isMess)
					{
					}
					result = "";
				}
				else
				{
					result = text.Substring(0, 2) + ":" + text.Substring(2, 2) + ":00";
				}
			}
			if (text.Length == 6)
			{
				if (!IsHour(text.Substring(0, 2)) || !IsMinute(text.Substring(2, 2)) || !IsSecond(text.Substring(4, 2)))
				{
					if (isMess)
					{
					}
					result = "";
				}
				else
				{
					result = text.Substring(0, 2) + ":" + text.Substring(2, 2) + ":" + text.Substring(4, 2);
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string toDateAndTimeDDMMYYYYHHMMSS(string value, bool isMess)
	{
		try
		{
			if (isNullString(value))
			{
				return "";
			}
			string text = "";
			string text2 = value;
			text2 = text2.Replace("\\", "").Replace("/", "").Replace("-", "")
				.Replace(" ", "")
				.Replace(":", "");
			if (text2.Length == 4)
			{
				text = (IsYear(text2) ? ("01/01/" + text2 + " 00:00:00") : "");
			}
			else if (text2.Length == 6)
			{
				text = ((IsMonth(text2.Substring(0, 2)) && IsYear(text2.Substring(2, 4))) ? ("01/" + text2.Substring(0, 2) + "/" + text2.Substring(2, 4) + " 00:00:00") : "");
			}
			else if (text2.Length == 8)
			{
				text = ((IsDay(text2.Substring(0, 2)) && IsMonth(text2.Substring(2, 2)) && IsYear(text2.Substring(4, 4))) ? (text2.Substring(0, 2) + "/" + text2.Substring(2, 2) + "/" + text2.Substring(4, 4) + " 00:00:00") : "");
			}
			else if (text2.Length == 10)
			{
				text = ((IsDay(text2.Substring(0, 2)) && IsMonth(text2.Substring(2, 2)) && IsYear(text2.Substring(4, 4)) && IsHour(text2.Substring(8, 2))) ? (text2.Substring(0, 2) + "/" + text2.Substring(2, 2) + "/" + text2.Substring(4, 4) + " " + text2.Substring(8, 2) + ":00:00") : "");
			}
			else if (text2.Length == 12)
			{
				text = ((IsDay(text2.Substring(0, 2)) && IsMonth(text2.Substring(2, 2)) && IsYear(text2.Substring(4, 4)) && IsHour(text2.Substring(8, 2)) && IsMinute(text2.Substring(10, 2))) ? (text2.Substring(0, 2) + "/" + text2.Substring(2, 2) + "/" + text2.Substring(4, 4) + " " + text2.Substring(8, 2) + ":" + text2.Substring(10, 2) + ":00") : "");
			}
			else if (text2.Length == 14)
			{
				text = ((IsDay(text2.Substring(0, 2)) && IsMonth(text2.Substring(2, 2)) && IsYear(text2.Substring(4, 4)) && IsHour(text2.Substring(8, 2)) && IsMinute(text2.Substring(10, 2)) && IsSecond(text2.Substring(12, 2))) ? (text2.Substring(0, 2) + "/" + text2.Substring(2, 2) + "/" + text2.Substring(4, 4) + " " + text2.Substring(8, 2) + ":" + text2.Substring(10, 2) + ":" + text2.Substring(12, 2)) : "");
			}
			if (!string.IsNullOrEmpty(text) || isMess)
			{
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static int FormatStringToDateNumberBegin(string value)
	{
		try
		{
			value = value.Replace("/", "");
			int result = 0;
			if (value.Length == 4)
			{
				result = int.Parse(value + "0101");
			}
			if (value.Length == 6)
			{
				result = int.Parse(value.Substring(2, 4) + value.Substring(0, 2) + "01");
			}
			if (value.Length == 8)
			{
				result = int.Parse(value.Substring(4, 4) + value.Substring(2, 2) + value.Substring(0, 2));
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return 0;
		}
	}

	public static int FormatStringToDateNumberEnd(string value)
	{
		try
		{
			value = value.Replace("/", "");
			int result = 0;
			if (value.Length == 4)
			{
				result = int.Parse(value + "1231");
			}
			if (value.Length == 6)
			{
				result = int.Parse(value.Substring(2, 4) + value.Substring(0, 2) + "31");
			}
			if (value.Length == 8)
			{
				result = int.Parse(value.Substring(4, 4) + value.Substring(2, 2) + "31");
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return 0;
		}
	}

	public static bool IsDateTime(string strDate, string strFormat, out DateTime dtDate)
	{
		try
		{
			if (!DateTime.TryParseExact(strDate, strFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out dtDate))
			{
				dtDate = DateTime.MinValue;
				return false;
			}
			dtDate = DateTime.ParseExact(strDate, strFormat, CultureInfo.InvariantCulture);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			dtDate = DateTime.MinValue;
			return false;
		}
	}

	public static string FormatDdmmyyyy(string sDate)
	{
		//IL_028e: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c6: Unknown result type (might be due to invalid IL or missing references)
		//IL_011f: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_01cb: Unknown result type (might be due to invalid IL or missing references)
		string result = sDate;
		try
		{
			string text = "";
			string text2 = "";
			string text3 = "";
			if (!sDate.Trim().Equals("") || sDate.Trim().Length > 0)
			{
				string[] array = sDate.Split(new char[1] { '/' });
				text = array[0];
				text2 = array[1];
				text3 = array[2].Replace(".", "");
				if (ConvertHelpers.ConvertToInt(text) < 1 || ConvertHelpers.ConvertToInt(text) > 31)
				{
					frmMessage frmMessage = new frmMessage("0071", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return result;
				}
				text = text.PadLeft(2, '0');
				if (ConvertHelpers.ConvertToInt(text2) < 1 || ConvertHelpers.ConvertToInt(text2) > 12)
				{
					frmMessage frmMessage = new frmMessage("0071", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return result;
				}
				text2 = text2.PadLeft(2, '0');
				if (ConvertHelpers.ConvertToInt(text3) < 0 || ConvertHelpers.ConvertToInt(text3) > 9999)
				{
					frmMessage frmMessage = new frmMessage("0071", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return result;
				}
				if (ConvertHelpers.ConvertToInt(text3) >= 0 && ConvertHelpers.ConvertToInt(text3) <= 999)
				{
					text3 = ConvertHelpers.ConvertToString(2000 + ConvertHelpers.ConvertToInt(text3));
				}
				if (ConvertHelpers.ConvertToInt(text3) < 1900)
				{
					frmMessage frmMessage = new frmMessage("0071", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return result;
				}
				result = text + "/" + text2 + "/" + text3;
				DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(text3), ConvertHelpers.ConvertToInt(text2), ConvertHelpers.ConvertToInt(text));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage = new frmMessage("0071", "3", "", "Thông báo");
			((Form)frmMessage).ShowDialog();
			result = "";
		}
		return result;
	}

	private static bool IsDateTime(string result, string p, DateTime dDate)
	{
		throw new NotImplementedException();
	}

	public static string ConvertToKhongDau(string chuCoDau)
	{
		try
		{
			int num = -1;
			char[] anyOf = "áàảãạâấầẩẫậăắằẳẵặđéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵÁÀẢÃẠÂẤẦẨẪẬĂẮẰẲẴẶĐÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴ".ToCharArray();
			while ((num = chuCoDau.IndexOfAny(anyOf)) != -1)
			{
				int index = "áàảãạâấầẩẫậăắằẳẵặđéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵÁÀẢÃẠÂẤẦẨẪẬĂẮẰẲẴẶĐÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴ".IndexOf(chuCoDau[num]);
				chuCoDau = chuCoDau.Replace(chuCoDau[num], "aaaaaaaaaaaaaaaaadeeeeeeeeeeeiiiiiooooooooooooooooouuuuuuuuuuuyyyyyAAAAAAAAAAAAAAAAADEEEEEEEEEEEIIIIIOOOOOOOOOOOOOOOOOUUUUUUUUUUUYYYYY"[index]);
			}
			return chuCoDau;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string CutWordLength(string str, int length)
	{
		try
		{
			if (string.IsNullOrEmpty(str))
			{
				return "";
			}
			if (str.Length > length)
			{
				string[] array = str.Split(new char[1] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
				if (array[0].Length > length)
				{
					throw new ArgumentException("Từ đầu tiên dài hơn chuỗi cần cắt");
				}
				StringBuilder stringBuilder = new StringBuilder();
				string[] array2 = array;
				foreach (string text in array2)
				{
					if (string.Concat(stringBuilder, text).Length < length)
					{
						stringBuilder.Append(text + " ");
					}
				}
				return $"{stringBuilder}...";
			}
			return str;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static int GetColunmIndexByColumName(string columnName)
	{
		try
		{
			int result = -1;
			if (HTKKConfig.CELL_DATA.ContainsKey(columnName))
			{
				foreach (KeyValuePair<string, string> cELL_DATum in HTKKConfig.CELL_DATA)
				{
					if (cELL_DATum.Key == columnName)
					{
						result = int.Parse(cELL_DATum.Value) - 1;
						break;
					}
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1;
		}
	}

	public static int GetTotalPage(int totalRecords, int pageSize)
	{
		try
		{
			int num = totalRecords % pageSize;
			int num2 = (totalRecords - num) / pageSize;
			if (num > 0)
			{
				num2++;
			}
			return num2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1;
		}
	}

	public static bool IsEmail(string strEmail)
	{
		try
		{
			string pattern = "^(([\\w-]+\\.)+[\\w-]+|([a-zA-Z]{1}|[\\w-]{2,}))@((([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\\.([0-1]?\r\n\t\t\t\t[0-9]{1,2}|25[0-5]|2[0-4][0-9])\\.([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\\.([0-1]?\r\n\t\t\t\t[0-9]{1,2}|25[0-5]|2[0-4][0-9])){1}|([a-zA-Z0-9]+[\\w-]+\\.)+[a-zA-Z]{1}[a-zA-Z0-9-]{1,23})$";
			if (strEmail != null)
			{
				return Regex.IsMatch(strEmail, pattern);
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static string GetTenToKhaiChinh(string fileName, string versionAPP, string versionXML)
	{
		try
		{
			string text = fileName.Replace(".xml", "");
			string oldValue = text.Substring(text.Length - 6 - 54 + 1);
			return text.Replace(versionAPP + "_" + versionXML + "_", "").Replace(oldValue, "");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetKyTinhThueShow(string fileName, string versionAPP, string versionXML, string strTenToKhai)
	{
		try
		{
			return fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[0];
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetQuy(string fileName, string versionAPP, string versionXML, string strTenToKhai)
	{
		try
		{
			string text = fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[0];
			if (text == "Q")
			{
				string text2 = fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[1];
				return text2.Substring(2, 2);
			}
			return string.Empty;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetThang(string fileName, string versionAPP, string versionXML, string strTenToKhai)
	{
		try
		{
			string text = fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[0];
			if (text == "M")
			{
				string text2 = fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[1];
				return text2.Substring(2, 2);
			}
			return string.Empty;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetNam(string fileName, string versionAPP, string versionXML, string strTenToKhai)
	{
		try
		{
			string text = fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[0];
			if (text == "Y")
			{
				string text2 = fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[1];
				return text2.Substring(4, 4);
			}
			if (text == "M")
			{
				string text2 = fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[1];
				return text2.Substring(4, 4);
			}
			return string.Empty;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetNgay(string fileName, string versionAPP, string versionXML, string strTenToKhai)
	{
		try
		{
			string text = fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[0];
			if (text == "D")
			{
				return fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[1];
			}
			return string.Empty;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetTuNgay(string fileName, string versionAPP, string versionXML, string strTenToKhai)
	{
		try
		{
			string text = fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[2];
			if (text == "00000000")
			{
				return string.Empty;
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetDenNgay(string fileName, string versionAPP, string versionXML, string strTenToKhai)
	{
		try
		{
			string text = fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[3];
			if (text == "00000000")
			{
				return string.Empty;
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetTuThang(string fileName, string versionAPP, string versionXML, string strTenToKhai)
	{
		try
		{
			string text = fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[4];
			if (text == "00000000")
			{
				return string.Empty;
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetDenThang(string fileName, string versionAPP, string versionXML, string strTenToKhai)
	{
		try
		{
			string text = fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[5];
			if (text == "00000000")
			{
				return string.Empty;
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetNgayLapKHBS(string fileName, string versionAPP, string versionXML, string strTenToKhai)
	{
		try
		{
			string text = fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[6];
			if (text == "00000000")
			{
				return string.Empty;
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetLan(string fileName, string versionAPP, string versionXML, string strTenToKhai)
	{
		try
		{
			return fileName.Replace(versionAPP + "_" + versionXML + "_" + strTenToKhai + "_", "").Split(new char[1] { '_' })[7].Replace(".xml", "");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetVersionApp(string fileName)
	{
		try
		{
			return fileName.Split(new char[1] { '_' })[0];
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetVersionXML(string fileName)
	{
		try
		{
			string text = fileName;
			text = text.Replace(".xml", "");
			string[] array = text.Split(new char[1] { '_' });
			if (array.Length > 0)
			{
				return array[^1];
			}
			return string.Empty;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string KieuKy(string kyTinhThue)
	{
		try
		{
			string empty = string.Empty;
			return kyTinhThue switch
			{
				"M" => "Tháng", 
				"Q" => "Quý", 
				"Y" => "Năm", 
				"D_Y" => "", 
				_ => kyTinhThue, 
			};
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetKieuKyByLengh(string kyTinhThue, string id)
	{
		try
		{
			string result = string.Empty;
			if (kyTinhThue.StartsWith("Q"))
			{
				return "Q";
			}
			if (kyTinhThue.StartsWith("K"))
			{
				return "K";
			}
			switch (kyTinhThue.Length)
			{
			case 4:
				result = "Y";
				break;
			case 6:
				result = "M";
				break;
			case 7:
				result = "Q";
				if (id == "95")
				{
					result = "Y";
				}
				break;
			case 8:
				result = "D";
				break;
			case 11:
				result = "D";
				break;
			case 13:
				result = "M";
				break;
			case 14:
				result = "Q";
				break;
			case 17:
				result = (("80;82;301;G71;C9;D6;Q4;X9;Y1;".IndexOf(id) <= -1) ? ((kyTinhThue.IndexOf("CSM") <= -1) ? "Y" : "D") : "D_D");
				break;
			case 18:
				result = "Y";
				break;
			case 21:
				result = "Q";
				break;
			case 22:
				result = "D_Y";
				break;
			case 9:
				result = "Y";
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

	public static bool isNullString(string str)
	{
		bool result = false;
		if (string.IsNullOrEmpty(str))
		{
			result = true;
		}
		else if (str.Trim().Length == 0)
		{
			result = true;
		}
		else if (str == "null")
		{
			result = true;
		}
		return result;
	}

	public static bool IsNullValueAc(string strValue)
	{
		strValue = strValue.Replace(".", "");
		strValue = strValue.Replace(" ", "");
		if ("".Equals(strValue))
		{
			return true;
		}
		return false;
	}
}
