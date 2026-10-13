using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class ValidateBase
{
	private static double Evaluate(string expression)
	{
		try
		{
			DataTable dataTable = new DataTable();
			dataTable.Columns.Add("expression", typeof(string), expression);
			DataRow dataRow = dataTable.NewRow();
			dataTable.Rows.Add(dataRow);
			return double.Parse((string)dataRow["expression"]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1.0;
		}
	}

	public static bool ValidateFormula(string strValidate, string range)
	{
		try
		{
			double num = Evaluate(range);
			if (double.Parse(strValidate) == num)
			{
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool NhoHonHoacBangCellLocation(string strValidate, string range)
	{
		try
		{
			double num = Evaluate(range);
			if (double.Parse(strValidate) <= num)
			{
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool ValidateEmail(string strValidate, string range)
	{
		try
		{
			Regex regex = new Regex("^[\\W]*([\\w+\\-.%]+@[\\w\\-.]+\\.[A-Za-z] {2,4}[\\W]*,{1}[\\W]*)*([\\w+\\-.%]+@[\\w\\-.]+\\.[A-Za-z]{2,4})[\\W]*$");
			return regex.IsMatch(strValidate);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool ValidateInt(string strValidate)
	{
		try
		{
			Regex regex = new Regex("^[0-9]");
			return regex.IsMatch(strValidate);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool ValidateDouble(string strValidate)
	{
		try
		{
			Regex regex = new Regex("^-?[0-9]{1,3}(,[0-9]{3})*(\\.[0-9]+)?$|^\\([0-9]{1,3}(,[0-9]{3})*(\\.[0-9]+)?\\)$|^-?[0-9]+(\\.[0-9]+)?$|^\\([0-9]+(\\.[0-9]+)?\\)$");
			return regex.IsMatch(strValidate);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool ValidateNumberBetween(string strValidate, string range)
	{
		try
		{
			bool result = false;
			if (!ValidateDouble(strValidate))
			{
				result = false;
			}
			else
			{
				string[] array = range.Split(new char[1] { ',' });
				double num = double.Parse(array[0]);
				double num2 = double.Parse(array[1]);
				double num3 = double.Parse(strValidate);
				if (num3 >= num && num3 <= num2)
				{
					result = true;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool ValidateMaxLen(string strValidate, string strDoDai)
	{
		try
		{
			if (strValidate.Length <= int.Parse(strDoDai))
			{
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool ValidateMaSoThue(string strValidate, string range)
	{
		try
		{
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool ValidateCMTND(string strValidate, string range)
	{
		try
		{
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool ValidateBangNhau(string strValidate, string range)
	{
		try
		{
			if (double.Parse(strValidate) == double.Parse(range))
			{
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool ValidateLonHon(string strValidate, string range)
	{
		try
		{
			if (double.Parse(strValidate) > double.Parse(range))
			{
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool ValidateNhoHon(string strValidate, string range)
	{
		try
		{
			if (double.Parse(strValidate) < double.Parse(range))
			{
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool ValidateNhoHonHoacBang(string strValidate, string range)
	{
		try
		{
			if (double.Parse(strValidate) <= double.Parse(range))
			{
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool ValidateMinValue(string strValidate, string range)
	{
		try
		{
			if (double.Parse(strValidate) >= double.Parse(range))
			{
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool ValidateFullDate(string strValidate, string format)
	{
		try
		{
			DateTime t = DateTime.ParseExact(strValidate, format, CultureInfo.InvariantCulture);
			if (DateTime.Compare(t, new DateTime(1900, 1, 1)) < 0)
			{
				return false;
			}
			return true;
		}
		catch (Exception)
		{
			return false;
		}
	}

	public static List<ReturnMessageInfo> ValidateDuplicate(List<ReturnMessageInfo> listCheckDuplicate)
	{
		try
		{
			Dictionary<string, ReturnMessageInfo> dictionary = new Dictionary<string, ReturnMessageInfo>();
			List<ReturnMessageInfo> list = new List<ReturnMessageInfo>();
			foreach (ReturnMessageInfo item in listCheckDuplicate)
			{
				if (!dictionary.ContainsKey(item.ValueCellIndex))
				{
					dictionary.Add(item.ValueCellIndex, item);
					continue;
				}
				list.Add(item);
				if (!list.Contains(dictionary[item.ValueCellIndex]))
				{
					list.Add(dictionary[item.ValueCellIndex]);
				}
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static bool CheckMaSoThue(string strMST)
	{
		try
		{
			bool flag = false;
			if (strMST.Length != 10 && strMST.Length != 14 && strMST.Length != 13)
			{
				return false;
			}
			if (strMST.Length == 10)
			{
				Regex regex = new Regex("^\\d{10}$|^\\d{10}$");
				if (!regex.IsMatch(strMST))
				{
					return false;
				}
			}
			else if (strMST.Length == 13)
			{
				Regex regex = new Regex("^\\d{10}$|^\\d{10}?\\d{3}$");
				if (!regex.IsMatch(strMST))
				{
					return false;
				}
			}
			else
			{
				Regex regex = new Regex("^\\d{10}$|^\\d{10}-?\\d{3}$");
				if (!regex.IsMatch(strMST))
				{
					return false;
				}
			}
			int[] array = new int[strMST.Length];
			for (int i = 0; i < 10; i++)
			{
				array[i] = Convert.ToInt32(strMST.Substring(i, 1));
			}
			int num = 31 * array[0] + 29 * array[1] + 23 * array[2] + 19 * array[3] + 17 * array[4] + 13 * array[5] + 7 * array[6] + 5 * array[7] + 3 * array[8];
			int num2 = num % 11;
			return (array[9] == 10 - num2) ? true : false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool CheckMaSoThue5So(string strMST)
	{
		if (CheckStringIsNumber(strMST) && strMST.Length == 5)
		{
			return true;
		}
		return false;
	}

	public static bool CheckStringIsNumber(string text)
	{
		return text.All(char.IsDigit);
	}

	public static bool CheckMaSoThueCCCD(string strMST)
	{
		try
		{
			bool flag = false;
			if (strMST.Length != 10 && strMST.Length != 14 && strMST.Length != 13 && strMST.Length != 12)
			{
				return false;
			}
			if (strMST.Length == 12)
			{
				flag = true;
				Regex regex = new Regex("^\\d{12}$|^\\d{12}$");
				return regex.IsMatch(strMST);
			}
			if (strMST.Length == 10)
			{
				Regex regex = new Regex("^\\d{10}$|^\\d{10}$");
				if (!regex.IsMatch(strMST))
				{
					return false;
				}
			}
			else if (strMST.Length == 13)
			{
				Regex regex = new Regex("^\\d{10}$|^\\d{10}?\\d{3}$");
				if (!regex.IsMatch(strMST))
				{
					return false;
				}
			}
			else
			{
				Regex regex = new Regex("^\\d{10}$|^\\d{10}-?\\d{3}$");
				if (!regex.IsMatch(strMST))
				{
					return false;
				}
			}
			int[] array = new int[strMST.Length];
			for (int i = 0; i < 10; i++)
			{
				array[i] = Convert.ToInt32(strMST.Substring(i, 1));
			}
			int num = 31 * array[0] + 29 * array[1] + 23 * array[2] + 19 * array[3] + 17 * array[4] + 13 * array[5] + 7 * array[6] + 5 * array[7] + 3 * array[8];
			int num2 = num % 11;
			return (array[9] == 10 - num2) ? true : false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}
}
