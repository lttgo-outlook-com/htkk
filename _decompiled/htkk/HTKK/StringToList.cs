using System;
using System.Collections.Generic;
using System.Reflection;
using TAX.HTKK.Utilities.Core;

namespace HTKK;

public class StringToList
{
	public static List<int> lstPages(string _trangIn, int _totalPage)
	{
		int num = 0;
		int num2 = 0;
		List<int> list = new List<int>();
		bool flag = false;
		list.Clear();
		if (InvalidPage(_trangIn, _totalPage))
		{
			return list = null;
		}
		if (_trangIn.Contains(","))
		{
			string[] array = _trangIn.Trim().Split(new char[1] { ',' });
			for (int i = 0; i < array.Length; i++)
			{
				if (array[i].Contains("-"))
				{
					string[] array2 = array[i].Trim().Split(new char[1] { '-' });
					int[] array3 = Array.ConvertAll(array2, (string s) => int.Parse(s));
					num = Lowest(num, array3);
					num2 = Highest(num2, array3);
					for (int num3 = num; num3 <= num2; num3++)
					{
						list.Add(num3);
					}
					if (HasDuplicates(list))
					{
						flag = true;
						break;
					}
				}
				else
				{
					list.Add(int.Parse(array[i]));
					if (HasDuplicates(list))
					{
						flag = true;
						break;
					}
				}
			}
		}
		else if (_trangIn.Contains("-") && !_trangIn.Contains(","))
		{
			string[] array = _trangIn.Trim().Split(new char[1] { '-' });
			int[] array3 = Array.ConvertAll(array, (string s) => int.Parse(s));
			num = Lowest(num, array3);
			num2 = Highest(num2, array3);
			for (int num3 = num; num3 <= num2; num3++)
			{
				list.Add(num3);
			}
			if (HasDuplicates(list))
			{
				flag = true;
			}
		}
		else
		{
			list.Add(int.Parse(_trangIn.Trim()));
		}
		if (flag)
		{
			return list = null;
		}
		return list;
	}

	private static bool InvalidPage(string pages, int _totalPage)
	{
		string[] array = pages.Trim().Split(',', '-');
		int[] array2 = Array.ConvertAll(array, (string s) => int.Parse(s));
		bool result = false;
		int[] array3 = array2;
		foreach (int num2 in array3)
		{
			if (num2 > _totalPage)
			{
				result = true;
			}
		}
		return result;
	}

	private static bool HasDuplicates(List<int> list)
	{
		List<int> list2 = new List<int>();
		bool result = false;
		try
		{
			foreach (int item in list)
			{
				if (list2.Contains(item))
				{
					result = true;
					break;
				}
				list2.Add(item);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	private static int Lowest(int lowest, int[] array)
	{
		try
		{
			lowest = array[0];
			for (int i = 0; i < array.Length; i++)
			{
				if (lowest > array[i])
				{
					lowest = array[i];
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return lowest;
	}

	private static int Highest(int highest, int[] array)
	{
		try
		{
			highest = array[0];
			for (int i = 0; i < array.Length; i++)
			{
				if (highest < array[i])
				{
					highest = array[i];
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return highest;
	}
}
