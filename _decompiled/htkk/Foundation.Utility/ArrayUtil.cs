using System;
using Foundation.Base;

namespace Foundation.Utility;

internal class ArrayUtil : BaseObject
{
	public static int[] Copy(int[] int_list)
	{
		if (int_list == null)
		{
			return null;
		}
		int[] array = new int[int_list.Length];
		for (int i = 0; i < int_list.Length; i++)
		{
			array[i] = int_list[i];
		}
		return array;
	}

	public static byte[] Copy(byte[] bytes)
	{
		if (bytes == null)
		{
			return null;
		}
		byte[] array = new byte[bytes.Length];
		for (int i = 0; i < bytes.Length; i++)
		{
			array[i] = bytes[i];
		}
		return array;
	}

	public static bool Compare(byte[] b1, byte[] b2)
	{
		if (b1 == null && b2 == null)
		{
			return true;
		}
		if (b1 == null || b2 == null)
		{
			return false;
		}
		if (b1.Length != b2.Length)
		{
			return false;
		}
		bool result = true;
		for (int i = 0; i < b1.Length; i++)
		{
			if (b1[i] != b2[i])
			{
				result = false;
				break;
			}
		}
		return result;
	}

	public static int Sort(int[] int_array, bool ascending = true)
	{
		if (int_array == null)
		{
			return -1;
		}
		int result = 0;
		try
		{
			if (ascending)
			{
				Array.Sort(int_array);
			}
			else
			{
				Array.Sort(int_array, (int i1, int i2) => i2.CompareTo(i1));
			}
		}
		catch
		{
			result = -1;
		}
		return result;
	}
}
