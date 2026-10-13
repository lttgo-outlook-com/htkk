using System;

namespace TAX.HTKK.Utilities.Core;

public static class MoneyHelpers
{
	private static string[] strSo = new string[10] { "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };

	private static int a;

	private static int b;

	private static int c;

	public static string MoneyToText(string number)
	{
		string text = converNumToString(slipArray(number)).Replace("  ", " ").Trim() + " đồng";
		return text.Substring(0, 1).ToUpper() + text.Substring(1, text.Length - 1);
	}

	public static string[] slipArray(string input)
	{
		int num = 0;
		int num2 = input.Length;
		string[] array = ((num2 % 3 != 0) ? new string[num2 / 3 + 1] : new string[num2 / 3]);
		if (num2 < 3)
		{
			array[0] = input;
		}
		else
		{
			while (num2 >= 3)
			{
				array[num] = input.Remove(0, num2 - 3);
				input = input.Remove(num2 - 3, 3);
				num++;
				num2 -= 3;
			}
			if (num2 > 0)
			{
				array[num] = input;
			}
		}
		return array;
	}

	public static string converNumToString(string[] list)
	{
		string text = "";
		int num = list.Length;
		if (num <= 4)
		{
			if (num == 1)
			{
				text = readThousand(list[0]);
			}
			if (num == 2)
			{
				if (ConvertHelpers.ConvertToDouble(list[1]) != 0.0)
				{
					text = readThousand(list[1]) + " nghìn ";
				}
				if (ConvertHelpers.ConvertToDouble(list[0]) != 0.0)
				{
					text += readThousand(list[0]);
				}
			}
			if (num == 3)
			{
				if (ConvertHelpers.ConvertToDouble(list[2]) != 0.0)
				{
					text = readThousand(list[2]) + " triệu ";
				}
				if (ConvertHelpers.ConvertToDouble(list[1]) != 0.0)
				{
					text = text + readThousand(list[1]) + " nghìn ";
				}
				if (ConvertHelpers.ConvertToDouble(list[0]) != 0.0)
				{
					text += readThousand(list[0]);
				}
			}
			if (num == 4)
			{
				if (ConvertHelpers.ConvertToDouble(list[3]) != 0.0)
				{
					text = readThousand(list[3]) + " tỷ ";
				}
				if (ConvertHelpers.ConvertToDouble(list[2]) != 0.0)
				{
					text = text + readThousand(list[2]) + " triệu ";
				}
				if (ConvertHelpers.ConvertToDouble(list[1]) != 0.0)
				{
					text = text + readThousand(list[1]) + " nghìn ";
				}
				if (ConvertHelpers.ConvertToDouble(list[0]) != 0.0)
				{
					text += readThousand(list[0]);
				}
			}
		}
		if (num > 4)
		{
			string[] array = new string[3];
			string[] array2 = new string[num - 3];
			for (int i = 0; i < 3; i++)
			{
				array[i] = list[i];
			}
			for (int i = 0; i < num - 3; i++)
			{
				array2[i] = list[3 + i];
			}
			text = converNumToString(array2) + " tỷ " + converNumToString(array);
		}
		return text;
	}

	public static string readThousand(string input)
	{
		string result = "";
		input = input.Trim();
		string text = input;
		int length = text.Length;
		if (length == 1)
		{
			result = strSo[Convert.ToInt32(input)];
		}
		if (length == 2)
		{
			a = Convert.ToInt32(Convert.ToString(text[0]));
			b = Convert.ToInt32(Convert.ToString(text[1]));
			if (a != 1)
			{
				result = ((b == 5) ? (strSo[a] + " mươi lăm") : ((b == 4) ? (strSo[a] + " mươi tư") : ((b == 1) ? (strSo[a] + " mươi mốt") : ((b != 0) ? (strSo[a] + " mươi " + readThousand(Convert.ToString(text[1]))) : (strSo[a] + " mươi ")))));
			}
			if (a == 1)
			{
				result = ((b == 5) ? " mười lăm" : ((b != 0) ? (" mười " + readThousand(Convert.ToString(text[1]))) : " mười "));
			}
		}
		if (length == 3)
		{
			a = Convert.ToInt32(Convert.ToString(text[0]));
			b = Convert.ToInt32(Convert.ToString(text[1]));
			c = Convert.ToInt32(Convert.ToString(text[2]));
			result = ((b != 0) ? (strSo[a] + " trăm " + readThousand(Convert.ToString(text[1]) + Convert.ToString(text[2]))) : ((c != 0) ? (strSo[a] + " trăm linh " + readThousand(Convert.ToString(text[2]))) : (strSo[a] + " trăm ")));
		}
		return result;
	}
}
