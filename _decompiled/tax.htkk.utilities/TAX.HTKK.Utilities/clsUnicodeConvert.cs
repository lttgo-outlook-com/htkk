using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Microsoft.VisualBasic;

namespace TAX.HTKK.Utilities;

public class clsUnicodeConvert
{
	public enum CharacterSet
	{
		UNICODE = 1,
		VPS,
		VISCII,
		VNI,
		TCVN,
		VIQR
	}

	private XmlDocument mDOMTextFile = null;

	private string mXMLPath = "";

	private string[] mPopulate = null;

	private static clsUnicodeConvert _DefaultInstance = null;

	public static clsUnicodeConvert DefaultInstance
	{
		get
		{
			if (_DefaultInstance == null)
			{
				_DefaultInstance = new clsUnicodeConvert();
			}
			return _DefaultInstance;
		}
	}

	private string GetLocalDirectory()
	{
		string text = Path.GetDirectoryName(Application.ExecutablePath);
		if (text.Substring(Math.Max(text.Length - 1, 0)) != "\\")
		{
			text += "\\";
		}
		return text;
	}

	private string ReadUnicode(string TXMLPath)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_000c: Expected O, but got Unknown
		mDOMTextFile = new XmlDocument();
		mXMLPath = TXMLPath;
		mDOMTextFile.Load(mXMLPath);
		XmlElement documentElement = mDOMTextFile.DocumentElement;
		return ((XmlNode)documentElement).Value;
	}

	private string ReadTextFile(string FileName)
	{
		return File.ReadAllText(FileName);
	}

	private string GetVowels(CharacterSet pCharSet, bool pFlag)
	{
		string result = "";
		switch (pCharSet)
		{
		case CharacterSet.TCVN:
			result = ReadTextFile(GetLocalDirectory() + "TCVNVowels.txt");
			break;
		case CharacterSet.UNICODE:
			result = ReadUnicode(GetLocalDirectory() + "UnicodeVowels.xml");
			break;
		case CharacterSet.VIQR:
			result = (pFlag ? ReadTextFile(GetLocalDirectory() + "VIQRVowelMap.txt") : ReadTextFile(GetLocalDirectory() + "VIQRVowels.txt"));
			break;
		case CharacterSet.VISCII:
			result = ReadTextFile(GetLocalDirectory() + "VISCIIVowels.txt");
			break;
		case CharacterSet.VNI:
			result = (pFlag ? ReadTextFile(GetLocalDirectory() + "VNIVowelMap.txt") : ReadTextFile(GetLocalDirectory() + "VNIVowels.txt"));
			break;
		case CharacterSet.VPS:
			result = ReadTextFile(GetLocalDirectory() + "VPSVowels.txt");
			break;
		}
		return result;
	}

	private string StringToString(string pString, ref string Vowel1, string Vowel2)
	{
		string text = "";
		StringBuilder stringBuilder = new StringBuilder();
		int num = 0;
		for (int i = 1; i <= Strings.Len(pString); i++)
		{
			text = pString.Substring(i - 1, Math.Min(1, pString.Length - (i - 1)));
			if (text == "\r")
			{
				stringBuilder.Append("\r");
				continue;
			}
			if (text == "\n")
			{
				stringBuilder.Append("\n");
				continue;
			}
			num = Vowel1.IndexOf(text) + 1;
			if (num <= 0)
			{
				stringBuilder.Append(text);
			}
			else
			{
				stringBuilder.Append(Vowel2.Substring(num - 1, Math.Min(1, Vowel2.Length - (num - 1))));
			}
		}
		return stringBuilder.ToString();
	}

	private void PopulateFromFile(CharacterSet pCharSet, bool pFlag)
	{
		string text = "";
		switch (pCharSet)
		{
		case CharacterSet.TCVN:
			text = GetVowels(CharacterSet.TCVN, pFlag: false);
			break;
		case CharacterSet.UNICODE:
			text = GetVowels(CharacterSet.UNICODE, pFlag: false);
			break;
		case CharacterSet.VIQR:
			text = GetVowels(CharacterSet.VIQR, pFlag);
			break;
		case CharacterSet.VISCII:
			text = GetVowels(CharacterSet.VISCII, pFlag: false);
			break;
		case CharacterSet.VNI:
			text = GetVowels(CharacterSet.VNI, pFlag);
			break;
		case CharacterSet.VPS:
			text = GetVowels(CharacterSet.VPS, pFlag: false);
			break;
		}
		mPopulate = text.Split(new string[1] { "\r\n" }, StringSplitOptions.None);
	}

	private string UnicodeToMultichar(string pString, ref string UVowels)
	{
		string text = "";
		StringBuilder stringBuilder = new StringBuilder();
		int num = 0;
		for (int i = 1; i <= Strings.Len(pString); i++)
		{
			text = pString.Substring(i - 1, Math.Min(1, pString.Length - (i - 1)));
			if (text == "\r")
			{
				stringBuilder.Append("\r");
				continue;
			}
			if (text == "\n")
			{
				stringBuilder.Append("\n");
				continue;
			}
			num = UVowels.IndexOf(text) + 1;
			if (num <= 0)
			{
				stringBuilder.Append(text);
			}
			else
			{
				stringBuilder.Append(mPopulate[num - 1]);
			}
		}
		return stringBuilder.ToString();
	}

	private string MulticharToUnicode(string pString, string UVowels)
	{
		string text = "";
		StringBuilder stringBuilder = new StringBuilder();
		int num = 0;
		string text2 = "";
		string text3 = "";
		double num2 = 0.0;
		string text4 = pString;
		for (num = 0; num <= mPopulate.GetUpperBound(0) - 1; num++)
		{
			text2 = mPopulate[num];
			text = text2.Substring(0, Math.Min(3, text2.Length)).TrimEnd(new char[0]);
			text3 = "|" + text2.Substring(Math.Max(text2.Length - 3, 0));
			text4 = Strings.Replace(text4, text, text3, 1, -1, (CompareMethod)0);
		}
		num = 1;
		int num3 = Strings.Len(text4);
		while (num <= num3)
		{
			text = text4.Substring(num - 1, Math.Min(1, text4.Length - (num - 1)));
			if (text == "\r")
			{
				stringBuilder.Append("\r");
				num++;
				continue;
			}
			if (text == "\n")
			{
				stringBuilder.Append("\n");
				num++;
				continue;
			}
			if (text != "|")
			{
				stringBuilder.Append(text);
				num++;
				continue;
			}
			num2 = Conversion.Val(text4.Substring(num, Math.Min(3, text4.Length - num)));
			if (num2 == 0.0)
			{
				stringBuilder.Append(text);
				num++;
			}
			else
			{
				stringBuilder.Append(UVowels.Substring(System.Convert.ToInt32(num2) - 1, Math.Min(1, UVowels.Length - (System.Convert.ToInt32(num2) - 1))));
				num += 4;
			}
		}
		return stringBuilder.ToString();
	}

	public string Convert(string pString, CharacterSet pFrom, CharacterSet pTo)
	{
		string result = "";
		string text = "";
		if (pString == "" || pFrom == pTo)
		{
			return pString;
		}
		string Vowel = GetVowels(CharacterSet.UNICODE, pFlag: false);
		if (pFrom == CharacterSet.UNICODE)
		{
			text = GetVowels(pTo, pFlag: false);
			switch (pTo)
			{
			case CharacterSet.VPS:
			case CharacterSet.VISCII:
			case CharacterSet.TCVN:
				result = StringToString(pString, ref Vowel, text);
				break;
			case CharacterSet.VIQR:
				PopulateFromFile(CharacterSet.VIQR, pFlag: false);
				result = UnicodeToMultichar(pString, ref Vowel);
				break;
			case CharacterSet.VNI:
				PopulateFromFile(CharacterSet.VNI, pFlag: false);
				result = UnicodeToMultichar(pString, ref Vowel);
				break;
			}
		}
		else
		{
			text = GetVowels(pFrom, pFlag: false);
			switch (pFrom)
			{
			case CharacterSet.VPS:
			case CharacterSet.VISCII:
			case CharacterSet.TCVN:
				result = StringToString(pString, ref text, Vowel);
				break;
			case CharacterSet.VIQR:
				PopulateFromFile(CharacterSet.VIQR, pFlag: true);
				result = MulticharToUnicode(pString, Vowel);
				break;
			case CharacterSet.VNI:
				PopulateFromFile(CharacterSet.VNI, pFlag: true);
				result = MulticharToUnicode(pString, Vowel);
				break;
			}
		}
		return result;
	}
}
