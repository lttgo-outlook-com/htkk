using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using FarPoint.Win;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using HTKKDanhmuc;
using HTKKSerializer;
using Newtonsoft.Json.Linq;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class GeneralFunctions
{
	private sealed class SheetTablePageRenderState
	{
		public readonly List<KeyValuePair<int, int>> InsertedRows = new List<KeyValuePair<int, int>>();

		public readonly Dictionary<int, bool> OriginalRowVisibility = new Dictionary<int, bool>();
	}

	private sealed class SheetTablePageSegment
	{
		public int TableIndex;

		public int FirstDataIndex;

		public int DataCount;

		public int TargetDataRow;
	}

	private sealed class SheetTablePage
	{
		public readonly List<SheetTablePageSegment> Segments = new List<SheetTablePageSegment>();

		public int DataCount;
	}

	public static string MA_COLOR_LOCKED = "#e8f5ff";

	public static string MA_COLOR_UNLOCK = "#FFFFFF";

	public static Keys KEY_PRESS;

	public static int _page = 0;

	public static int _pageSize = 0;

	public static int _totalPage = 0;

	public static HTKKCMCBigDataInfo dataInfoOld = null;

	public static string strMauSoHD_01GTKT;

	public static Dictionary<string, string> dicVar = new Dictionary<string, string>();

	public static Dictionary<string, Dictionary<string, string>> dicDms = new Dictionary<string, Dictionary<string, string>>();

	private static readonly Dictionary<FpSpread, Dictionary<int, SheetTablePageRenderState>> SheetTablePageRenderStates = new Dictionary<FpSpread, Dictionary<int, SheetTablePageRenderState>>();

	public static bool CheckNamTaiChinh(string strNgayBatDauNamTaiChinh)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		bool result = true;
		try
		{
			if (strNgayBatDauNamTaiChinh.Length == 0)
			{
				frmMessage frmMessage = new frmMessage("0061", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return false;
			}
			if (strNgayBatDauNamTaiChinh.Length == 0)
			{
				frmMessage frmMessage = new frmMessage("0061", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public static DateTime getDateKHBS(string strFileName, HTKKHeaderInfo hInfo, int xSoLan)
	{
		DateTime result = default(DateTime);
		try
		{
			string[] files = Directory.GetFiles("..\\DataFiles\\" + hInfo.mst);
			for (int i = 0; i < files.Length; i++)
			{
				if (files[i].IndexOf(strFileName) <= 0)
				{
					continue;
				}
				int num = files[i].LastIndexOf('_');
				string s = files[i].Substring(num + 1, 2);
				if (int.Parse(s) == xSoLan)
				{
					int num2 = files[i].LastIndexOf('_');
					string text = files[i].Substring(num2 - 8, 8);
					string text2 = text.Substring(0, 2);
					string text3 = text.Substring(2, 2);
					string text4 = text.Substring(4, 4);
					if (int.Parse(text2) > 0 && int.Parse(text3) > 0 && int.Parse(text4) > 0)
					{
						text = DateTime.ParseExact(text2 + "/" + text3 + "/" + text4, "d/M/yyyy", null).ToString("dd/MM/yyyy");
						result = DateTime.Parse(text);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static bool CheckKHBSLanTruoc(string strFileName, HTKKHeaderInfo hInfo, int xSoLan)
	{
		bool result = false;
		try
		{
			string[] files = Directory.GetFiles("..\\DataFiles\\" + hInfo.mst);
			for (int i = 0; i < files.Length; i++)
			{
				if (files[i].Equals(strFileName))
				{
					string s = files[i].Split(new char[1] { '\\' })[3].Split(new char[1] { '_' })[0].Replace("bs", "");
					if (int.Parse(s) == xSoLan)
					{
						result = true;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
		return result;
	}

	public static string GetXMLFileName(string path, string filename, string version, string fileExten)
	{
		if (string.Compare(filename.ToUpper(), "HEADER") == 0)
		{
			return Path.Combine(path, $"{filename}.{fileExten}");
		}
		return Path.Combine(path, string.Format("{0}_{1}.{2}", filename, version.Replace(".", ""), fileExten));
	}

	public static bool CheckTrungKy(string fromDate1, string toDate1, string fromDate2, string toDate2)
	{
		bool flag = true;
		try
		{
			int num = int.Parse(fromDate1);
			int num2 = int.Parse(toDate1);
			int num3 = int.Parse(fromDate2);
			int num4 = int.Parse(toDate2);
			if (num3 == num && num2 == num4)
			{
				return true;
			}
			if (num3 <= num && num <= num4)
			{
				return false;
			}
			if (num <= num3 && num3 <= num2)
			{
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			flag = false;
		}
		return flag;
	}

	public static string GetValue(SheetView sView, string index)
	{
		string result = null;
		try
		{
			if (index.Length > 0)
			{
				string[] array = index.Split(new char[1] { '_' });
				if (array.Length == 2)
				{
					int num = -1;
					int num2 = -1;
					num2 = int.Parse(array[1]) - 1;
					num = HTKKFile.ColumnLettersToIndex(array[0]) - 1;
					result = ((num < 0 || num2 < 0) ? null : ((sView.Cells[num2, num].Value == null) ? "" : sView.Cells[num2, num].Value.ToString()));
				}
				else
				{
					result = null;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = null;
		}
		return result;
	}

	public static string GetValue(SheetView sView, int iRow, int iCol)
	{
		string text = null;
		try
		{
			if (iCol >= 0 && iRow >= 0)
			{
				if (sView.Cells[iRow - 1, iCol - 1].Value != null)
				{
					return sView.Cells[iRow - 1, iCol - 1].Value.ToString();
				}
				return "";
			}
			return null;
		}
		catch (Exception)
		{
			return null;
		}
	}

	public static string GetValueCustom(SheetView sView, int iRow, int iCol, string controlType)
	{
		string result = string.Empty;
		try
		{
			if (iCol >= 0 && iRow >= 0)
			{
				if (sView.Cells[iRow - 1, iCol - 1].Value != null)
				{
					string text = "";
					switch (controlType)
					{
					case "2":
						text = sView.Cells[iRow - 1, iCol - 1].Value.ToString();
						result = ((!(text.ToLower() == "true")) ? "False" : "True");
						break;
					case "14":
						text = sView.Cells[iRow - 1, iCol - 1].Value.ToString();
						if (text.Trim().Length == 8)
						{
							result = text.Substring(0, 2) + "/" + text.Substring(2, 2) + "/" + text.Substring(4, 4);
						}
						else if (text.Trim().Length == 10)
						{
							result = text.Trim();
						}
						break;
					case "96":
						text = sView.Cells[iRow - 1, iCol - 1].Value.ToString();
						if (text.Trim().Length == 6)
						{
							result = text.Substring(0, 2) + ":" + text.Substring(2, 2) + ":" + text.Substring(4, 2);
						}
						else if (text.Trim().Length == 8)
						{
							result = text.Trim();
						}
						break;
					case "94":
						text = sView.Cells[iRow - 1, iCol - 1].Value.ToString();
						if (text.Trim().Length == 14)
						{
							result = text.Substring(0, 2) + "/" + text.Substring(2, 2) + "/" + text.Substring(4, 4) + " " + text.Substring(8, 2) + ":" + text.Substring(10, 2) + ":" + text.Substring(12, 2);
						}
						else if (text.Trim().Length == 18)
						{
							result = text.Trim();
						}
						break;
					case "0":
						result = sView.Cells[iRow - 1, iCol - 1].Text;
						break;
					case "16":
						if (sView.Cells[iRow - 1, iCol - 1].Value != null)
						{
							return sView.Cells[iRow - 1, iCol - 1].Value.ToString();
						}
						return "";
					default:
					{
						ICellType cellType = sView.Cells[iRow - 1, iCol - 1].CellType;
						if (cellType.ToString() == "GeneralCellType")
						{
							GeneralCellType generalCellType = (GeneralCellType)sView.Cells[iRow - 1, iCol - 1].CellType;
							if (generalCellType.NumberFormat != null)
							{
								if (sView.Cells[iRow - 1, iCol - 1].Value != null)
								{
									return sView.Cells[iRow - 1, iCol - 1].Value.ToString();
								}
								return "";
							}
							return sView.Cells[iRow - 1, iCol - 1].Text.ToString().Replace("&", "&amp;").Replace("\"", "&quot;")
								.Replace(">", "&gt;")
								.Replace("<", "&lt;");
						}
						if (cellType.ToString() == "NumberCellType")
						{
							if (sView.Cells[iRow - 1, iCol - 1].Value != null)
							{
								return sView.Cells[iRow - 1, iCol - 1].Value.ToString();
							}
							return "";
						}
						if (cellType.ToString() == "ButtonCellType")
						{
							ButtonCellType buttonCellType = (ButtonCellType)cellType;
							if (buttonCellType.Text != null)
							{
								return "TaiBangKe" + buttonCellType.Text.ToString();
							}
							return "";
						}
						return sView.Cells[iRow - 1, iCol - 1].Text.Replace("&", "&amp;").Replace("\"", "&quot;").Replace(">", "&gt;")
							.Replace("<", "&lt;");
					}
					}
				}
				else
				{
					result = "";
				}
			}
			else
			{
				result = string.Empty;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = string.Empty;
		}
		return result;
	}

	public static string GetValueCustomMV(SheetView sView, int iRow, int iCol, string controlType)
	{
		string result = string.Empty;
		try
		{
			if (iCol >= 0 && iRow >= 0)
			{
				if (sView.Cells[iRow - 1, iCol - 1].Value != null)
				{
					string text = "";
					switch (controlType)
					{
					case "2":
						text = sView.Cells[iRow - 1, iCol - 1].Value.ToString();
						result = ((!(text.ToLower() == "true")) ? "" : "x");
						break;
					case "14":
						text = sView.Cells[iRow - 1, iCol - 1].Value.ToString();
						if (text.Trim().Length == 8)
						{
							result = text.Substring(0, 2) + "/" + text.Substring(2, 2) + "/" + text.Substring(4, 4);
						}
						else if (text.Trim().Length == 10)
						{
							result = text.Trim();
						}
						break;
					case "0":
						result = sView.Cells[iRow - 1, iCol - 1].Text;
						break;
					case "16":
						if (sView.Cells[iRow - 1, iCol - 1].Value != null)
						{
							return sView.Cells[iRow - 1, iCol - 1].Value.ToString();
						}
						return "";
					default:
					{
						ICellType cellType = sView.Cells[iRow - 1, iCol - 1].CellType;
						if (cellType.ToString() == "GeneralCellType")
						{
							GeneralCellType generalCellType = (GeneralCellType)sView.Cells[iRow - 1, iCol - 1].CellType;
							if (generalCellType.NumberFormat != null)
							{
								if (sView.Cells[iRow - 1, iCol - 1].Value != null)
								{
									return sView.Cells[iRow - 1, iCol - 1].Value.ToString();
								}
								return "";
							}
							return sView.Cells[iRow - 1, iCol - 1].Text.ToString().Replace("&", "&amp;").Replace("'", "&apos;")
								.Replace("\"", "&quot;")
								.Replace(">", "&gt;")
								.Replace("<", "&lt;");
						}
						if (cellType.ToString() == "NumberCellType")
						{
							if (sView.Cells[iRow - 1, iCol - 1].Value != null)
							{
								return sView.Cells[iRow - 1, iCol - 1].Value.ToString();
							}
							return "";
						}
						return sView.Cells[iRow - 1, iCol - 1].Text.Replace("&", "&amp;").Replace("'", "&apos;").Replace("\"", "&quot;")
							.Replace(">", "&gt;")
							.Replace("<", "&lt;");
					}
					}
				}
				else
				{
					result = "";
				}
			}
			else
			{
				result = string.Empty;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = string.Empty;
		}
		return result;
	}

	public static string[] GetCatalogueFileName(DateTime dtPeriod)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Expected O, but got Unknown
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0045: Expected O, but got Unknown
		try
		{
			string[] array = new string[2];
			string text = "Menu.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNode val2 = ((XmlNode)val).SelectSingleNode("Root/Menu[@ID='100_7']");
			foreach (XmlNode item in val2)
			{
				XmlNode val3 = item;
				DateTime result = default(DateTime);
				string value = ((XmlNode)val3.Attributes["StartDate"]).Value;
				if (DateTime.TryParseExact(value, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out result) && dtPeriod >= result)
				{
					array[0] = ((XmlNode)val3.ChildNodes[0].Attributes["InterfaceIni"]).Value;
					array[1] = ((XmlNode)val3.ChildNodes[1].Attributes["InterfaceIni"]).Value;
				}
			}
			return array;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string GetValueExportXML(SheetView sView, string controlType, int iRow, int iCol, string defauValue, string sheetName = "")
	{
		string text = string.Empty;
		try
		{
			text = GetValueExport(sView, controlType, iRow, iCol, defauValue);
			if (controlType == "16" && Convert.ToDouble(text) == 0.0)
			{
				text = "0";
			}
			if (controlType == "16" && (sheetName == "01-KHBS" || sheetName == "01-1/KHBS"))
			{
				int digits = ((defauValue.Length > 2) ? (defauValue.Length - 2) : 0);
				text = Math.Round(Convert.ToDouble(text), digits).ToString();
				if (Math.Abs(ConvertHelpers.ConvertToDouble(Math.Round(Convert.ToDouble(text), digits))) < 0.0001 && Math.Abs(ConvertHelpers.ConvertToDouble(Math.Round(Convert.ToDouble(text), digits))) != 0.0)
				{
					text = $"{Math.Round(Convert.ToDouble(text), digits):F6}";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return text;
		}
		return text;
	}

	public static string GetValueExport(SheetView sView, string controlType, int iRow, int iCol, string defauValue)
	{
		try
		{
			int num = 0;
			if (iRow == 32 && iCol == HTKKFile.ColumnLettersToIndex("E"))
			{
				num = iRow;
			}
			if (sView.Cells[iRow - 1, iCol - 1].Value != null && HTKKFormater.isNullString(sView.Cells[iRow - 1, iCol - 1].Value.ToString()))
			{
				if (defauValue == null)
				{
					return "";
				}
				if (controlType == "16")
				{
					string text = defauValue;
					return text.Replace(".", "").Replace(",", ".");
				}
			}
			if (iCol >= 0 && iRow >= 0)
			{
				switch (controlType)
				{
				default:
					if (!(controlType == "102"))
					{
						if (string.IsNullOrEmpty(controlType))
						{
							controlType = "0";
						}
						switch (controlType)
						{
						case "2":
						{
							string text3 = ((sView.Cells[iRow - 1, iCol - 1].Value != null) ? sView.Cells[iRow - 1, iCol - 1].Value.ToString() : string.Empty);
							if (text3 == "True" || text3 == "1")
							{
								return "1";
							}
							return "0";
						}
						case "14":
							if (sView.Cells[iRow - 1, iCol - 1].Value != null && !string.IsNullOrEmpty(sView.Cells[iRow - 1, iCol - 1].Text))
							{
								DateTime dateTime = DateTime.ParseExact(sView.Cells[iRow - 1, iCol - 1].Text, "dd/MM/yyyy", null);
								return dateTime.Year + "-" + dateTime.Month.ToString().PadLeft(2, '0') + "-" + dateTime.Day.ToString().PadLeft(2, '0');
							}
							return "";
						case "0":
							return sView.Cells[iRow - 1, iCol - 1].Text;
						case "16":
						{
							string text = sView.Cells[iRow - 1, iCol - 1].Text;
							if (!HTKKFormater.isNullString(text))
							{
								text = ((defauValue == null || defauValue.IndexOf(",") <= 0) ? text.Replace(".", "").Replace(",", "") : text.Replace(".", "").Replace(",", "."));
								return text.StartsWith("(") ? ("-" + text.Replace("(", "").Replace(")", "")) : text;
							}
							object value = sView.Cells[iRow - 1, iCol - 1].Value;
							if (value == null || value.ToString() == "")
							{
								return "0";
							}
							return value.ToString();
						}
						case "33":
						{
							string text2 = sView.Cells[iRow - 1, iCol - 1].Value.ToString();
							return text2.Equals("0") ? "" : text2;
						}
						default:
						{
							ICellType cellType = sView.Cells[iRow - 1, iCol - 1].CellType;
							switch (cellType.ToString())
							{
							case "GeneralCellType":
							{
								GeneralCellType generalCellType = (GeneralCellType)sView.Cells[iRow - 1, iCol - 1].CellType;
								if (generalCellType.NumberFormat != null)
								{
									if (sView.Cells[iRow - 1, iCol - 1].Value != null)
									{
										return sView.Cells[iRow - 1, iCol - 1].Value.ToString();
									}
									return "";
								}
								if (generalCellType.DateTimeFormat != null)
								{
									DateTime dateTime = (DateTime)sView.Cells[iRow - 1, iCol - 1].Value;
									return dateTime.Year + "-" + dateTime.Month + "-" + dateTime.Day;
								}
								return sView.Cells[iRow - 1, iCol - 1].Text.ToString();
							}
							case "NumberCellType":
								if (sView.Cells[iRow - 1, iCol - 1].Value != null)
								{
									return sView.Cells[iRow - 1, iCol - 1].Value.ToString();
								}
								return "";
							case "DateTimeCellType":
								if (sView.Cells[iRow - 1, iCol - 1].Value != null)
								{
									DateTime dateTime = (DateTime)sView.Cells[iRow - 1, iCol - 1].Value;
									return dateTime.Year + "-" + dateTime.Month.ToString().PadLeft(2, '0') + "-" + dateTime.Day.ToString().PadLeft(2, '0');
								}
								return "";
							default:
								return sView.Cells[iRow - 1, iCol - 1].Text;
							}
						}
						}
					}
					goto case "4";
				case "4":
				case "5":
				case "100":
				case "6":
				case "7":
				case "101":
					return sView.Cells[iRow - 1, iCol - 1].Text;
				}
			}
			return string.Empty;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetValueImport(SheetView sView, string controlType, int iRow, int iCol, string defauValue)
	{
		try
		{
			int num = 0;
			if (iRow == 28)
			{
				num = iRow;
			}
			if (HTKKFormater.isNullString(sView.Cells[iRow - 1, iCol - 1].Text))
			{
				return defauValue;
			}
			if (iCol >= 0 && iRow >= 0)
			{
				switch (controlType)
				{
				default:
					if (!(controlType == "102"))
					{
						if (string.IsNullOrEmpty(controlType))
						{
							controlType = "0";
						}
						switch (controlType)
						{
						case "2":
						{
							string text = sView.Cells[iRow - 1, iCol - 1].Value.ToString();
							if (text == "True" || text.ToUpper() == "X")
							{
								return "1";
							}
							return "0";
						}
						case "14":
							if (sView.Cells[iRow - 1, iCol - 1].Value != null)
							{
								return sView.Cells[iRow - 1, iCol - 1].Value.ToString();
							}
							return "";
						case "0":
							return sView.Cells[iRow - 1, iCol - 1].Text;
						case "16":
						{
							string text2 = sView.Cells[iRow - 1, iCol - 1].Value.ToString();
							text2 = ((text2.IndexOf(",") <= 0) ? text2.Replace(".", "") : text2.Replace(".", "").Replace(",", "."));
							return text2.StartsWith("(") ? ("-" + text2.Replace("(", "").Replace(")", "")) : text2;
						}
						case "33":
						{
							string text3 = sView.Cells[iRow - 1, iCol - 1].Value.ToString();
							return text3.Equals("0") ? "" : text3;
						}
						default:
							if (sView.Cells[iRow - 1, iCol - 1].Value != null)
							{
								return sView.Cells[iRow - 1, iCol - 1].Value.ToString();
							}
							return "";
						}
					}
					goto case "4";
				case "4":
				case "5":
				case "100":
				case "6":
				case "7":
				case "101":
					return sView.Cells[iRow - 1, iCol - 1].Text;
				}
			}
			return string.Empty;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static object GetObjectValue(SheetView sView, string index)
	{
		object result = null;
		try
		{
			if (index.Length > 0)
			{
				string[] array = index.Split(new char[1] { '_' });
				if (array.Length == 2)
				{
					int num = -1;
					int num2 = -1;
					num2 = int.Parse(array[1]) - 1;
					num = HTKKFile.ColumnLettersToIndex(array[0]) - 1;
					result = ((num < 0 || num2 < 0) ? null : sView.Cells[num2, num].Value);
				}
				else
				{
					result = null;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = null;
		}
		return result;
	}

	public static int QuarterToFirstMonth(int quarter)
	{
		int result = 0;
		try
		{
			switch (quarter)
			{
			case 1:
				result = 1;
				break;
			case 2:
				result = 4;
				break;
			case 3:
				result = 7;
				break;
			case 4:
				result = 10;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static int QuarterToLastMonth(int quarter)
	{
		int result = 0;
		try
		{
			switch (quarter)
			{
			case 1:
				result = 3;
				break;
			case 2:
				result = 6;
				break;
			case 3:
				result = 9;
				break;
			case 4:
				result = 12;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static List<string> GetFileNameCheckThangQuy(string fileNameExtend)
	{
		List<string> list = new List<string>();
		string text = fileNameExtend.Substring(0, 6);
		string text2 = fileNameExtend.Substring(6, 4);
		switch (text)
		{
		case "Q_0001":
			list.Add("M_0001" + text2);
			list.Add("M_0002" + text2);
			list.Add("M_0003" + text2);
			return list;
		case "Q_0002":
			list.Add("M_0004" + text2);
			list.Add("M_0005" + text2);
			list.Add("M_0006" + text2);
			return list;
		case "Q_0003":
			list.Add("M_0007" + text2);
			list.Add("M_0008" + text2);
			list.Add("M_0009" + text2);
			return list;
		case "Q_0004":
			list.Add("M_0010" + text2);
			list.Add("M_0011" + text2);
			list.Add("M_0012" + text2);
			return list;
		case "M_0001":
			list.Add("Q_0001" + text2);
			return list;
		case "M_0002":
			list.Add("Q_0001" + text2);
			return list;
		case "M_0003":
			list.Add("Q_0001" + text2);
			return list;
		case "M_0004":
			list.Add("Q_0002" + text2);
			return list;
		case "M_0005":
			list.Add("Q_0002" + text2);
			return list;
		case "M_0006":
			list.Add("Q_0002" + text2);
			return list;
		case "M_0007":
			list.Add("Q_0003" + text2);
			return list;
		case "M_0008":
			list.Add("Q_0003" + text2);
			return list;
		case "M_0009":
			list.Add("Q_0003" + text2);
			return list;
		case "M_0010":
			list.Add("Q_0004" + text2);
			return list;
		case "M_0011":
			list.Add("Q_0004" + text2);
			return list;
		case "M_0012":
			list.Add("Q_0004" + text2);
			return list;
		default:
			return null;
		}
	}

	public static string GetTuNgay(string fileNameExtend)
	{
		if (fileNameExtend.Length > 20)
		{
			return fileNameExtend.Substring(15, 4) + fileNameExtend.Substring(13, 2) + fileNameExtend.Substring(11, 2);
		}
		return string.Empty;
	}

	public static string GetDenNgay(string fileNameExtend)
	{
		if (fileNameExtend.Length > 0)
		{
			return fileNameExtend.Substring(24, 4) + fileNameExtend.Substring(22, 2) + fileNameExtend.Substring(20, 2);
		}
		return string.Empty;
	}

	public static void ChangeLockedCell(SheetView sView, int rowIndex, int columnIndex, bool valueChange)
	{
		try
		{
			sView.Cells[rowIndex, columnIndex].Locked = valueChange;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static List<string> GetPreperiodFileName_ThangAndQuy(HTKKPeriodInfo periodInfo, HTKKCMCToKhaiInfo toKhaiInfo, int value)
	{
		try
		{
			List<string> list = new List<string>();
			HTKKCMCSheetInfo hTKKCMCSheetInfo = toKhaiInfo.SheetInfos[0];
			string versionXML = toKhaiInfo.VersionXML;
			string dataFile = hTKKCMCSheetInfo.DataFile;
			string thang = periodInfo.Thang;
			string quy = periodInfo.Quy;
			int num = Convert.ToInt32(periodInfo.Nam);
			switch (periodInfo.KieuKy)
			{
			case "M":
			{
				int num5 = Convert.ToInt32(thang);
				int num6 = ((num5 == 1) ? 12 : (num5 - 1));
				int num7 = ((num5 == 1) ? (num - 1) : num);
				if (value == 1)
				{
					num6++;
				}
				string text2 = string.Concat(num6, num7);
				list.Add(string.Format("{0}_{1}_{2}", dataFile, text2.PadLeft(6, '0'), versionXML.Replace(".", "")));
				break;
			}
			case "Q":
			{
				int num2 = Convert.ToInt32(quy);
				int num3 = ((num2 == 1) ? 4 : (num2 - 1));
				int num4 = ((num2 == 1) ? (num - 1) : num);
				if (value == 1)
				{
					num3++;
				}
				string text = string.Concat(num3, num4);
				list.Add(string.Format("{0}_Q{1}_{2}", dataFile, text.PadLeft(6, '0'), versionXML.Replace(".", "")));
				break;
			}
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw;
		}
	}

	public static List<string> GetPreperiodFileName_ThangQuy(HTKKPeriodInfo periodInfo, HTKKCMCToKhaiInfo toKhaiInfo)
	{
		try
		{
			List<string> list = new List<string>();
			HTKKCMCSheetInfo hTKKCMCSheetInfo = toKhaiInfo.SheetInfos[0];
			string versionXML = toKhaiInfo.VersionXML;
			string dataFile = hTKKCMCSheetInfo.DataFile;
			string text = string.Format("{0}_{1}_{2}", dataFile, "", versionXML);
			string thang = periodInfo.Thang;
			string quy = periodInfo.Quy;
			int num = Convert.ToInt32(periodInfo.Nam);
			switch (periodInfo.KieuKy)
			{
			case "M":
			{
				int num5 = Convert.ToInt32(thang);
				int num6 = ((num5 == 1) ? 12 : (num5 - 1));
				int num7 = ((num5 == 1) ? (num - 1) : num);
				string text3 = string.Concat(num6, num7);
				if (periodInfo.menuID.Equals("01") || periodInfo.menuID.Equals("I5"))
				{
					if (num >= 2021)
					{
						list.Add(string.Format("{0}_{1}_{2}_{3}", dataFile, periodInfo.TieuMucHachToanVal, text3.PadLeft(6, '0'), versionXML.Replace(".", "")));
					}
					else
					{
						list.Add(string.Format("{0}_{1}_{2}", dataFile, text3.PadLeft(6, '0'), versionXML.Replace(".", "")));
					}
				}
				else if (periodInfo.menuID.Equals("D1") || periodInfo.menuID.Equals("Z5"))
				{
					list.Add(string.Format("{0}_{1}_{2}_{3}_{4}", dataFile, GetNameNopTheoTinhCQT(toKhaiInfo, periodInfo), periodInfo.TieuMucHachToanVal, text3.PadLeft(6, '0'), versionXML.Replace(".", "")));
				}
				else if (periodInfo.menuID.Equals("02"))
				{
					if (num >= 2021)
					{
						list.Add(string.Format("{0}_{1}_{2}_{3}", dataFile, periodInfo.MaHSDuAnDT, text3.PadLeft(6, '0'), versionXML.Replace(".", "")));
					}
					else
					{
						list.Add(string.Format("{0}_{1}_{2}", dataFile, text3.PadLeft(6, '0'), versionXML.Replace(".", "")));
					}
				}
				else if (periodInfo.menuID.Equals("D2"))
				{
					list.Add(string.Format("{0}_{1}_{2}_{3}_{4}", dataFile, "CQTQL" + periodInfo.maCQTQuanLy, periodInfo.MaHSDuAnDT, text3.PadLeft(6, '0'), versionXML.Replace(".", "")));
				}
				else
				{
					list.Add(string.Format("{0}_{1}_{2}", dataFile, text3.PadLeft(6, '0'), versionXML.Replace(".", "")));
				}
				break;
			}
			case "Q":
			{
				int num2 = Convert.ToInt32(quy);
				int num3 = ((num2 == 1) ? 4 : (num2 - 1));
				int num4 = ((num2 == 1) ? (num - 1) : num);
				string text2 = string.Concat(num3, num4);
				if (periodInfo.menuID.Equals("01") || periodInfo.menuID.Equals("I5"))
				{
					if (num >= 2021)
					{
						list.Add(string.Format("{0}_{1}_Q{2}_{3}", dataFile, periodInfo.TieuMucHachToanVal, text2.PadLeft(6, '0'), versionXML.Replace(".", "")));
					}
					else
					{
						list.Add(string.Format("{0}_Q{1}_{2}", dataFile, text2.PadLeft(6, '0'), versionXML.Replace(".", "")));
					}
				}
				else if (periodInfo.menuID.Equals("D1") || periodInfo.menuID.Equals("Z5"))
				{
					list.Add(string.Format("{0}_{1}_{2}_Q{3}_{4}", dataFile, GetNameNopTheoTinhCQT(toKhaiInfo, periodInfo), periodInfo.TieuMucHachToanVal, text2.PadLeft(6, '0'), versionXML.Replace(".", "")));
				}
				else if (periodInfo.menuID.Equals("02"))
				{
					if (num >= 2021)
					{
						list.Add(string.Format("{0}_{1}_Q{2}_{3}", dataFile, periodInfo.MaHSDuAnDT, text2.PadLeft(6, '0'), versionXML.Replace(".", "")));
					}
					else
					{
						list.Add(string.Format("{0}_Q{1}_{2}", dataFile, text2.PadLeft(6, '0'), versionXML.Replace(".", "")));
					}
				}
				else if (periodInfo.menuID.Equals("D2"))
				{
					list.Add(string.Format("{0}_{1}_{2}_Q{3}_{4}", dataFile, "CQTQL" + periodInfo.maCQTQuanLy, periodInfo.MaHSDuAnDT, text2.PadLeft(6, '0'), versionXML.Replace(".", "")));
				}
				else
				{
					list.Add(string.Format("{0}_Q{1}_{2}", dataFile, text2.PadLeft(6, '0'), versionXML.Replace(".", "")));
				}
				break;
			}
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw;
		}
	}

	public static string GetPreperiodFileName_Quy(HTKKPeriodInfo periodInfo, HTKKCMCToKhaiInfo toKhaiInfo)
	{
		try
		{
			string text = "";
			HTKKCMCSheetInfo hTKKCMCSheetInfo = toKhaiInfo.SheetInfos[0];
			string versionXML = toKhaiInfo.VersionXML;
			string dataFile = hTKKCMCSheetInfo.DataFile;
			int num;
			int num4;
			string text2;
			if (periodInfo.KieuKy == "Q")
			{
				string quy = periodInfo.Quy;
				num = ConvertHelpers.ConvertToInt(periodInfo.Nam);
				int num2 = ConvertHelpers.ConvertToInt(quy);
				int num3 = ((num2 == 1) ? 4 : (num2 - 1));
				num4 = ((num2 == 1) ? (num - 1) : num);
				text2 = string.Concat(num3, num4);
				return string.Format("{0}_Q{1}_{2}", dataFile, text2.PadLeft(6, '0'), versionXML.Replace(".", ""));
			}
			int num5 = ConvertHelpers.ConvertToInt(periodInfo.Thang);
			num = ConvertHelpers.ConvertToInt(periodInfo.Nam);
			int num6 = ((num5 == 1) ? 12 : (num5 - 1));
			num4 = ((num5 == 1) ? (num - 1) : num);
			text2 = string.Concat(num6, num4);
			return string.Format("{0}_{1}_{2}", dataFile, text2.PadLeft(6, '0'), versionXML.Replace(".", ""));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw;
		}
	}

	public static DateTime getDateKeKhai(HTKKPeriodInfo periodInfo)
	{
		try
		{
			int num = 1;
			int year = DateTime.Now.Year;
			if (ConvertHelpers.ConvertToInt(periodInfo.Nam) > 0)
			{
				year = int.Parse(periodInfo.Nam);
			}
			num = periodInfo.KieuKy switch
			{
				"M" => int.Parse(periodInfo.Thang), 
				"Q" => HTKKFormater.isNullString(periodInfo.Quy) ? 1 : (int.Parse(periodInfo.Quy) * 3), 
				"D" => int.Parse(periodInfo.Thang), 
				_ => 1, 
			};
			if (num == 0)
			{
				num = 1;
			}
			return new DateTime(year, num, 1);
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	public static DateTime getDateEffectByAttributes(XmlNode xMenuNode, string AttributesName)
	{
		DateTime result = new DateTime(1900, 1, 1);
		try
		{
			XmlNode val = xMenuNode.SelectSingleNode("Validity");
			if (val.Attributes[AttributesName] != null)
			{
				if (((XmlNode)val.Attributes[AttributesName]).Value.Length > 0)
				{
					result = DateTime.ParseExact(((XmlNode)val.Attributes[AttributesName]).Value, "d/M/yyyy", null);
					return result;
				}
				result = new DateTime(2100, 1, 1);
				return result;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static bool checkVersionDateEffect(ref XmlNode xMenuNode, HTKKPeriodInfo TTPeriodInfo, XmlNodeList menuList)
	{
		//IL_0068: Unknown result type (might be due to invalid IL or missing references)
		//IL_006f: Expected O, but got Unknown
		bool result = false;
		try
		{
			if (menuList.Count == 0)
			{
				return false;
			}
			DateTime dateEffectByAttributes = getDateEffectByAttributes(xMenuNode, "StartDate");
			DateTime dateEffectByAttributes2 = getDateEffectByAttributes(xMenuNode, "EndDate");
			DateTime dateKeKhai = getDateKeKhai(TTPeriodInfo);
			if (DateTime.Compare(dateKeKhai, dateEffectByAttributes) < 0)
			{
				foreach (XmlNode menu in menuList)
				{
					XmlNode val = menu;
					dateEffectByAttributes = getDateEffectByAttributes(val, "StartDate");
					dateEffectByAttributes2 = getDateEffectByAttributes(val, "EndDate");
					if (dateKeKhai >= dateEffectByAttributes && dateKeKhai <= dateEffectByAttributes2)
					{
						xMenuNode = val;
						result = true;
						break;
					}
				}
			}
		}
		catch (Exception)
		{
			result = false;
		}
		return result;
	}

	public static bool checkVersionEffectByVerXml(ref XmlNode xMenuNode, XmlNodeList menuList, string versionXml)
	{
		//IL_006f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0075: Expected O, but got Unknown
		bool result = false;
		try
		{
			string text = ((xMenuNode.Attributes["VersionXML"] == null) ? "" : ((XmlNode)xMenuNode.Attributes["VersionXML"]).Value.Replace(".", ""));
			if (versionXml == text)
			{
				return false;
			}
			foreach (XmlNode menu in menuList)
			{
				XmlNode val = menu;
				string text2 = ((val.Attributes["VersionXML"] == null) ? "" : ((XmlNode)val.Attributes["VersionXML"]).Value.Replace(".", ""));
				if (versionXml == text2)
				{
					xMenuNode = val;
					result = true;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static string GetValueInXml(string filename, string cellID)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		try
		{
			string result = "";
			XmlDocument val = new XmlDocument();
			val.Load(filename);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			for (int i = 0; i < elementsByTagName.Count; i++)
			{
				string value = ((XmlNode)elementsByTagName[i].Attributes["CellID"]).Value;
				if (value == cellID)
				{
					result = ((XmlNode)elementsByTagName[i].Attributes["Value"]).Value;
					break;
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

	public static Dictionary<string, string> GetValueInXml(string filename, List<string> lstCellId)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		try
		{
			XmlDocument val = new XmlDocument();
			val.Load(filename);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			for (int i = 0; i < elementsByTagName.Count; i++)
			{
				string value = ((XmlNode)elementsByTagName[i].Attributes["CellID"]).Value;
				if (lstCellId.Contains(value))
				{
					dictionary.Add(value, ((XmlNode)elementsByTagName[i].Attributes["Value"]).Value);
					lstCellId.Remove(value);
				}
				if (lstCellId.Count == 0)
				{
					break;
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return dictionary;
		}
	}

	public static List<Dictionary<string, string>> GetRowDataValueInXml(string filename, List<string> lstCellId)
	{
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0022: Expected O, but got Unknown
		List<Dictionary<string, string>> list = new List<Dictionary<string, string>>();
		try
		{
			if (!File.Exists(filename))
			{
				return null;
			}
			XmlDocument val = new XmlDocument();
			val.Load(filename);
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("Sections/Section/RowData");
			for (int i = 0; i < val2.Count; i++)
			{
				string text = "###";
				string[] array = ConvertHelpers.ConvertToString(val2[i].InnerXml).Split(new string[1] { text }, StringSplitOptions.None);
				Dictionary<string, string> dictionary = new Dictionary<string, string>();
				for (int j = 0; j < array.Length; j++)
				{
					dictionary.Add(lstCellId[j], array[j]);
				}
				list.Add(dictionary);
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return list;
		}
	}

	public static string GetValAttributeStaticCell(HTKKCMCSheetInfo sheetInfo, string cellID, string attributeName)
	{
		string result = string.Empty;
		try
		{
			int[] array = ArrayCellId(cellID);
			foreach (HTKKCMCCellInfo staticCell in sheetInfo.StaticCells)
			{
				if ((staticCell.CellTemplate.ColIndex - 1 == array[0]) & (staticCell.CellTemplate.RowIndex - 1 == array[1]))
				{
					switch (attributeName)
					{
					case "Path":
						result = staticCell.Path;
						break;
					case "Value":
						result = staticCell.Value;
						break;
					}
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static int[] ArrayCellId(string cellId)
	{
		try
		{
			int[] array = new int[2];
			string[] array2 = cellId.Split(new char[1] { '_' });
			array[0] = HTKKFile.ColumnLettersToIndex(array2[0]) - 1;
			array[1] = int.Parse(array2[1]) - 1;
			return array;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return new int[2] { -1, -1 };
		}
	}

	public static bool HaveToValidateRow(HTKKCMCRowInfo rowsInfo, string[] lstRowValue)
	{
		try
		{
			bool flag = false;
			for (int i = 0; i < rowsInfo.Cells.Count; i++)
			{
				string text = ((rowsInfo.Cells[i].DefaultValue == null) ? "" : rowsInfo.Cells[i].DefaultValue);
				if (i < lstRowValue.Length && !text.Equals("NA"))
				{
					if (rowsInfo.Cells[i].Controltype == "16")
					{
						text = ConvertHelpers.ConvertToDouble(rowsInfo.Cells[i].DefaultValue).ToString();
						lstRowValue[i] = ConvertHelpers.ConvertToDouble(lstRowValue[i]).ToString();
					}
					if (rowsInfo.Cells[i].Controltype == "2")
					{
						text = "False";
					}
					if (!string.IsNullOrEmpty(lstRowValue[i]) && text != null && text != lstRowValue[i])
					{
						flag = true;
						break;
					}
				}
			}
			if (!flag)
			{
			}
			return flag;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static bool HaveToValidateHeaderRow(HTKKCMCStaticOfTable header, string[] lstRowValue)
	{
		try
		{
			string text = "";
			bool result = false;
			for (int i = 0; i < header.StaticCells.Count; i++)
			{
				string text2 = header.StaticCells[i].DefaultValue;
				text = lstRowValue[i];
				if (header.StaticCells[i].Controltype == "16")
				{
					text2 = ConvertHelpers.ConvertToDouble(header.StaticCells[i].DefaultValue).ToString();
					text = ConvertHelpers.ConvertToDouble(text).ToString();
				}
				if (!string.IsNullOrEmpty(text) && text2 != null && text2 != text)
				{
					result = true;
					break;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	public static bool HaveToValidateHeaderRow(SheetView sView, HTKKCMCStaticOfTable header, int odrStt, string[] lstRowValue)
	{
		try
		{
			bool result = false;
			for (int i = 0; i < header.StaticCells.Count; i++)
			{
				ICellType cellType = sView.Cells[header.StaticCells[i].CellTemplate.RowIndex - 1, header.StaticCells[i].CellTemplate.ColIndex - 1].CellType;
				if (cellType.ToString() == "NumberCellType" || header.StaticCells[i].Controltype == "16")
				{
					lstRowValue[i] = (string.IsNullOrEmpty(lstRowValue[i]) ? "" : ConvertHelpers.ConvertToDouble(lstRowValue[i]).ToString());
				}
				if (header.StaticCells[i].CellTemplate.ColIndex != odrStt && !string.IsNullOrEmpty(lstRowValue[i]) && header.StaticCells[i].DefaultValue != null && header.StaticCells[i].DefaultValue != lstRowValue[i])
				{
					result = true;
					break;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	public static string[] GetHeaderRowValue(SheetView sView, HTKKCMCStaticOfTable header)
	{
		try
		{
			string[] array = new string[header.StaticCells.Count];
			for (int i = 0; i < header.StaticCells.Count; i++)
			{
				ICellType cellType = sView.Cells[header.StaticCells[i].CellTemplate.RowIndex - 1, header.StaticCells[i].CellTemplate.ColIndex - 1].CellType;
				string text = sView.Cells[header.StaticCells[i].CellTemplate.RowIndex - 1, header.StaticCells[i].CellTemplate.ColIndex - 1].Text;
				if (cellType.ToString() == "NumberCellType")
				{
					text = (string.IsNullOrEmpty(text) ? "" : ConvertHelpers.ConvertToDouble(text).ToString());
				}
				array[i] = text;
			}
			return array;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	public static string[] GetRowValue(SheetView sView, HTKKCMCRowInfo rowsInfo, int odrDataRowFirst)
	{
		try
		{
			string[] array = new string[rowsInfo.Cells.Count];
			for (int i = 0; i < rowsInfo.Cells.Count; i++)
			{
				ICellType cellType = sView.Cells[odrDataRowFirst - 1, rowsInfo.Cells[i].CellTemplate.ColIndex - 1].CellType;
				string text = sView.Cells[odrDataRowFirst - 1, rowsInfo.Cells[i].CellTemplate.ColIndex - 1].Text;
				if (cellType.ToString() == "NumberCellType")
				{
					text = (string.IsNullOrEmpty(text) ? "" : ConvertHelpers.ConvertToDouble(text).ToString());
				}
				array[i] = text;
			}
			return array;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static int SoNgayNopCham(HTKKPeriodInfo periodInfo, HTKKHeaderInfo HeaderInfo = null)
	{
		int result = 0;
		DateTime hanNopToKhai = GetHanNopToKhai(periodInfo, HeaderInfo);
		try
		{
			if (string.IsNullOrEmpty(periodInfo.NgayLapKHBS))
			{
				return result;
			}
			DateTime dateTime = DateTime.ParseExact(periodInfo.NgayLapKHBS, "d/M/yyyy", null);
			result = (dateTime - hanNopToKhai).Days;
			if (result < 0)
			{
				result = 0;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		return result;
	}

	public static DateTime GetHanNopToKhai(HTKKPeriodInfo periodInfo, HTKKHeaderInfo headerInfo = null)
	{
		//IL_15e2: Unknown result type (might be due to invalid IL or missing references)
		//IL_15e9: Expected O, but got Unknown
		//IL_1617: Unknown result type (might be due to invalid IL or missing references)
		//IL_161e: Expected O, but got Unknown
		try
		{
			DateTime dateTime = default(DateTime);
			bool flag = false;
			int num = 90;
			int num2 = 45;
			int num3 = 10;
			int num4 = int.Parse(periodInfo.Nam);
			switch (periodInfo.KieuKy)
			{
			case "M":
			{
				int num8 = int.Parse(periodInfo.Thang);
				if (periodInfo.menuID.Equals("01") || periodInfo.menuID.Equals("I5"))
				{
					if ((ConvertHelpers.ConvertToInt(periodInfo.Thang) != 4 && ConvertHelpers.ConvertToInt(periodInfo.Thang) != 5 && ConvertHelpers.ConvertToInt(periodInfo.Thang) != 6) || ConvertHelpers.ConvertToInt(periodInfo.Nam) != 2012 || !flag)
					{
						dateTime = ((ConvertHelpers.ConvertToInt(periodInfo.Thang) != 12) ? new DateTime(num4, num8 + 1, 20) : new DateTime(num4 + 1, 1, 20));
					}
					else if (ConvertHelpers.ConvertToInt(periodInfo.Thang) == 4)
					{
						dateTime = new DateTime(num4, 11, 20);
					}
					else if (ConvertHelpers.ConvertToInt(periodInfo.Thang) == 5)
					{
						dateTime = new DateTime(num4, 12, 20);
					}
					else if (ConvertHelpers.ConvertToInt(periodInfo.Thang) == 6)
					{
						dateTime = new DateTime(num4 + 1, 1, 21);
					}
				}
				else
				{
					dateTime = ((ConvertHelpers.ConvertToInt(periodInfo.Thang) != 12) ? ((!periodInfo.menuID.Equals("Q9") && !periodInfo.menuID.Equals("R4") && !periodInfo.menuID.Equals("X8") && !periodInfo.menuID.Equals("Q8") && !periodInfo.menuID.Equals("U5")) ? new DateTime(num4, num8 + 1, 20) : ((num4 != 2026 || (num8 != 1 && num8 != 2 && num8 != 3)) ? new DateTime(num4, num8 + 1, 20) : new DateTime(num4, 4, 20))) : new DateTime(num4 + 1, 1, 20));
				}
				break;
			}
			case "Q":
				if (ConvertHelpers.ConvertToInt(periodInfo.Quy) == 4)
				{
					dateTime = new DateTime(num4 + 1, 1, 30);
				}
				else if (ConvertHelpers.ConvertToInt(periodInfo.Quy) == 3)
				{
					dateTime = new DateTime(num4, 10, 30);
				}
				else if (ConvertHelpers.ConvertToInt(periodInfo.Quy) == 2)
				{
					dateTime = new DateTime(num4, 7, 30);
				}
				else if (ConvertHelpers.ConvertToInt(periodInfo.Quy) == 1)
				{
					dateTime = new DateTime(num4, 5, 1);
				}
				else if (periodInfo.menuID.Equals("30"))
				{
					DateTime minValue = DateTime.MinValue;
					string tuNgay = periodInfo.TuNgay;
					string[] array3 = tuNgay.Split(new char[1] { '/' });
					dateTime = new DateTime(ConvertHelpers.ConvertToInt(array3[2]), ConvertHelpers.ConvertToInt(array3[1]), ConvertHelpers.ConvertToInt(array3[0])).AddDays(10.0);
				}
				break;
			case "Y":
			{
				string tuThang = periodInfo.TuThang;
				string denThang = periodInfo.DenThang;
				string tuNgay2 = periodInfo.TuNgay;
				string denNgay = periodInfo.DenNgay;
				int num9 = 0;
				DateTime dateTime2 = DateTime.MinValue;
				if (!string.IsNullOrEmpty(tuThang) && !string.IsNullOrEmpty(denThang))
				{
					num9 = GetSoThang(tuThang.Replace("/", ""), denThang.Replace("/", ""));
					string[] array4 = denThang.Split(new char[1] { '/' });
					int num10 = ConvertHelpers.ConvertToInt(array4[0]);
					if (num9 < 12 && num10 != 12)
					{
						dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(array4[1]), ConvertHelpers.ConvertToInt(array4[0]), 1).AddMonths(1).AddDays(-1.0);
						num = 45;
					}
				}
				if (!string.IsNullOrEmpty(tuNgay2) && !string.IsNullOrEmpty(denNgay))
				{
					num9 = GetSoThang(tuNgay2.Substring(3).Replace("/", ""), denNgay.Substring(3).Replace("/", ""));
					string[] array5 = denNgay.Split(new char[1] { '/' });
					string text7 = array5[0] + "/" + array5[1];
					if (num9 >= 12 || text7 == "31/12")
					{
						dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(array5[2]), ConvertHelpers.ConvertToInt(array5[1]), ConvertHelpers.ConvertToInt(array5[0]));
					}
					else
					{
						dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(array5[2]), ConvertHelpers.ConvertToInt(array5[1]), ConvertHelpers.ConvertToInt(array5[0]));
						num = 45;
					}
				}
				if (periodInfo.menuID.Equals("03") || periodInfo.menuID.Equals("77"))
				{
					if (num == 90)
					{
						string text8 = null;
						int iThangTc = ((headerInfo == null) ? 1 : Convert.ToInt32(headerInfo.NgayBatDauNamTaiChinh.Split(new char[1] { '/' })[1]));
						int iNgayTc = ((headerInfo == null) ? 1 : Convert.ToInt32(headerInfo.NgayBatDauNamTaiChinh.Split(new char[1] { '/' })[0]));
						if (periodInfo.dicVar.ContainsKey("NGAY_BAT_DAU_NAM_TAI_CHINH"))
						{
							text8 = (string)periodInfo.dicVar["NGAY_BAT_DAU_NAM_TAI_CHINH"];
						}
						if (!HTKKFormater.isNullString(text8))
						{
							iThangTc = ConvertHelpers.ConvertToInt(text8.Split(new char[1] { '/' })[1]);
							iNgayTc = ConvertHelpers.ConvertToInt(text8.Split(new char[1] { '/' })[0]);
						}
						dateTime = NgayCuoiNamTaiChinh(num4, iThangTc, iNgayTc);
						DateTime dateTime3 = new DateTime(2020, 6, 1);
						if (dateTime2 >= dateTime3)
						{
							DateTime dateTime4 = dateTime.AddMonths(3);
							dateTime = new DateTime(dateTime4.Year, dateTime4.Month, 1).AddMonths(1).AddDays(-1.0);
						}
						else
						{
							dateTime = dateTime.AddDays(90.0);
						}
					}
					else
					{
						dateTime = dateTime2.AddDays(num);
					}
				}
				else if (periodInfo.menuID.Equals("T5") || periodInfo.menuID.Equals("Y4") || periodInfo.menuID.Equals("C6") || periodInfo.menuID.Equals("C7") || periodInfo.menuID.Equals("N8") || periodInfo.menuID.Equals("H4"))
				{
					DateTime dateTime5 = DateTime.ParseExact(periodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					if (periodInfo.MaQuyetToan != "01" && periodInfo.MaQuyetToan != "04")
					{
						dateTime = dateTime5.AddDays(num2).AddDays(-1.0);
						break;
					}
					int iThangTc = ((headerInfo == null) ? 1 : Convert.ToInt32(headerInfo.NgayBatDauNamTaiChinh.Split(new char[1] { '/' })[1]));
					int iNgayTc = ((headerInfo == null) ? 1 : Convert.ToInt32(headerInfo.NgayBatDauNamTaiChinh.Split(new char[1] { '/' })[0]));
					List<string> quyByNamTC = GetQuyByNamTC(dateTime5.Month, dateTime5.Year, iNgayTc + "/" + iThangTc);
					if (quyByNamTC[0].Equals("4"))
					{
						quyByNamTC[0] = "1";
						quyByNamTC[1] = Convert.ToString(Convert.ToInt32(quyByNamTC[1]) + 1);
					}
					else
					{
						quyByNamTC[0] = Convert.ToString(Convert.ToInt32(quyByNamTC[0]) + 1);
					}
					dateTime = GetNgayCuoiQuy(Convert.ToInt32(quyByNamTC[0]), Convert.ToInt32(quyByNamTC[1]), iNgayTc, iThangTc);
				}
				else if (periodInfo.menuID.Equals("83"))
				{
					if (num == 90)
					{
						string text8 = null;
						int iThangTc = 1;
						int iNgayTc = 1;
						if (periodInfo.dicVar.ContainsKey("NGAY_BAT_DAU_NAM_TAI_CHINH"))
						{
							text8 = (string)periodInfo.dicVar["NGAY_BAT_DAU_NAM_TAI_CHINH"];
						}
						if (!HTKKFormater.isNullString(text8))
						{
							iThangTc = ConvertHelpers.ConvertToInt(text8.Split(new char[1] { '/' })[1]);
							iNgayTc = ConvertHelpers.ConvertToInt(text8.Split(new char[1] { '/' })[0]);
						}
						dateTime = NgayCuoiNamTaiChinh(num4, iThangTc, iNgayTc).AddDays(90.0);
					}
					else
					{
						dateTime = dateTime2.AddDays(num);
					}
				}
				else if (periodInfo.menuID.Equals("A3"))
				{
					dateTime = new DateTime(num4 - 1, 12, 31);
				}
				else if (periodInfo.menuID.Equals("63"))
				{
					dateTime = new DateTime(num4, 1, 1);
				}
				else if (periodInfo.menuID.Equals("L1"))
				{
					bool flag15 = false;
					if (periodInfo.dicVar.Count > 0)
					{
						flag15 = Convert.ToBoolean(periodInfo.dicVar["CoSoMoiThanhLap"]);
					}
					dateTime = ((!flag15) ? new DateTime(num4, 1, 30) : new DateTime(ConvertHelpers.ConvertToInt(num4) + 1, 1, 30));
				}
				else
				{
					dateTime = ((num != 90) ? dateTime2.AddDays(num) : new DateTime(num4, 12, 31).AddDays(num));
				}
				break;
			}
			case "D":
			{
				int month = int.Parse(periodInfo.Thang);
				int day = int.Parse(periodInfo.Ngay);
				bool flag2 = false;
				bool flag3 = false;
				bool flag4 = false;
				bool flag5 = false;
				bool flag6 = false;
				bool flag7 = false;
				bool flag8 = false;
				bool flag9 = false;
				bool flag10 = false;
				bool flag11 = false;
				bool flag12 = false;
				bool flag13 = false;
				bool flag14 = false;
				if (periodInfo.menuID.Equals("O6") || periodInfo.menuID.Equals("O5") || periodInfo.menuID.Equals("O4"))
				{
					num3 = 9;
				}
				else if (periodInfo.menuID.Equals("O2"))
				{
					num3 = 4;
				}
				else if (periodInfo.menuID.Equals("O3"))
				{
					string text = "";
					string text2 = HTKKConfig.DATAFILES + "\\" + headerInfo.mst;
					string text3 = "01_CNV_TT57";
					string text4 = "";
					if (!string.IsNullOrEmpty(periodInfo.maCQTQuanLy))
					{
						text4 = "CQTQL" + periodInfo.maCQTQuanLy;
					}
					string versionXML = periodInfo.VersionXML;
					text = text2 + "\\" + text3 + ((!string.IsNullOrEmpty(text4)) ? ("_" + text4) : "") + ((!string.IsNullOrEmpty(periodInfo.MaHSDuAnDT.Trim())) ? ("_" + periodInfo.MaHSDuAnDT.Trim()) : "") + "_" + periodInfo.Ngay + periodInfo.Thang + periodInfo.Nam + "_" + versionXML.Replace(".", "") + ".xml";
					flag11 = ConvertHelpers.ConvertToBool(GetValueInXml(text, "L_5"));
					flag12 = ConvertHelpers.ConvertToBool(GetValueInXml(text, "L_9"));
					flag13 = ConvertHelpers.ConvertToBool(GetValueInXml(text, "X_9"));
					flag14 = ConvertHelpers.ConvertToBool(GetValueInXml(text, "X_5"));
					if (flag13 || flag12)
					{
						num3 = 9;
					}
					else if (flag11)
					{
						num3 = ((!flag14) ? 9 : 14);
					}
				}
				else if (periodInfo.menuID.Equals("O1"))
				{
					string text = "";
					string text5 = "";
					string text2 = HTKKConfig.DATAFILES + "\\" + headerInfo.mst;
					string text3 = "01_CPH_TT57";
					string text4 = "";
					if (!string.IsNullOrEmpty(periodInfo.maCQTQuanLy))
					{
						text4 = "CQTQL" + periodInfo.maCQTQuanLy;
					}
					string versionXML = periodInfo.VersionXML;
					text = text2 + "\\" + text3 + ((!string.IsNullOrEmpty(text4)) ? ("_" + text4) : "") + ((!string.IsNullOrEmpty(periodInfo.MaHSDuAnDT.Trim())) ? ("_" + periodInfo.MaHSDuAnDT.Trim()) : "") + "_" + periodInfo.Ngay + periodInfo.Thang + periodInfo.Nam + "_" + versionXML.Replace(".", "") + ".xml";
					text5 = text2 + "\\bs" + periodInfo.SoLan + "_" + text3 + ((!string.IsNullOrEmpty(text4)) ? ("_" + text4) : "") + ((!string.IsNullOrEmpty(periodInfo.MaHSDuAnDT.Trim())) ? ("_" + periodInfo.MaHSDuAnDT.Trim()) : "") + "_" + periodInfo.Ngay + periodInfo.Thang + periodInfo.Nam + "_" + versionXML.Replace(".", "") + ".xml";
					string text6 = "48~49~50~51~52~53~55~56~57";
					string[] array2 = text6.Split(new char[1] { '~' });
					for (int i = 0; i < array2.Length; i++)
					{
						double num5 = ConvertHelpers.ConvertToDouble(GetValueInXml(text, "U_" + array2[i]));
						if (num5 > 0.0)
						{
							switch (i)
							{
							case 0:
								flag2 = true;
								break;
							case 1:
								flag3 = true;
								break;
							case 2:
								flag4 = true;
								break;
							case 3:
								flag5 = true;
								break;
							case 4:
								flag6 = true;
								break;
							case 5:
								flag7 = true;
								break;
							case 6:
								flag8 = true;
								break;
							case 7:
								flag9 = true;
								break;
							case 8:
								flag10 = true;
								break;
							}
							break;
						}
					}
					if (!flag2 && !flag3 && !flag4 && !flag5 && !flag6 && !flag7 && !flag8 && !flag9 && !flag10)
					{
						for (int i = 0; i < array2.Length; i++)
						{
							double num5 = ConvertHelpers.ConvertToDouble(GetValueInXml(text5, "U_" + array2[i]));
							if (num5 > 0.0)
							{
								switch (i)
								{
								case 0:
									flag2 = true;
									break;
								case 1:
									flag3 = true;
									break;
								case 2:
									flag4 = true;
									break;
								case 3:
									flag5 = true;
									break;
								case 4:
									flag6 = true;
									break;
								case 5:
									flag7 = true;
									break;
								case 6:
									flag8 = true;
									break;
								case 7:
									flag9 = true;
									break;
								case 8:
									flag10 = true;
									break;
								}
								break;
							}
						}
					}
					if (flag2 || flag3 || flag5 || flag6)
					{
						num3 = 4;
					}
					else if (flag4 || flag9 || flag10)
					{
						num3 = 9;
					}
					else if (flag7)
					{
						num3 = 89;
					}
					else if (flag8)
					{
						num3 = 29;
					}
				}
				dateTime = new DateTime(num4, month, day).AddDays(num3);
				if (periodInfo.menuID.Equals("O2") || periodInfo.menuID.Equals("O5") || periodInfo.menuID.Equals("O4"))
				{
					int num6 = CountWeekendDays(new DateTime(num4, month, day), dateTime, DayOfWeek.Saturday);
					int num7 = CountWeekendDays(new DateTime(num4, month, day), dateTime, DayOfWeek.Sunday);
					dateTime = dateTime.AddDays(num6 + num7);
				}
				if (periodInfo.menuID.Equals("O1") && (flag2 || flag3 || flag5 || flag6 || flag9 || flag10))
				{
					int num6 = CountWeekendDays(new DateTime(num4, month, day), dateTime, DayOfWeek.Saturday);
					int num7 = CountWeekendDays(new DateTime(num4, month, day), dateTime, DayOfWeek.Sunday);
					dateTime = dateTime.AddDays(num6 + num7);
				}
				if (periodInfo.menuID.Equals("O3") && (flag13 || flag12 || (flag11 && !flag14)))
				{
					int num6 = CountWeekendDays(new DateTime(num4, month, day), dateTime, DayOfWeek.Saturday);
					int num7 = CountWeekendDays(new DateTime(num4, month, day), dateTime, DayOfWeek.Sunday);
					dateTime = dateTime.AddDays(num6 + num7);
				}
				break;
			}
			case "D_D":
			{
				string[] array = periodInfo.DenNgay.Split(new char[1] { '/' });
				dateTime = new DateTime(ConvertHelpers.ConvertToInt(array[2]), ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[0])).AddDays(45.0);
				break;
			}
			}
			if (periodInfo.menuID.Equals("47") || periodInfo.menuID == "K7")
			{
				int month = int.Parse(periodInfo.Thang);
				int day = int.Parse(periodInfo.Ngay);
				DateTime dateTime6 = new DateTime(2020, 12, 5);
				if (DateTime.ParseExact(periodInfo.NgayChiaCTLN, "dd/MM/yyyy", CultureInfo.CurrentCulture) < dateTime6)
				{
					dateTime = DateTime.ParseExact(periodInfo.NgayChiaCTLN, "dd/MM/yyyy", CultureInfo.CurrentCulture).AddDays(num2);
				}
			}
			if (periodInfo.menuID.Equals("21"))
			{
				int month = int.Parse(periodInfo.Thang);
				int day = int.Parse(periodInfo.Ngay);
				dateTime = new DateTime(num4, month, day).AddDays(periodInfo.SoNgayNopChamKHBS);
			}
			string text9 = "DMNgayLe.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text9);
			XmlNode val2 = ((XmlNode)val).SelectSingleNode("Root");
			foreach (XmlNode item in val2)
			{
				XmlNode val3 = item;
				string value = ((XmlNode)val3.Attributes["KieuKy"]).Value;
				string value2 = ((XmlNode)val3.Attributes["HanNop"]).Value;
				string[] array6 = value2.Split(new char[1] { '/' });
				string value3 = ((XmlNode)val3.Attributes["GiaTri"]).Value;
				string[] array7 = value3.Split(new char[1] { '/' });
				DateTime dateTime7 = new DateTime(ConvertHelpers.ConvertToInt(array6[2]), ConvertHelpers.ConvertToInt(array6[1]), ConvertHelpers.ConvertToInt(array6[0]));
				DateTime dateTime8 = new DateTime(ConvertHelpers.ConvertToInt(array7[2]), ConvertHelpers.ConvertToInt(array7[1]), ConvertHelpers.ConvertToInt(array7[0]));
				if (periodInfo.KieuKy.Equals("D") && value.Equals("D"))
				{
					if (dateTime == dateTime7)
					{
						dateTime = dateTime8;
						break;
					}
				}
				else if (periodInfo.KieuKy.Equals("Q") && value.Equals("Q") && dateTime == dateTime7)
				{
					dateTime = dateTime8;
					break;
				}
			}
			if (dateTime.DayOfWeek == DayOfWeek.Saturday)
			{
				dateTime = dateTime.AddDays(2.0);
			}
			else if (dateTime.DayOfWeek == DayOfWeek.Sunday)
			{
				dateTime = dateTime.AddDays(1.0);
			}
			if (dateTime.ToString("dd/MM/yyyy") == "20/02/2018")
			{
				dateTime = dateTime.AddDays(1.0);
			}
			return dateTime;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static DateTime MoveWeekendToMonday(DateTime sourceDate)
	{
		try
		{
			if (sourceDate.DayOfWeek == DayOfWeek.Saturday)
			{
				sourceDate = sourceDate.AddDays(2.0);
			}
			else if (sourceDate.DayOfWeek == DayOfWeek.Sunday)
			{
				sourceDate = sourceDate.AddDays(1.0);
			}
			return sourceDate;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static DateTime GetNgayBatDauDuocNopKHBS(HTKKPeriodInfo periodInfo)
	{
		try
		{
			DateTime sourceDate = GetHanNopToKhai(periodInfo).AddDays(1.0);
			return MoveWeekendToMonday(sourceDate);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string GetDataFileNameExtend(string strDataFileNameSaveExtend, string strVersionXML, string strDataFile, string mst)
	{
		string empty = string.Empty;
		try
		{
			if (strDataFile == "HEADER")
			{
				return HTKKConfig.DATAFILES + "\\" + mst + "\\" + strDataFile + "_.xml";
			}
			return HTKKConfig.DATAFILES + "\\" + mst + "\\" + HTKKConfig.VERSION + "_" + strVersionXML + "_" + strDataFile + "_" + strDataFileNameSaveExtend + ".xml";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static long getMaxNumber(string strNumber)
	{
		long num = 0L;
		try
		{
			if (!HTKKFormater.isNullString(strNumber))
			{
				return long.Parse(strNumber);
			}
			return 999999999999999L;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return 999999999999999L;
		}
	}

	public static long getMinNumber(string strNumber)
	{
		long num = 0L;
		try
		{
			if (!HTKKFormater.isNullString(strNumber))
			{
				return long.Parse(strNumber);
			}
			return -999999999999999L;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -999999999999999L;
		}
	}

	public static int getMaxLen(string strNumber)
	{
		int num = 0;
		try
		{
			if (!HTKKFormater.isNullString(strNumber))
			{
				return int.Parse(strNumber);
			}
			return 150;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return 150;
		}
	}

	public static void axfpsPhuLucSetAllValues(FpSpread axfpsPhuLuc, CheckBox chkChonPhuLucKeKhaiAll, bool value)
	{
		try
		{
			for (int i = 0; i < axfpsPhuLuc.Sheets[0].RowCount; i++)
			{
				if (!axfpsPhuLuc.Sheets[0].Cells[i, 0].Locked)
				{
					axfpsPhuLuc.Sheets[0].Cells[i, 0].Value = value;
					axfpsPhuLuc.Sheets[0].Cells[i, 0].Locked = value;
				}
			}
			chkChonPhuLucKeKhaiAll.Checked = value;
			((Control)chkChonPhuLucKeKhaiAll).Enabled = !value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static CheckBox getCheckBoxByName(List<CheckBox> listCheckBox, string strName)
	{
		for (int i = 0; i <= listCheckBox.Count; i++)
		{
			if (((Control)listCheckBox[i]).Name == strName)
			{
				return listCheckBox[i];
			}
		}
		return null;
	}

	public static MaskedTextBox getMaskedTextBoxByName(List<MaskedTextBox> listMaskedTextBox, string strName)
	{
		for (int i = 0; i <= listMaskedTextBox.Count; i++)
		{
			if (((Control)listMaskedTextBox[i]).Name == strName)
			{
				return listMaskedTextBox[i];
			}
		}
		return null;
	}

	public static ComboBox getComboBoxByName(List<ComboBox> listComboBox, string strName)
	{
		try
		{
			for (int i = 0; i <= listComboBox.Count; i++)
			{
				if (((Control)listComboBox[i]).Name == strName)
				{
					return listComboBox[i];
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return null;
	}

	public static RadioButton getRadioButtonByName(List<RadioButton> listRadioButton, string strName)
	{
		try
		{
			for (int i = 0; i <= listRadioButton.Count; i++)
			{
				if (((Control)listRadioButton[i]).Name == strName)
				{
					return listRadioButton[i];
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return null;
	}

	public static List<HTKKCMCRowValue> GetDataFromExcelFile(string filePath, HTKKCMCRowInfo rowInfo, int DataLocationTaiBangKeRowStart, string strEnd, int iColCheck)
	{
		List<HTKKCMCRowValue> list;
		try
		{
			FpSpread fpSpread = new FpSpread();
			((ISupportInitialize)fpSpread).BeginInit();
			fpSpread.Name = "winFpSpread";
			((ISupportInitialize)fpSpread).EndInit();
			fpSpread.ActiveSheetIndex = 0;
			fpSpread.OpenExcel(filePath);
			list = new List<HTKKCMCRowValue>();
			for (int i = DataLocationTaiBangKeRowStart - 1; i <= fpSpread.ActiveSheet.NonEmptyRowCount - 1 && fpSpread.ActiveSheet.Cells[i, iColCheck - 1].Text != strEnd; i++)
			{
				HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
				for (int j = 0; j < rowInfo.Cells.Count; j++)
				{
					if (rowInfo.Cells[j].CellBangKe != null && rowInfo.Cells[j].CellBangKe.ColIndex >= 1)
					{
						if (fpSpread.ActiveSheet.Cells[i, rowInfo.Cells[j].CellBangKe.ColIndex - 1].Value != null)
						{
							hTKKCMCRowValue.RowValue.Add(fpSpread.ActiveSheet.Cells[i, rowInfo.Cells[j].CellBangKe.ColIndex - 1].Value.ToString());
						}
						else
						{
							hTKKCMCRowValue.RowValue.Add("");
						}
					}
				}
				hTKKCMCRowValue.RowDisplay = rowInfo.RowDefaultDisplay;
				list.Add(hTKKCMCRowValue);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
		return list;
	}

	public static void InsertHyperLink(SheetView sView, int iRow, int iCol, string strCaption, string link)
	{
		try
		{
			HyperLinkCellType hyperLinkCellType = new HyperLinkCellType();
			hyperLinkCellType.Text = strCaption;
			hyperLinkCellType.Link = link;
			hyperLinkCellType.StringTrim = (StringTrimming)0;
			sView.Cells[iRow - 1, iCol - 1].Locked = false;
			sView.Cells[iRow - 1, iCol - 1].CellType = hyperLinkCellType;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void setHindenRow(SheetView sView, int beginR, int endR)
	{
		try
		{
			for (int i = beginR - 1; i < endR; i++)
			{
				Row row = sView.Rows[i];
				row.Visible = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ChangeValueCell(SheetView sView, int rowIndex, int columnIndex, bool locked, string formula, object value)
	{
		try
		{
			if (!string.IsNullOrEmpty(formula))
			{
				sView.Cells[rowIndex, columnIndex].Formula = formula;
			}
			else
			{
				sView.Cells[rowIndex, columnIndex].ResetFormula();
				sView.Cells[rowIndex, columnIndex].Value = value;
			}
			sView.Cells[rowIndex, columnIndex].Locked = locked;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static DateTime NgayDauKyTinhThue(HTKKPeriodInfo periodInfo)
	{
		try
		{
			DateTime result = default(DateTime);
			switch (periodInfo.KieuKy)
			{
			case "M":
				result = new DateTime(int.Parse(periodInfo.Nam), int.Parse(periodInfo.Thang), 1);
				break;
			case "Q":
			{
				int month = 3 * (int.Parse(periodInfo.Quy) - 1) + 1;
				result = new DateTime(int.Parse(periodInfo.Nam), month, 1);
				break;
			}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static DateTime GetNgayDauQuy(int iQuy, int iNam, int iNgayBdNamTc, int iThangBdNamTc)
	{
		DateTime result = default(DateTime);
		try
		{
			int num = (iQuy - 1) * 3 + iThangBdNamTc;
			int year = iNam;
			if (num > 12)
			{
				num -= 12;
				year = iNam + 1;
			}
			result = new DateTime(year, num, iNgayBdNamTc);
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
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

	public static DateTime NgayCuoiNamTaiChinh(int iNam, int iThangTc, int iNgayTc)
	{
		DateTime dateTime = new DateTime(iNam, iThangTc, iNgayTc);
		try
		{
			return dateTime.AddMonths(12).AddDays(-1.0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static DateTime NgayCuoiNamTaiChinh(HTKKHeaderInfo headerInfo, string nam)
	{
		string[] array = headerInfo.NgayBatDauNamTaiChinh.Split(new char[1] { '/' });
		return NgayCuoiNamTaiChinh(Convert.ToInt32(nam), ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[0]));
	}

	public static string FormatSoHD(string str)
	{
		try
		{
			if (str.IsNulOrEmpty())
			{
				return string.Empty;
			}
			string text;
			str = (text = ConvertHelpers.ConvertToInt(str).ToString());
			if (!text.Equals("") || (str.Length > 0 && str.Length < 7))
			{
				for (int i = 1; i <= 7 - str.Length; i++)
				{
					text = "0" + text;
				}
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string FormatKyHieu(string str)
	{
		try
		{
			string text = "";
			string text2 = "";
			text = str.ToUpper();
			if (str.Length == 5)
			{
				return text.Substring(0, 2) + "/" + text.Substring(2);
			}
			if (str.Length == 7)
			{
				if (HTKKFormater.IsNumber(str.Substring(0, 2)))
				{
					return text.Substring(0, 4) + "/" + text.Substring(4);
				}
				return text.Substring(0, 2) + "/" + text.Substring(2);
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string CheckSoHD(string str, string strLoai)
	{
		try
		{
			string text = "";
			string text2 = "ETP";
			string text3 = "TP";
			string text4 = "ABCDEGHKLMNPQRSTUVXY";
			text = "0";
			if (!strLoai.Equals("3") && str.Trim().Length != 6 && str.Trim().Length != 8)
			{
				return "2";
			}
			if (!strLoai.Equals("3"))
			{
				string text5;
				if (str.Trim().Length == 8)
				{
					text5 = str.Substring(0, 2);
					str = str.Substring(2, 6);
				}
				else
				{
					text5 = "01";
				}
				text = ((text4.IndexOf(str.Substring(0, 1).ToUpper()) >= 0 && text4.IndexOf(str.Substring(1, 1).ToUpper()) >= 0 && HTKKFormater.IsNumber(text5) && int.Parse(text5) > 0 && int.Parse(text5) < 65) ? ((str.IndexOf("/") <= -1 || !HTKKFormater.IsNumber(str.Substring(3, 2))) ? "1" : ((ConvertHelpers.ConvertToInt(str.Substring(3, 2)) < 0) ? "1" : ((text2.IndexOf(str.Substring(str.Length - 1, 1)) < 0) ? "1" : "0"))) : ((strMauSoHD_01GTKT == null || !strMauSoHD_01GTKT.Trim().StartsWith("01GTKT")) ? "1" : ((str.IndexOf("/") <= -1 || !"NN".Equals(text5) || !"XC".Equals(str.Substring(1, 2))) ? "1" : ((text3.IndexOf(str.Substring(str.Length - 1, 1).ToUpper()) <= -1) ? "1" : "0"))));
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string CheckMauSoHD(string str, string strLoai, string strTemp)
	{
		try
		{
			string text = "";
			string result = "0";
			str = str.Trim();
			if ("0".Equals(strLoai))
			{
				if (str.Length != 11 && str.Length != 13)
				{
					result = "1";
				}
				else if (!strTemp.Equals(str.Substring(0, 6)))
				{
					result = "1";
				}
				else
				{
					string text2 = str.Substring(6, 1);
					string value = str.Substring(7, 1);
					string text3 = str.Substring(8, 3);
					if (str.Length == 13)
					{
						text = str.Substring(11, 2);
					}
					if (HTKKFormater.IsNumber(text2))
					{
						if ("01BHDT".Equals(strTemp))
						{
							if (int.Parse(text2) < 0 || int.Parse(text2) > 9)
							{
								result = "2";
							}
						}
						else if (int.Parse(text2) < 2 || int.Parse(text2) > 9)
						{
							result = "2";
						}
					}
					else
					{
						result = "2";
					}
					if (!"/".Equals(value))
					{
						result = "1";
					}
					if (!HTKKFormater.IsNumber(text3) || int.Parse(text3) < 0)
					{
						result = "1";
					}
					if (!"".Equals(text.Trim()) && !"BD".Equals(text.Trim()) && !"IV".Equals(text.Trim()))
					{
						result = "1";
					}
				}
			}
			else if ("3".Equals(strLoai))
			{
				result = "0";
			}
			else if ("4".Equals(strLoai))
			{
				if (!strTemp.Equals(str.Substring(0, 6)))
				{
					result = "1";
				}
			}
			else
			{
				if (!strTemp.Equals(str.Substring(0, 3)))
				{
					result = "1";
				}
				if (str.Length > 20 || str.Length < 3)
				{
					result = "1";
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string CheckSoLienHDDT(string str, string strLoai, string strTemp, string strKyHieuHD)
	{
		try
		{
			string result = "0";
			str = str.Trim();
			if ("0".Equals(strLoai) && (str.Length == 11 || str.Length == 13) && strTemp.Trim().Equals(str.Substring(0, 6)))
			{
				string text = str.Substring(6, 1);
				if (!HTKKFormater.IsNumber(text))
				{
					return "3";
				}
				if (!"E".Equals(strKyHieuHD.Trim().Substring(strKyHieuHD.Trim().Length - 1, 1)))
				{
					return "0";
				}
				if (int.Parse(text) != 0)
				{
					return "3";
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string CheckSoLienBLPDT(string str, string strTemp, string strKyHieuBLP)
	{
		try
		{
			string result = "0";
			str = str.Trim();
			if (strTemp.Trim().Equals(str.Substring(0, 5)))
			{
				string text = str.Substring(5, 1);
				if (!HTKKFormater.IsNumber(text))
				{
					return "3";
				}
				if ("E".Equals(strKyHieuBLP.Trim().Substring(strKyHieuBLP.Trim().Length - 1, 1)))
				{
					if (int.Parse(text) != 0)
					{
						return "3";
					}
				}
				else if ("T,P".Contains(strKyHieuBLP.Trim().Substring(strKyHieuBLP.Trim().Length - 1, 1)) && int.Parse(text) == 0)
				{
					return "3";
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string FormatKyHieuBLP(string str)
	{
		try
		{
			string text = "";
			string text2 = str.ToUpper();
			if (str.Length == 5)
			{
				return text2.Substring(0, 2) + "-" + text2.Substring(2, 3);
			}
			if (str.Length == 7)
			{
				if (HTKKFormater.IsNumber(text2.Substring(0, 2)))
				{
					return text2.Substring(0, 4) + "-" + text2.Substring(4, 3);
				}
				return text2.Substring(0, 2) + "-" + text2.Substring(2, 5);
			}
			return text2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string CheckSoBLP(string str, string strMaLoaiBLP, bool isAllowE)
	{
		try
		{
			string text = ((!isAllowE) ? "TP" : "TPE");
			string text2 = "ABCDEGHKLMNPQRSTUVXY";
			string text3 = "0";
			str = str.Trim();
			if ("01BLP".Equals(strMaLoaiBLP))
			{
				if (str.Length != 6 && str.Length != 8)
				{
					return "2";
				}
			}
			else if (str.Length != 6)
			{
				return "2";
			}
			string text4;
			if (str.Length == 8)
			{
				text4 = str.Substring(0, 2);
				str = str.Substring(2, 6);
			}
			else
			{
				text4 = "01";
			}
			return "01BLP".Equals(strMaLoaiBLP) ? ((text2.IndexOf(str.Substring(0, 1).ToUpper()) < 0 || text2.IndexOf(str.Substring(1, 1).ToUpper()) < 0 || !HTKKFormater.IsNumber(text4) || int.Parse(text4) <= 0 || int.Parse(text4) >= 65) ? "1" : ((str.IndexOf("-") <= 0 || !HTKKFormater.IsNumber(str.Substring(3, 2))) ? "1" : ((ConvertHelpers.ConvertToInt(str.Substring(3, 2)) < 0) ? "1" : ((text.IndexOf(str.Substring(str.Length - 1, 1)) < 0) ? "1" : "0")))) : ((text2.IndexOf(str.Substring(0, 1).ToUpper()) < 0 || text2.IndexOf(str.Substring(1, 1).ToUpper()) < 0) ? "1" : ((str.IndexOf("-") < 0 || !HTKKFormater.IsNumber(str.Substring(3, 2))) ? "1" : ((ConvertHelpers.ConvertToInt(str.Substring(3, 2)) < 0) ? "1" : ((text.IndexOf(str.Substring(str.Length - 1), 0) < 0) ? "1" : "0"))));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string CheckMauSoBLP(string str, string strLoai, string strTemp)
	{
		try
		{
			string result = "0";
			str = str.Trim();
			if ("0".Equals(strLoai))
			{
				if (str.Length != 10)
				{
					return "1";
				}
				if (!strTemp.Trim().Equals(str.Substring(0, 5)))
				{
					return "1";
				}
				string text = str.Substring(5, 1);
				string value = str.Substring(6, 1);
				string text2 = str.Substring(7, 3);
				if (HTKKFormater.IsNumber(text))
				{
					if (int.Parse(text) < 1 || int.Parse(text) > 9)
					{
						result = "2";
					}
				}
				else
				{
					result = "2";
				}
				if (!"-".Equals(value))
				{
					return "1";
				}
				if (!HTKKFormater.IsNumber(text2) || int.Parse(text2) < 0)
				{
					return "1";
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string CheckSoTemRuou(string str)
	{
		try
		{
			string text = "P";
			string text2 = "ABCDEGHKLMNPQRSTUVXY";
			string text3 = "0";
			str = str.Trim();
			if (str.Length != 6)
			{
				return "2";
			}
			return (text2.IndexOf(str.Substring(0, 1).ToUpper()) <= -1 || text2.IndexOf(str.Substring(1, 1).ToUpper()) <= -1) ? "1" : ((str.IndexOf("/") <= 0 || !HTKKFormater.IsNumber(str.Substring(3, 2))) ? "1" : ((text.IndexOf(str.Substring(str.Length - 1, 1)) <= -1) ? "1" : "0"));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string CheckSoChungTu(string str)
	{
		try
		{
			string text = "T";
			string text2 = "ABCDEGHKLMNPQRSTUVXY";
			string text3 = "0";
			str = str.Trim();
			if (str.Length != 9)
			{
				return "2";
			}
			return (text2.IndexOf(str.Substring(0, 1).ToUpper()) <= -1 || text2.IndexOf(str.Substring(1, 1).ToUpper()) <= -1) ? "1" : ((str.IndexOf("/") <= 0 || !HTKKFormater.IsNumber(str.Substring(3, 4))) ? "1" : ((text.IndexOf(str.Substring(str.Length - 1, 1)) <= -1 || !(str.Substring(str.Length - 2, 1) == "/")) ? "1" : "0"));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static bool isDuLieuHL(string str)
	{
		try
		{
			bool result = true;
			bool flag = false;
			bool flag2 = false;
			if (HTKKFormater.isNullString(str))
			{
				return true;
			}
			if (str.IndexOf(";") >= 0)
			{
				flag = true;
			}
			if (str.IndexOf("-") >= 0)
			{
				flag2 = true;
				string[] array = str.Split(new char[1] { '-' });
				for (int i = 1; i < array.Length - 1; i++)
				{
					if (array[i].IndexOf(";") < 0)
					{
						return false;
					}
				}
			}
			if (flag && !flag2)
			{
				string[] array2 = str.Split(new char[1] { ';' });
				for (int i = 0; i < array2.Length; i++)
				{
					if (HTKKFormater.IsNullValueAc(array2[i]))
					{
						result = false;
						break;
					}
					if (array2[i].IndexOf(" ") >= 0)
					{
						result = false;
						break;
					}
					if (!HTKKFormater.isNullString(array2[i]))
					{
						if (!HTKKFormater.IsNumber(array2[i]))
						{
							result = false;
							break;
						}
						if (array2[i].IndexOf(",") >= 0 || array2[i].IndexOf(".") >= 0)
						{
							result = false;
							break;
						}
					}
				}
			}
			else if (!flag && flag2)
			{
				string[] array3 = str.Split(new char[1] { '-' });
				for (int i = 0; i < array3.Length; i++)
				{
					if (HTKKFormater.IsNullValueAc(array3[i]))
					{
						result = false;
						break;
					}
					if (array3[i].IndexOf(" ") >= 0)
					{
						result = false;
						break;
					}
					if (!HTKKFormater.IsNullValueAc(array3[i]))
					{
						if (!HTKKFormater.IsNumber(array3[i]))
						{
							result = false;
							break;
						}
						if (array3[i].IndexOf(",") >= 0 || array3[i].IndexOf(".") >= 0)
						{
							result = false;
							break;
						}
					}
				}
			}
			else if (flag && flag2)
			{
				string[] array2 = str.Split(new char[1] { ';' });
				for (int i = 0; i < array2.Length; i++)
				{
					if (HTKKFormater.IsNullValueAc(array2[i]))
					{
						result = false;
						break;
					}
					if (array2[i].IndexOf("-") < 0)
					{
						continue;
					}
					string[] array3 = array2[i].Split(new char[1] { '-' });
					for (int j = 0; j < array3.Length; j++)
					{
						if (HTKKFormater.IsNullValueAc(array3[j]))
						{
							result = false;
							break;
						}
						if (array3[j].IndexOf(" ") >= 0)
						{
							result = false;
							break;
						}
						if (!HTKKFormater.isNullString(array3[j]))
						{
							if (!HTKKFormater.IsNumber(array3[j]))
							{
								result = false;
								break;
							}
							if (array3[j].IndexOf(",") >= 0 || array3[j].IndexOf(".") >= 0)
							{
								result = false;
								break;
							}
						}
					}
				}
			}
			else
			{
				if ("".Equals(str))
				{
					result = false;
				}
				if (!HTKKFormater.IsNumber(str))
				{
					result = false;
				}
				else if (str.IndexOf(",") >= 0 || str.IndexOf(".") >= 0)
				{
					result = false;
				}
				if (str.IndexOf(" ") >= 0)
				{
					result = false;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string GetSoLuongHD(string str)
	{
		try
		{
			double num;
			if (isDuLieuHL(str) && !"".Equals(str))
			{
				bool flag = false;
				bool flag2 = false;
				if (str.IndexOf(";") > 0)
				{
					flag = true;
				}
				if (str.IndexOf("-") > 0)
				{
					flag2 = true;
				}
				if (flag && !flag2)
				{
					string[] array = str.Split(new char[1] { ';' });
					num = array.Length;
				}
				else if (!flag && flag2)
				{
					string[] array2 = str.Split(new char[1] { '-' });
					num = ConvertHelpers.ConvertToDouble(array2[1]) - ConvertHelpers.ConvertToDouble(array2[0]) + 1.0;
				}
				else if (flag && flag2)
				{
					string[] array3 = str.Split(new char[1] { ';' });
					double num2 = 0.0;
					for (int i = 0; i < array3.Length; i++)
					{
						if (array3[i].IndexOf("-") > 0)
						{
							string[] array2 = array3[i].Split(new char[1] { '-' });
							num2 = num2 + ConvertHelpers.ConvertToDouble(array2[1]) - ConvertHelpers.ConvertToDouble(array2[0]) + 1.0;
						}
						else
						{
							num2 += 1.0;
						}
					}
					num = num2;
				}
				else
				{
					num = 1.0;
				}
			}
			else
			{
				num = 0.0;
			}
			return num.ToString();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string[] getTusoDenso(string Tuso1, string denso1, string Tuso2, string denso2)
	{
		try
		{
			string[] array = new string[6];
			bool flag = false;
			bool flag2 = false;
			bool flag3 = false;
			bool flag4 = false;
			if ("".Equals(Tuso1))
			{
				flag = true;
			}
			if ("".Equals(denso1))
			{
				flag2 = true;
			}
			if ("".Equals(Tuso2))
			{
				flag3 = true;
			}
			if ("".Equals(denso2))
			{
				flag4 = true;
			}
			double num = ConvertHelpers.ConvertToDouble(Tuso1);
			double num2 = ConvertHelpers.ConvertToDouble(denso1);
			double num3 = ConvertHelpers.ConvertToDouble(Tuso2);
			double num4 = ConvertHelpers.ConvertToDouble(denso2);
			if (num2 - num < 0.0 || num4 - num3 < 0.0)
			{
				array[0] = "-1";
				return array;
			}
			if ((num >= 0.0 || num2 >= 0.0) && flag3 && flag4)
			{
				if (num >= 0.0 && !flag && num2 == 0.0 && flag2)
				{
					array[0] = num.ToString();
					array[1] = "";
					array[2] = "";
					array[4] = "2";
				}
				else if (num == 0.0 && flag && num2 >= 0.0 && !flag2)
				{
					array[0] = "";
					array[1] = num2.ToString();
					array[2] = "";
					array[4] = "1";
				}
				else if (num2 - num >= 0.0 && !flag && !flag2)
				{
					array[0] = num.ToString();
					array[1] = num2.ToString();
					array[2] = (num2 - num + 1.0).ToString();
					array[4] = "";
				}
				else
				{
					array[0] = num.ToString();
					array[1] = num2.ToString();
					array[2] = "0";
					array[4] = "3";
				}
				array[3] = "1";
			}
			else if ((num3 >= 0.0 || num4 >= 0.0) && flag && flag2)
			{
				if (num3 == 0.0 && flag3 && num4 >= 0.0 && !flag4)
				{
					array[0] = "";
					array[1] = num4.ToString();
					array[2] = "0";
					array[4] = "4";
				}
				else if (num3 >= 0.0 && !flag3 && num4 == 0.0 && flag4)
				{
					array[0] = num3.ToString();
					array[1] = "";
					array[2] = "0";
					array[4] = "5";
				}
				else if (num4 - num3 >= 0.0)
				{
					array[0] = num3.ToString();
					array[1] = num4.ToString();
					array[2] = (num4 - num3 + 1.0).ToString();
					array[4] = "";
				}
				else
				{
					array[0] = num3.ToString();
					array[1] = num4.ToString();
					array[2] = "0";
					array[4] = "6";
				}
				array[3] = "2";
			}
			else if ((num >= 0.0 || num2 >= 0.0) && (num3 >= 0.0 || num4 >= 0.0))
			{
				if (num >= 0.0 && !flag && "".Equals(denso1) && flag2)
				{
					array[0] = num.ToString();
					array[1] = "";
					array[2] = "0";
					array[4] = "1";
				}
				else if (num == 0.0 && flag && num2 >= 0.0 && !flag2)
				{
					array[0] = num.ToString();
					array[1] = "";
					array[2] = "0";
					array[4] = "2";
				}
				else if (num2 - num >= 0.0)
				{
					array[0] = num.ToString();
					array[1] = num2.ToString();
					array[2] = (num2 - num + 1.0).ToString();
					array[4] = "";
				}
				else
				{
					array[0] = num.ToString();
					array[1] = num2.ToString();
					array[2] = "0";
					array[4] = "3";
				}
				if (num >= 0.0 && !flag && "".Equals(denso1) && flag2)
				{
					array[0] = num.ToString();
					array[1] = "";
					array[2] = "0";
					array[4] = "1";
				}
				else if ("".Equals(Tuso1) && flag && num2 >= 0.0 && !flag2)
				{
					array[0] = num.ToString();
					array[1] = "";
					array[2] = "0";
					array[4] = "2";
				}
				else if (num2 - num >= 0.0)
				{
					array[0] = num.ToString();
					array[1] = num2.ToString();
					array[2] = (num2 - num + 1.0).ToString();
					array[4] = "";
				}
				else
				{
					array[0] = num.ToString();
					array[1] = num2.ToString();
					array[2] = "0";
					array[4] = "3";
				}
				if (num3 - num2 != 1.0)
				{
					array[0] = "-1";
				}
				else
				{
					array[0] = num.ToString();
					array[1] = num4.ToString();
					array[2] = (num4 - num + 1.0).ToString();
					array[3] = "3";
				}
			}
			else
			{
				array[0] = "-1";
			}
			return array;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string checkHDXoa(string Tuso1, string denso1, string strCheck)
	{
		try
		{
			string result = "0";
			if (HTKKFormater.isNullString(strCheck))
			{
				return result;
			}
			if (isDuLieuHL(strCheck))
			{
				if (strCheck.IndexOf(";") > 0)
				{
					string[] array = strCheck.Split(new char[1] { ';' });
					for (int i = 0; i < array.Length; i++)
					{
						if (array[i].IndexOf("-") > 0)
						{
							string[] array2 = array[i].Split(new char[1] { '-' });
							if (ConvertHelpers.ConvertToDouble(array2[0]) < ConvertHelpers.ConvertToDouble(Tuso1) || ConvertHelpers.ConvertToDouble(array2[1]) > ConvertHelpers.ConvertToDouble(denso1))
							{
								return "-1";
							}
						}
						else if (ConvertHelpers.ConvertToDouble(array[i]) < ConvertHelpers.ConvertToDouble(Tuso1) || ConvertHelpers.ConvertToDouble(array[i]) > ConvertHelpers.ConvertToDouble(denso1))
						{
							return "-1";
						}
					}
				}
				else if (strCheck.IndexOf("-") > 0)
				{
					string[] array2 = strCheck.Split(new char[1] { '-' });
					if (ConvertHelpers.ConvertToDouble(array2[0]) < ConvertHelpers.ConvertToDouble(Tuso1) || ConvertHelpers.ConvertToDouble(array2[1]) > ConvertHelpers.ConvertToDouble(denso1))
					{
						return "-1";
					}
				}
				else if (ConvertHelpers.ConvertToDouble(strCheck) < ConvertHelpers.ConvertToDouble(Tuso1) || ConvertHelpers.ConvertToDouble(strCheck) > ConvertHelpers.ConvertToDouble(denso1))
				{
					return "-1";
				}
			}
			if (isTrungSoHD(strCheck))
			{
				result = "-1";
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static bool isTrungSoHD(string str)
	{
		try
		{
			bool result = false;
			if (HTKKFormater.isNullString(str))
			{
				return false;
			}
			if (isDuLieuHL(str) && str.IndexOf(";") > 0)
			{
				string[] array = str.Split(new char[1] { ';' });
				for (int i = 0; i < array.Length; i++)
				{
					if (array[i].IndexOf("-") > 0)
					{
						string[] array2 = array[i].Split(new char[1] { '-' });
						double num = ConvertHelpers.ConvertToDouble(array2[0]);
						double num2 = ConvertHelpers.ConvertToDouble(array2[1]);
						for (int j = i + 1; j < array.Length; j++)
						{
							if (array[j].IndexOf("-") > 0)
							{
								string[] array3 = array[j].Split(new char[1] { '-' });
								double num3 = ConvertHelpers.ConvertToDouble(array3[0]);
								double num4 = ConvertHelpers.ConvertToDouble(array3[1]);
								if ((num - num3) * (num2 - num3) <= 0.0 || (num - num4) * (num2 - num4) <= 0.0 || (num > num3 && num2 < num4))
								{
									return true;
								}
							}
							else if (ConvertHelpers.ConvertToDouble(array[j]) >= num && ConvertHelpers.ConvertToDouble(array[j]) <= num2)
							{
								return true;
							}
						}
						continue;
					}
					for (int j = i + 1; j < array.Length; j++)
					{
						if (array[j].IndexOf("-") > 0)
						{
							string[] array3 = array[j].Split(new char[1] { '-' });
							double num3 = ConvertHelpers.ConvertToDouble(array3[0]);
							double num4 = ConvertHelpers.ConvertToDouble(array3[1]);
							if (num3 <= ConvertHelpers.ConvertToDouble(array[i]) && num4 >= ConvertHelpers.ConvertToDouble(array[i]))
							{
								return true;
							}
						}
						else if (ConvertHelpers.ConvertToDouble(array[j]) == ConvertHelpers.ConvertToDouble(array[i]))
						{
							return true;
						}
					}
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string checkTrungXoa(string strXoa, string strMat, string strHuy)
	{
		try
		{
			bool flag = false;
			bool flag2 = false;
			bool flag3 = false;
			string result = "0";
			if (!HTKKFormater.isNullString(strXoa) && !HTKKFormater.isNullString(strMat) && HTKKFormater.isNullString(strHuy))
			{
				if (isTrungSoHD(strXoa + ";" + strMat))
				{
					return "1";
				}
			}
			else if (!HTKKFormater.isNullString(strXoa) && HTKKFormater.isNullString(strMat) && !HTKKFormater.isNullString(strHuy))
			{
				if (isTrungSoHD(strXoa + ";" + strHuy))
				{
					return "2";
				}
			}
			else if (HTKKFormater.isNullString(strXoa) && !HTKKFormater.isNullString(strMat) && !HTKKFormater.isNullString(strHuy))
			{
				if (isTrungSoHD(strMat + ";" + strHuy))
				{
					return "3";
				}
			}
			else if (!HTKKFormater.isNullString(strXoa) && !HTKKFormater.isNullString(strMat) && !HTKKFormater.isNullString(strHuy))
			{
				if (isTrungSoHD(strXoa + ";" + strMat))
				{
					flag2 = true;
				}
				if (isTrungSoHD(strXoa + ";" + strHuy))
				{
					flag3 = true;
				}
				if (isTrungSoHD(strMat + ";" + strHuy))
				{
					flag = true;
				}
				result = ((flag && !flag2 && !flag3) ? "3" : ((flag2 && !flag && !flag3) ? "1" : ((flag3 && !flag && !flag2) ? "2" : (((!flag3 || !flag || !flag2) && (flag3 || !flag || !flag2) && (!flag3 || flag || !flag2) && (!flag3 || !flag || flag2)) ? "0" : "4"))));
			}
			else
			{
				result = "0";
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string[] checkSoTonSoPH(string Tuso1, string denso1, string Tuso2, string denso2)
	{
		try
		{
			string[] array = new string[5] { "", "", "", "", "" };
			if ((!HTKKFormater.isNullString(Tuso1) || !HTKKFormater.isNullString(denso1)) && HTKKFormater.isNullString(Tuso2) && HTKKFormater.isNullString(denso2))
			{
				array[0] = "1";
				if (!HTKKFormater.isNullString(Tuso1) && HTKKFormater.isNullString(denso1))
				{
					array[1] = "2";
				}
				else if (HTKKFormater.isNullString(Tuso1) && !HTKKFormater.isNullString(denso1))
				{
					array[1] = "1";
				}
				else if (ConvertHelpers.ConvertToDouble(denso1) - ConvertHelpers.ConvertToDouble(Tuso1) < 0.0)
				{
					array[1] = "3";
				}
				else
				{
					array[1] = "4";
				}
			}
			else if (HTKKFormater.isNullString(Tuso1) && HTKKFormater.isNullString(denso1) && (!HTKKFormater.isNullString(Tuso2) || !HTKKFormater.isNullString(denso2)))
			{
				array[0] = "2";
				if (!HTKKFormater.isNullString(Tuso2) && HTKKFormater.isNullString(denso2))
				{
					array[2] = "2";
				}
				else if (HTKKFormater.isNullString(Tuso2) && !HTKKFormater.isNullString(denso2))
				{
					array[2] = "1";
				}
				else if (ConvertHelpers.ConvertToDouble(denso2) - ConvertHelpers.ConvertToDouble(Tuso2) < 0.0)
				{
					array[2] = "3";
				}
				else
				{
					array[2] = "4";
				}
			}
			else if ((!HTKKFormater.isNullString(Tuso1) || !HTKKFormater.isNullString(denso1)) && (!HTKKFormater.isNullString(Tuso2) || !HTKKFormater.isNullString(denso2)))
			{
				array[0] = "3";
				if (!HTKKFormater.isNullString(Tuso1) && HTKKFormater.isNullString(denso1))
				{
					array[1] = "2";
				}
				else if (HTKKFormater.isNullString(Tuso1) && !HTKKFormater.isNullString(denso1))
				{
					array[1] = "1";
				}
				else if (ConvertHelpers.ConvertToDouble(denso1) - ConvertHelpers.ConvertToDouble(Tuso1) < 0.0)
				{
					array[1] = "3";
				}
				else
				{
					array[1] = "4";
				}
				if (!HTKKFormater.isNullString(Tuso2) && HTKKFormater.isNullString(denso2))
				{
					array[2] = "2";
				}
				else if (HTKKFormater.isNullString(Tuso2) && !HTKKFormater.isNullString(denso2))
				{
					array[2] = "1";
				}
				else if (ConvertHelpers.ConvertToDouble(denso2) - ConvertHelpers.ConvertToDouble(Tuso2) < 0.0)
				{
					array[2] = "3";
				}
				else
				{
					array[2] = "4";
				}
				if ("4".Equals(array[2]) && "4".Equals(array[1]) && ConvertHelpers.ConvertToDouble(Tuso2) - ConvertHelpers.ConvertToDouble(denso1) != 1.0 && !HTKKFormater.isNullString(Tuso1) && !HTKKFormater.isNullString(denso1) && !HTKKFormater.isNullString(Tuso2) && !HTKKFormater.isNullString(denso2))
				{
					array[3] = "1";
				}
			}
			else
			{
				array[0] = "0";
			}
			return array;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string[] checkTrung(FpSpread winFpSpread, int startRow, string charTuSo, string charDenSo, string charSTT, string charTenHD, string charMso, string charKH, string charMST, bool isMST, int tSheet)
	{
		try
		{
			string text = "";
			SheetView sheetView = winFpSpread.Sheets[tSheet - 1];
			int column = HTKKFile.ColumnLettersToIndex(charTuSo) - 1;
			int column2 = HTKKFile.ColumnLettersToIndex(charDenSo) - 1;
			int column3 = HTKKFile.ColumnLettersToIndex(charSTT) - 1;
			int column4 = HTKKFile.ColumnLettersToIndex(charTenHD) - 1;
			int column5 = HTKKFile.ColumnLettersToIndex(charMso) - 1;
			int column6 = HTKKFile.ColumnLettersToIndex(charKH) - 1;
			int column7 = -1;
			if (isMST)
			{
				column7 = HTKKFile.ColumnLettersToIndex(charMST) - 1;
			}
			int num = startRow;
			do
			{
				string text2 = ((sheetView.Cells[num, column].Value == null) ? "" : sheetView.Cells[num, column].Value.ToString());
				string text3 = ((sheetView.Cells[num, column2].Value == null) ? "" : sheetView.Cells[num, column2].Value.ToString());
				string text4 = "";
				if (isMST)
				{
					text4 = sheetView.Cells[num, column7].Text;
				}
				double num2 = ConvertHelpers.ConvertToDouble(text2);
				double num3 = ConvertHelpers.ConvertToDouble(text3);
				string text5 = sheetView.Cells[num, column4].Text;
				string text6 = sheetView.Cells[num, column5].Text;
				string text7 = sheetView.Cells[num, column6].Text;
				int num4 = num + 1;
				if ("AA".Equals(sheetView.Cells[num4, column3].Value.ToString().ToUpper()))
				{
					break;
				}
				do
				{
					string text8 = ((sheetView.Cells[num4, column].Value == null) ? "" : sheetView.Cells[num4, column].Value.ToString());
					string text9 = ((sheetView.Cells[num4, column2].Value == null) ? "" : sheetView.Cells[num4, column2].Value.ToString());
					double num5 = ConvertHelpers.ConvertToDouble(text8);
					double num6 = ConvertHelpers.ConvertToDouble(text9);
					string text10 = sheetView.Cells[num4, column4].Text;
					string text11 = sheetView.Cells[num4, column5].Text;
					string text12 = sheetView.Cells[num4, column6].Text;
					string value = "";
					if (isMST)
					{
						value = sheetView.Cells[num4, column7].Text;
					}
					bool flag = !isMST || (text4.Equals(value) ? true : false);
					if (((num2 - num5) * (num3 - num5) <= 0.0 || (num2 - num6) * (num3 - num6) <= 0.0 || (num2 > num5 && num3 < num6)) && text5.Equals(text10) && text6.Equals(text11) && text7.Equals(text12) && flag && !HTKKFormater.isNullString(text8) && !HTKKFormater.isNullString(text9) && !HTKKFormater.isNullString(text2) && !HTKKFormater.isNullString(text3))
					{
						if (text.IndexOf(num.ToString()) < 0)
						{
							text = text + "~" + num;
						}
						if (text.IndexOf(num4.ToString()) < 0)
						{
							text = text + "~" + num4;
						}
					}
					num4++;
				}
				while (!"AA".Equals(sheetView.Cells[num4 - 1, column3].Value.ToString().ToUpper()));
				num++;
			}
			while (!"AA".Equals(sheetView.Cells[num, column3].Value.ToString().ToUpper()));
			if (text.Length > 2)
			{
				text = text.Substring(1, text.Length - 1);
			}
			return text.Split(new char[1] { '~' });
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string[] checkTrung_01(FpSpread winFpSpread, int startRow, string charTuSo, string charDenSo, string charSTT, string charTenHD, string charMso, string charKH, string charMST, bool isMST, int tSheet, string charEnd)
	{
		try
		{
			string text = "";
			SheetView sheetView = winFpSpread.Sheets[tSheet - 1];
			int column = HTKKFile.ColumnLettersToIndex(charTuSo) - 1;
			int column2 = HTKKFile.ColumnLettersToIndex(charDenSo) - 1;
			int column3 = HTKKFile.ColumnLettersToIndex(charSTT) - 1;
			int column4 = HTKKFile.ColumnLettersToIndex(charTenHD) - 1;
			int column5 = HTKKFile.ColumnLettersToIndex(charMso) - 1;
			int column6 = HTKKFile.ColumnLettersToIndex(charKH) - 1;
			int column7 = -1;
			if (isMST)
			{
				column7 = HTKKFile.ColumnLettersToIndex(charMST);
			}
			int num = startRow;
			do
			{
				string text2 = ((sheetView.Cells[num, column].Value == null) ? "" : sheetView.Cells[num, column].Value.ToString());
				string text3 = ((sheetView.Cells[num, column2].Value == null) ? "" : sheetView.Cells[num, column2].Value.ToString());
				string text4 = "";
				if (isMST)
				{
					text4 = sheetView.Cells[num, column7].Text;
				}
				double num2 = ConvertHelpers.ConvertToDouble(text2);
				double num3 = ConvertHelpers.ConvertToDouble(text3);
				string text5 = sheetView.Cells[num, column4].Text;
				string text6 = sheetView.Cells[num, column5].Text;
				string text7 = sheetView.Cells[num, column6].Text;
				int num4 = num + 1;
				if (charEnd.ToUpper().Equals(sheetView.Cells[num4, column3].Value.ToString().ToUpper()))
				{
					break;
				}
				do
				{
					string text8 = ((sheetView.Cells[num4, column].Value == null) ? "" : sheetView.Cells[num4, column].Value.ToString());
					string text9 = ((sheetView.Cells[num4, column2].Value == null) ? "" : sheetView.Cells[num4, column2].Value.ToString());
					double num5 = ConvertHelpers.ConvertToDouble(text8);
					double num6 = ConvertHelpers.ConvertToDouble(text9);
					string text10 = sheetView.Cells[num4, column4].Text;
					string text11 = sheetView.Cells[num4, column5].Text;
					string text12 = sheetView.Cells[num4, column6].Text;
					string value = "";
					if (isMST)
					{
						value = sheetView.Cells[num4, column7].Text;
					}
					bool flag = !isMST || (text4.Equals(value) ? true : false);
					if (((num2 - num5) * (num3 - num5) <= 0.0 || (num2 - num6) * (num3 - num6) <= 0.0 || (num2 > num5 && num3 < num6)) && text5.Equals(text10) && text6.Equals(text11) && text7.Equals(text12) && flag && !HTKKFormater.isNullString(text8) && !HTKKFormater.isNullString(text9) && !HTKKFormater.isNullString(text2) && !HTKKFormater.isNullString(text3))
					{
						if (text.IndexOf(num.ToString()) < 0)
						{
							text = text + "~" + num;
						}
						if (text.IndexOf(num4.ToString()) < 0)
						{
							text = text + "~" + num4;
						}
					}
					num4++;
				}
				while (!charEnd.ToUpper().Equals(sheetView.Cells[num4 - 1, column3].Value.ToString().ToUpper()));
				num++;
			}
			while (!charEnd.ToUpper().Equals(sheetView.Cells[num, column3].Value.ToString().ToUpper()));
			if (text.Length > 2)
			{
				text = text.Substring(1, text.Length - 1);
			}
			return text.Split(new char[1] { '~' });
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void CompareKeycode(KeyPressEventArgs e, int styleCell)
	{
		try
		{
			int hashCode = e.KeyChar.GetHashCode();
			int hashCode2 = '0'.GetHashCode();
			int hashCode3 = '9'.GetHashCode();
			int hashCode4 = '%'.GetHashCode();
			int hashCode5 = '('.GetHashCode();
			switch (styleCell)
			{
			case 0:
				if (char.IsDigit(e.KeyChar) || char.IsControl(e.KeyChar) || Convert.ToInt32(e.KeyChar) == 59 || Convert.ToInt32(e.KeyChar) == 45)
				{
					e.Handled = false;
				}
				else
				{
					e.Handled = true;
				}
				break;
			case 1:
				if (char.IsDigit(e.KeyChar) || char.IsControl(e.KeyChar))
				{
					e.Handled = false;
				}
				else
				{
					e.Handled = true;
				}
				break;
			case 2:
				if (char.IsDigit(e.KeyChar) || char.IsControl(e.KeyChar) || Convert.ToInt32(e.KeyChar) == 59)
				{
					e.Handled = false;
				}
				else
				{
					e.Handled = true;
				}
				break;
			default:
				e.Handled = true;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static double CountSoLuongHD(string input)
	{
		double num = 0.0;
		try
		{
			string[] array = input.Split(new char[1] { ';' });
			string[] array2 = array;
			foreach (string text in array2)
			{
				if (!text.Contains("-"))
				{
					num += 1.0;
					continue;
				}
				string[] array3 = text.Split(new char[1] { '-' });
				num = num + (ConvertHelpers.ConvertToDouble(array3[1]) - ConvertHelpers.ConvertToDouble(array3[0])) + 1.0;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return num;
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

	public static string GetStrByLengh(string strInput, int strLengh)
	{
		try
		{
			string text = strInput.Trim();
			if (HTKKFormater.isNullString(text))
			{
				return null;
			}
			if (text.Length > strLengh)
			{
				return text;
			}
			return text.PadLeft(strLengh, '0');
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static string GetCellValild(string cellFomular)
	{
		try
		{
			string text = "";
			string[] array = cellFomular.Replace("'", "").Split(new char[1] { '!' });
			string text2 = "";
			string text3 = "";
			char[] array2 = array[1].ToCharArray();
			int num = array2.Length - 1;
			while (num >= 0)
			{
				int num2 = -1;
				try
				{
					num2 = int.Parse(array2[num].ToString());
				}
				catch
				{
					num2 = -1;
				}
				if (num2 >= 0)
				{
					text3 = array2[num] + text3;
					num--;
					continue;
				}
				text2 = array[1].Replace(text3, "");
				break;
			}
			return array[0] + "#" + text3 + "#" + text2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static void CheckTuSoDenSo(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, long iRow, long iCol, int iCol_TuSo, int iCol_DenSo, int iCol_SoLuong, int iCol_MauSoHD, int iCol_KyHieuHD, Dictionary<string, string[]> dicDMHoaDon)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			string text = "";
			if (iRow - 1 < num || iRow - 1 > num2)
			{
				return;
			}
			if (iCol - 1 == iCol_TuSo || iCol - 1 == iCol_DenSo)
			{
				text = sheetView.Cells[(int)iCol - 1, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = FormatSoHD(text);
				string text2 = sheetView.Cells[(int)iRow - 1, iCol_TuSo].Text;
				string text3 = sheetView.Cells[(int)iRow - 1, iCol_DenSo].Text;
				if (!string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text3))
				{
					double num3 = ConvertHelpers.ConvertToDouble(text2);
					double num4 = ConvertHelpers.ConvertToDouble(text3);
					if (num3 == 0.0 && num4 == 0.0)
					{
						sheetView.Cells[(int)iRow - 1, iCol_SoLuong].Text = "0";
					}
					else if (num3 == 0.0 && num4 != 0.0)
					{
						sheetView.Cells[(int)iRow - 1, iCol_SoLuong].Text = (num4 - num3).ToString();
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, iCol_SoLuong].Text = (num4 - num3 + 1.0).ToString();
					}
				}
			}
			if (iCol - 1 == iCol_MauSoHD)
			{
				string text4 = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Text;
				string text5 = "";
				if (!string.IsNullOrEmpty(text4))
				{
					text5 = dicDMHoaDon[text4][0];
				}
				if (!text5.Equals("3"))
				{
					string text6 = sheetView.Cells[(int)iRow - 1, iCol_MauSoHD].Text;
					sheetView.Cells[(int)iRow - 1, iCol_MauSoHD].Text = text6.ToUpper();
				}
			}
			if (iCol - 1 == iCol_KyHieuHD)
			{
				string text7 = sheetView.Cells[(int)iRow - 1, iCol_KyHieuHD].Text;
				sheetView.Cells[(int)iRow - 1, iCol_KyHieuHD].Text = FormatKyHieu(text7);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetTabStripColor(SheetView sView, int indexSheet, int iwarningColor)
	{
		if (!(sView.FpSpread.TabStrip[indexSheet].BackColor == Color.Pink))
		{
			switch (iwarningColor)
			{
			case 2:
				sView.FpSpread.TabStrip[indexSheet].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[indexSheet].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[indexSheet].BackColor = Color.White;
				break;
			}
		}
	}

	internal static int GetCatalogueFileNam(DateTime dateTime)
	{
		throw new NotImplementedException();
	}

	public static Dictionary<string, string> GetQuyNamByThangNam(int thang, int nam, string ngayBatDauTC)
	{
		if (HTKKFormater.isNullString(ngayBatDauTC) || ngayBatDauTC.IndexOf("/") == -1)
		{
			return null;
		}
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		try
		{
			int num = ConvertHelpers.ConvertToInt(ngayBatDauTC.Split(new char[1] { '/' })[1]);
			if (thang < num)
			{
				nam--;
			}
			if (num <= 12 && num > 9)
			{
				HTKKFile.AddOrUpdateDicStr(dictionary, "4", nam.ToString());
				HTKKFile.AddOrUpdateDicStr(dictionary, "1", (nam + 1).ToString());
				HTKKFile.AddOrUpdateDicStr(dictionary, "2", (nam + 1).ToString());
				HTKKFile.AddOrUpdateDicStr(dictionary, "3", (nam + 1).ToString());
			}
			if (num <= 9 && num > 6)
			{
				HTKKFile.AddOrUpdateDicStr(dictionary, "3", nam.ToString());
				HTKKFile.AddOrUpdateDicStr(dictionary, "4", nam.ToString());
				HTKKFile.AddOrUpdateDicStr(dictionary, "1", (nam + 1).ToString());
				HTKKFile.AddOrUpdateDicStr(dictionary, "2", (nam + 1).ToString());
			}
			if (num <= 6 && num > 3)
			{
				HTKKFile.AddOrUpdateDicStr(dictionary, "2", nam.ToString());
				HTKKFile.AddOrUpdateDicStr(dictionary, "3", nam.ToString());
				HTKKFile.AddOrUpdateDicStr(dictionary, "4", nam.ToString());
				HTKKFile.AddOrUpdateDicStr(dictionary, "1", (nam + 1).ToString());
			}
			if (num <= 3 && num > 0)
			{
				HTKKFile.AddOrUpdateDicStr(dictionary, "1", nam.ToString());
				HTKKFile.AddOrUpdateDicStr(dictionary, "2", nam.ToString());
				HTKKFile.AddOrUpdateDicStr(dictionary, "3", nam.ToString());
				HTKKFile.AddOrUpdateDicStr(dictionary, "4", nam.ToString());
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return dictionary;
	}

	public static List<string> GetQuyByNamTC(int thang, int nam, string ngayBatDauTC)
	{
		if (HTKKFormater.isNullString(ngayBatDauTC) || ngayBatDauTC.IndexOf("/") == -1)
		{
			return null;
		}
		List<string> list = new List<string>();
		try
		{
			int num = ConvertHelpers.ConvertToInt(ngayBatDauTC.Split(new char[1] { '/' })[1]);
			if (num <= 12 && num > 9)
			{
				if (thang <= 12 && thang > 9)
				{
					list.Add("1");
					list.Add(nam.ToString());
				}
				if (thang <= 9 && thang > 6)
				{
					list.Add("4");
					list.Add((nam - 1).ToString());
				}
				if (thang <= 6 && thang > 3)
				{
					list.Add("3");
					list.Add((nam - 1).ToString());
				}
				if (thang <= 3 && thang > 0)
				{
					list.Add("2");
					list.Add((nam - 1).ToString());
				}
			}
			if (num <= 9 && num > 6)
			{
				if (thang <= 12 && thang > 9)
				{
					list.Add("2");
					list.Add(nam.ToString());
				}
				if (thang <= 9 && thang > 6)
				{
					list.Add("1");
					list.Add(nam.ToString());
				}
				if (thang <= 6 && thang > 3)
				{
					list.Add("4");
					list.Add((nam - 1).ToString());
				}
				if (thang <= 3 && thang > 0)
				{
					list.Add("3");
					list.Add((nam - 1).ToString());
				}
			}
			if (num <= 6 && num > 3)
			{
				if (thang <= 12 && thang > 9)
				{
					list.Add("3");
					list.Add(nam.ToString());
				}
				if (thang <= 9 && thang > 6)
				{
					list.Add("2");
					list.Add(nam.ToString());
				}
				if (thang <= 6 && thang > 3)
				{
					list.Add("1");
					list.Add(nam.ToString());
				}
				if (thang <= 3 && thang > 0)
				{
					list.Add("4");
					list.Add((nam - 1).ToString());
				}
			}
			if (num <= 3 && num > 0)
			{
				if (thang <= 12 && thang > 9)
				{
					list.Add("4");
					list.Add(nam.ToString());
				}
				if (thang <= 9 && thang > 6)
				{
					list.Add("3");
					list.Add(nam.ToString());
				}
				if (thang <= 6 && thang > 3)
				{
					list.Add("2");
					list.Add(nam.ToString());
				}
				if (thang <= 3 && thang > 0)
				{
					list.Add("1");
					list.Add(nam.ToString());
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	public static bool CheckThongTinNNT(HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			Type typeFromHandle = typeof(HTKKHeaderInfo);
			bool flag = ConvertHelpers.ConvertToBool(HeaderInfo.ThongTinDonViChuQuan);
			bool flag2 = ConvertHelpers.ConvertToBool(HeaderInfo.ThongTinDaiLy);
			PropertyInfo[] properties = typeFromHandle.GetProperties();
			string text = "~tenNNT~dchiNNT~tenXaPhuongNNT~tenTinhNNT~NgayBatDauNamTaiChinh~tenCQTCapCuc~tenCQTNoiNop~maCQTCapCuc~maCQTNoiNop~";
			string text2 = "~TenDonViChuQuan~MaSoThueDonViChuQuan~";
			string text3 = "~mstDLyThue~tenDLyThue~dchiDLyThue~tenXaPhuongDLyThue~tenTinhDLyThue~soHDongDLyThue~ngayKyHDDLyThue~tenNVienDLyThue~cchiHNghe~";
			PropertyInfo[] array = properties;
			foreach (PropertyInfo propertyInfo in array)
			{
				if (text.IndexOf("~" + propertyInfo.Name + "~") > -1 && HTKKFormater.isNullString(propertyInfo.GetValue(HeaderInfo, null).ToString()))
				{
					return false;
				}
				if (flag && text2.IndexOf("~" + propertyInfo.Name + "~") > -1)
				{
					if (HTKKFormater.isNullString(propertyInfo.GetValue(HeaderInfo, null).ToString()))
					{
						return false;
					}
					if (propertyInfo.Name == "MaSoThueDonViChuQuan" && !ValidateBase.CheckMaSoThue(propertyInfo.GetValue(HeaderInfo, null).ToString()))
					{
						return false;
					}
				}
				if (flag2 && text3.IndexOf("~" + propertyInfo.Name + "~") > -1)
				{
					if (HTKKFormater.isNullString(propertyInfo.GetValue(HeaderInfo, null).ToString()))
					{
						return false;
					}
					if (propertyInfo.Name == "mstDLyThue" && !ValidateBase.CheckMaSoThue(propertyInfo.GetValue(HeaderInfo, null).ToString()))
					{
						return false;
					}
				}
			}
			if (HeaderInfo.mst == HeaderInfo.mstDLyThue && flag2)
			{
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
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
				result = ((num4 != num2) ? ((num4 - num2) * 12 + (num3 - num) + 1) : (num3 - num + 1));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = 0;
		}
		return result;
	}

	public static string getMaHoaDon(string kyHieuMau)
	{
		string result = "";
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = "";
		}
		return result;
	}

	public static List<item> getDmTaiNguyenForTinh(List<string> listPaths, HTKKPeriodInfo PeriodInfo)
	{
		List<item> list = new List<item>();
		try
		{
			if (listPaths != null && listPaths.Count > 0)
			{
				foreach (string listPath in listPaths)
				{
					if (!File.Exists(listPath))
					{
						continue;
					}
					HTKKSerializer.Serializer serializer = new HTKKSerializer.Serializer();
					string input = File.ReadAllText(listPath);
					danhmuc danhmuc = serializer.Deserialize<danhmuc>(input);
					if (danhmuc.items == null || danhmuc.items.Count <= 0)
					{
						continue;
					}
					foreach (item item in danhmuc.items)
					{
						if (!HTKKFormater.isNullString(item.cqthltu) && !HTKKFormater.isNullString(item.cqthlden))
						{
							DateTime dateTime = DateTime.ParseExact(item.cqthltu, "dd/MM/yyyy", CultureInfo.CurrentCulture);
							DateTime dateTime2 = DateTime.ParseExact(item.cqthlden, "dd/MM/yyyy", CultureInfo.CurrentCulture);
							DateTime[] array = (DateTime[])PeriodInfo.dicVar["DATE_KY_KE_KHAI"];
							if (array != null && ((array[0] >= dateTime && array[0] <= dateTime2) || (dateTime >= array[0] && dateTime <= array[1])))
							{
								list.Add(item);
							}
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			list = null;
		}
		return list;
	}

	public static item getTaiNguyenForCombobox(item itCurrent, List<item> itChecks)
	{
		item item = null;
		try
		{
			item item2 = null;
			item item3 = null;
			string text = "";
			string text2 = "";
			bool flag = true;
			bool flag2 = true;
			for (int i = 0; i < itChecks.Count; i++)
			{
				item2 = itChecks[i];
				for (int j = i + 1; j < itChecks.Count; j++)
				{
					item3 = itChecks[j];
					if (itChecks[i].ma == itChecks[j].ma)
					{
						flag2 = false;
					}
					text = ((item2 == null || !(item2.cqtgia.ToNumber() > 0m)) ? itChecks[i].giatoithieu.ToStr() : item2.cqtgia);
					text2 = ((item3 == null || !(item3.cqtgia.ToNumber() > 0m)) ? itChecks[j].giatoithieu.ToStr() : item3.cqtgia);
					if (text == text2 && itChecks[i].thuesuat == itChecks[j].thuesuat)
					{
						flag = false;
					}
				}
			}
			if (!flag)
			{
				if (!flag2)
				{
					return null;
				}
				item = HTKKSerializer.Serializer.DeepCopy(itCurrent);
				item.ten = item.ten + "(" + item.ma + ")";
			}
			else
			{
				item = HTKKSerializer.Serializer.DeepCopy(itCurrent);
				string text3 = "";
				string text4 = "";
				text3 = ((itCurrent.thuesuat.ToStr().Length <= 0) ? "0" : itCurrent.thuesuat.ToStr());
				text4 = ((itCurrent.cqtgia.ToStr().Length > 0) ? itCurrent.cqtgia.ToStr() : ((itCurrent.giatoithieu.ToStr().Length <= 0) ? "0" : itCurrent.giatoithieu.ToStr()));
				CultureInfo cultureInfo = CultureInfo.GetCultureInfo("vi-VN");
				item.ten = item.ten + "(" + text3 + "% - " + string.Format(cultureInfo, "{0:N0}", new object[1] { text4.ToNumber() }) + "đ)";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return item;
	}

	public static void InsertTableRow(SheetView sView, int frRow, int frCol, int toRow, int toCol, int colCount, HTKKCMCSheetInfo sInfo, int iCurTable)
	{
		try
		{
			string sheetName = sView.SheetName;
			sView.Rows.Add(frRow, 1);
			sView.CopyRange(frRow - 1, frCol, toRow, toCol, 1, colCount, dataOnly: false);
			sInfo.IncrementStaticCell(frRow, 1);
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (table.Location.DataLocationTemplateRowStart > frRow)
				{
					table.Increment(1);
				}
				table.IncrementFooter(frRow, 1);
			}
			for (int i = 0; i < sInfo.Tables[iCurTable].RowsInfo.Cells.Count; i++)
			{
				if (sView.Cells[frRow, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex - 1].BackColor == Color.Pink || sView.Cells[frRow, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex - 1].BackColor == Color.Yellow)
				{
					sView.Cells[frRow, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex - 1].ResetBackColor();
				}
				sView.Cells[frRow, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex - 1].ResetNote();
				fpSpBindDataCell(sView, frRow + 1, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex, sInfo.Tables[iCurTable].RowsInfo.Cells[i].Controltype, sInfo.Tables[iCurTable].RowsInfo.Cells[i].DefaultValue);
			}
			sInfo.Tables[iCurTable].addNewDefaultRow(frRow - sInfo.Tables[iCurTable].Location.DataLocationTemplateRowStart + 1);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertTableRows(SheetView sView, int frRow, int frCol, int toRow, int toCol, int colCount, HTKKCMCSheetInfo sInfo, int iCurTable, int rowCount)
	{
		try
		{
			sView.Rows.Add(toRow, rowCount);
			sView.CopyRange(frRow - 1, frCol, toRow, toCol, rowCount, colCount, dataOnly: false);
			sInfo.IncrementStaticCell(frRow, rowCount);
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (table.Location.DataLocationTemplateRowStart > frRow)
				{
					table.Increment(rowCount);
				}
				table.IncrementFooter(frRow, rowCount);
			}
			sInfo.Tables[iCurTable].addNewDefaultRow(rowCount);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpBindDataCell(SheetView sView, int iRow, int iCol, string style, string values)
	{
		try
		{
			if (!string.IsNullOrEmpty(values))
			{
				values = values.Trim();
			}
			if (iRow == 57 && iCol == 18)
			{
				int num = 0;
			}
			if (iRow <= 0 || iCol <= 0)
			{
				return;
			}
			switch (style)
			{
			case "2":
				if (values.ToUpper() == "TRUE" || values.ToUpper() == "X" || values == "1")
				{
					sView.Cells[iRow - 1, iCol - 1].Value = true;
				}
				else
				{
					sView.Cells[iRow - 1, iCol - 1].Value = false;
				}
				break;
			case "7":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				break;
			case "101":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				break;
			case "102":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				break;
			case "6":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				break;
			case "5":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				break;
			case "100":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				break;
			case "103":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				break;
			case "14":
				if (!sView.Cells[iRow - 1, iCol - 1].Locked && !string.IsNullOrEmpty(values) && !string.IsNullOrEmpty(sView.Cells[iRow - 1, iCol - 1].Formula))
				{
					sView.Cells[iRow - 1, iCol - 1].ResetFormula();
				}
				InsertDateCell(sView, iRow, iCol, values);
				break;
			case "96":
				if (!sView.Cells[iRow - 1, iCol - 1].Locked && !string.IsNullOrEmpty(values) && !string.IsNullOrEmpty(sView.Cells[iRow - 1, iCol - 1].Formula))
				{
					sView.Cells[iRow - 1, iCol - 1].ResetFormula();
				}
				InsertTimeCell(sView, iRow, iCol, values);
				break;
			case "94":
				if (!sView.Cells[iRow - 1, iCol - 1].Locked && !string.IsNullOrEmpty(values) && !string.IsNullOrEmpty(sView.Cells[iRow - 1, iCol - 1].Formula))
				{
					sView.Cells[iRow - 1, iCol - 1].ResetFormula();
				}
				InsertDateTimeCell(sView, iRow, iCol, values);
				break;
			case "33":
				if (!string.IsNullOrEmpty(values))
				{
					sView.Cells[iRow - 1, iCol - 1].Value = ConvertHelpers.ConvertToDouble(values);
				}
				else
				{
					sView.Cells[iRow - 1, iCol - 1].Value = values;
				}
				break;
			case "19":
				InsertMonthYearCell(sView, iRow, iCol, values);
				break;
			default:
				if (!sView.Cells[iRow - 1, iCol - 1].Locked && !string.IsNullOrEmpty(values) && !string.IsNullOrEmpty(sView.Cells[iRow - 1, iCol - 1].Formula))
				{
					sView.Cells[iRow - 1, iCol - 1].ResetFormula();
				}
				if (sView.Cells[iRow - 1, iCol - 1].CellType != null && sView.Cells[iRow - 1, iCol - 1].CellType.GetType().Name == "NumberCellType")
				{
					sView.Cells[iRow - 1, iCol - 1].Value = ConvertHelpers.ConvertToDouble(values);
				}
				else
				{
					sView.Cells[iRow - 1, iCol - 1].Text = values;
				}
				break;
			}
			if (!string.IsNullOrEmpty(values))
			{
				Row row = sView.Rows[iRow - 1];
				float preferredHeight = row.GetPreferredHeight();
				if (row.Height < preferredHeight - 5f || row.Height > preferredHeight + 5f)
				{
					row.Height = preferredHeight;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpReindexingTableTemplate(SheetView sView, HTKKCMCTableInfo tableInfo)
	{
		try
		{
			sView.FpSpread.SuspendLayout();
			for (int i = 0; i < tableInfo.RowsValues.Count; i++)
			{
				_ = tableInfo.iColumnSTT;
				bool flag = 1 == 0;
				if (tableInfo.CellTableSTT != null)
				{
					if (tableInfo.CellTableSTT.Value == "abc")
					{
						string[] array = "a~b~c~d~đ~e~g~h~i~k~l~m~n~o~p~q~r~s~t~u~v~x~y".Split(new char[1] { '~' }).ToArray();
						sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].ResetFormula();
						sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].Text = array[(i >= 22) ? 22 : i];
					}
					else
					{
						sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].ResetFormula();
						sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].Text = tableInfo.CellTableSTT.Value + "." + (i + 1);
					}
				}
				else
				{
					sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].ResetFormula();
					sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].Text = (i + 1).ToString();
				}
			}
			sView.FpSpread.ResumeLayout();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertDateCell(SheetView sView, int iRow, int iCol, string value)
	{
		try
		{
			string text = HTKKFormater.toDateDDMMYYYY(value, isMess: false);
			if (!HTKKFormater.isNullString(text))
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/00/0000";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = text;
			}
			else
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/00/0000";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertTimeCell(SheetView sView, int iRow, int iCol, string value)
	{
		try
		{
			string text = HTKKFormater.toTimeHHMMSS(value, isMess: false);
			if (!HTKKFormater.isNullString(text))
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00:00:00";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = text;
			}
			else
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00:00:00";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertDateTimeCell(SheetView sView, int iRow, int iCol, string value)
	{
		try
		{
			string text = HTKKFormater.toDateAndTimeDDMMYYYYHHMMSS(value, isMess: false);
			if (!HTKKFormater.isNullString(text))
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/00/0000 00:00:00";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = text;
			}
			else
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/00/0000 00:00:00";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertMonthYearCell(SheetView sView, int iRow, int iCol, string value)
	{
		try
		{
			string text = HTKKFormater.toDateMMYYYY(string.IsNullOrEmpty(value) ? "" : value, isMess: false);
			if (!HTKKFormater.isNullString(text))
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/0000";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = text;
			}
			else
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/0000";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void RemoveDataRows(SheetView sView, int tableIndex, int frRow, int rowCount, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (rowCount <= 0)
			{
				return;
			}
			sView.Rows.Remove(frRow, rowCount);
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (table.Location.DataLocationTemplateRowStart > frRow + 1)
				{
					table.Increment(rowCount * -1);
				}
			}
			sInfo.IncrementStaticCell(frRow, rowCount * -1);
			int num = sInfo.Tables[tableIndex].RowsValues.Count - 1;
			for (int num2 = num; num2 > num - rowCount; num2--)
			{
				sInfo.Tables[tableIndex].RowsValues.RemoveAt(num2);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void RemoveDataOfTable(SheetView sView, int tableIndex, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[tableIndex];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			if (hTKKCMCTableInfo.RowsValues.Count > 1)
			{
				RemoveDataRows(sView, tableIndex, dataLocationTemplateRowStart - 1, sInfo.Tables[tableIndex].RowsValues.Count - 1, sInfo);
			}
			foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo.RowsInfo.Cells)
			{
				sView.Cells[dataLocationTemplateRowStart - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void LockUnLockCols(SheetView sView, int row, List<int> cacCotAn, bool locked, bool deleteData)
	{
		try
		{
			foreach (int item in cacCotAn)
			{
				sView.Cells[row - 1, item - 1].Locked = locked;
				if (deleteData)
				{
					sView.Cells[row - 1, item - 1].Text = "";
				}
				sView.Cells[row - 1, item - 1].BackColor = (locked ? ColorTranslator.FromHtml(MA_COLOR_LOCKED) : ColorTranslator.FromHtml(MA_COLOR_UNLOCK));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void LockUnLockColsInTable(SheetView sView, HTKKCMCTableInfo table, List<int> cacCotAn, bool locked, bool deleteData, bool lockAllHeader = false)
	{
		try
		{
			if (lockAllHeader)
			{
				foreach (HTKKCMCCellInfo staticCell in table.Header.StaticCells)
				{
					LockUnLockCol(sView, staticCell.CellTemplate.RowIndex, staticCell.CellTemplate.ColIndex, locked, deleteData);
				}
			}
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				foreach (int col in cacCotAn)
				{
					List<HTKKCMCTableCellInfo> cells = table.RowsInfo.Cells;
					Predicate<HTKKCMCTableCellInfo> match = (HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == col;
					int index = cells.FindIndex(match);
					sView.Cells[dataLocationTemplateRowStart + i - 1, col - 1].Locked = locked;
					if (deleteData)
					{
						sView.Cells[dataLocationTemplateRowStart + i - 1, col - 1].Text = table.RowsInfo.Cells[index].DefaultValue;
					}
					sView.Cells[dataLocationTemplateRowStart + i - 1, col - 1].BackColor = (locked ? ColorTranslator.FromHtml(MA_COLOR_LOCKED) : ColorTranslator.FromHtml(MA_COLOR_UNLOCK));
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void LockUnLockCol(SheetView sView, int row, int colHide, bool locked, bool deleteData)
	{
		try
		{
			NumberCellType numberCellType = new NumberCellType();
			CurrencyCellType currencyCellType = new CurrencyCellType();
			CheckBoxCellType checkBoxCellType = new CheckBoxCellType();
			sView.Cells[row - 1, colHide - 1].Locked = locked;
			if (deleteData && locked)
			{
				ICellType cellType = sView.Cells[row - 1, colHide - 1].CellType;
				if (cellType != null)
				{
					if ((object)cellType.GetType() == numberCellType.GetType() || (object)cellType.GetType() == currencyCellType.GetType())
					{
						sView.Cells[row - 1, colHide - 1].Value = 0;
					}
					else if ((object)cellType.GetType() == checkBoxCellType.GetType())
					{
						sView.Cells[row - 1, colHide - 1].Value = false;
					}
					else
					{
						sView.Cells[row - 1, colHide - 1].Text = "";
					}
				}
			}
			sView.Cells[row - 1, colHide - 1].BackColor = (locked ? ColorTranslator.FromHtml(MA_COLOR_LOCKED) : ColorTranslator.FromHtml(MA_COLOR_UNLOCK));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void LockUnLockCol2(SheetView sView, int row, int colHide, bool locked, bool deleteData, object deleteToData)
	{
		try
		{
			sView.Cells[row - 1, colHide - 1].Locked = locked;
			sView.Cells[row - 1, colHide - 1].BackColor = (locked ? ColorTranslator.FromHtml(MA_COLOR_LOCKED) : ColorTranslator.FromHtml(MA_COLOR_UNLOCK));
			if (deleteData)
			{
				sView.Cells[row - 1, colHide - 1].Value = deleteToData;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
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

	public static void InsertCombobox(SheetView sView, int iRow, int iCol, string[] Values, bool deletText = true)
	{
		try
		{
			if (iCol >= 1 && iRow >= 1)
			{
				ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
				comboBoxCellType.Items = Values;
				comboBoxCellType.MaxDrop = 10;
				comboBoxCellType.ListWidth = 500;
				sView.Cells[iRow - 1, iCol - 1].CellType = comboBoxCellType;
				if (deletText)
				{
					sView.Cells[iRow - 1, iCol - 1].Text = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void TinhSttTheoDieuKienBool(SheetView sView, HTKKCMCSheetInfo sheet, int indexTable, int colStt, int colSoSanh, bool valueSoSanh)
	{
		try
		{
			if (sheet == null)
			{
				return;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = sheet.Tables[indexTable];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			int num = 0;
			int num2 = 0;
			for (int i = 0; i < rowsValues.Count; i++)
			{
				bool flag = Convert.ToBoolean(sView.Cells[dataLocationTemplateRowStart + i - 1, colSoSanh - 1].Value);
				if (flag == valueSoSanh || i == 0)
				{
					num++;
					num2 = 0;
					sView.Cells[dataLocationTemplateRowStart + i - 1, colStt - 1].Text = num.ToString();
				}
				else
				{
					num2++;
					sView.Cells[dataLocationTemplateRowStart + i - 1, colStt - 1].Text = num + "." + num2;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static Dictionary<string, List<int>> GroupTableByStt(SheetView sView, HTKKCMCSheetInfo sheet, int indexTable, int colStt)
	{
		Dictionary<string, List<int>> dictionary = new Dictionary<string, List<int>>();
		if (sheet == null)
		{
			return dictionary;
		}
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sheet.Tables[indexTable];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			for (int i = 0; i < rowsValues.Count; i++)
			{
				int num = dataLocationTemplateRowStart + i;
				string text = sView.Cells[num - 1, colStt - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text))
				{
					string[] array = text.Split(new char[1] { '.' });
					if (dictionary.ContainsKey(array[0]))
					{
						dictionary[array[0]].Add(num);
						continue;
					}
					List<int> list = new List<int>();
					list.Add(num);
					dictionary.Add(array[0], list);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return dictionary;
	}

	public static Dictionary<string, List<int>> GroupTableByCol(SheetView sView, HTKKCMCSheetInfo sInfo, int indexTable, int col)
	{
		Dictionary<string, List<int>> dictionary = new Dictionary<string, List<int>>();
		if (sInfo == null)
		{
			return dictionary;
		}
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[indexTable];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			for (int i = 0; i < rowsValues.Count; i++)
			{
				int num = dataLocationTemplateRowStart + i;
				string text = sView.Cells[num - 1, col - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text))
				{
					if (dictionary.ContainsKey(text))
					{
						dictionary[text].Add(num);
						continue;
					}
					List<int> list = new List<int>();
					list.Add(num);
					dictionary.Add(text, list);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return dictionary;
	}

	public static bool DetectIsRowNew(SheetView sView, int newRow, int colSoSanh1, bool valueSoSanh1, int colSoSanh2, string valueSoSanh2)
	{
		try
		{
			bool flag = Convert.ToBoolean(sView.Cells[newRow - 1, colSoSanh1 - 1].Value);
			string text = sView.Cells[newRow - 1, colSoSanh2 - 1].Text.Trim();
			if (flag == valueSoSanh1 && text == valueSoSanh2)
			{
				return true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return false;
	}

	public static double CalculateSumByGroup(SheetView sView, List<HTKKCMCTableInfo> tables, int colSum, bool isView = true)
	{
		double num = 0.0;
		try
		{
			foreach (HTKKCMCTableInfo table in tables)
			{
				int num2 = -1;
				if (table.Header != null && table.Header.StaticCells.Count > 0)
				{
					num2 = table.Header.StaticCells[0].CellTemplate.RowIndex;
				}
				if (num2 >= 0)
				{
					double num3 = CalculateColSumTable(sView, table, colSum);
					num += num3;
					if (isView)
					{
						sView.Cells[num2 - 1, colSum - 1].Value = num3;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return num;
	}

	public static double CalculateColSumTable(SheetView sView, HTKKCMCTableInfo table, int colSum)
	{
		double num = 0.0;
		try
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				double num2 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + i - 1, colSum - 1].Value);
				num += num2;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return num;
	}

	public static int DetectIndexOfTable(HTKKCMCSheetInfo sInfo, int row)
	{
		try
		{
			if (sInfo == null || sInfo.Tables.Count == 0)
			{
				return -1;
			}
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				try
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
					HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
					int num = location.DataLocationTemplateRowStart;
					int num2 = 0;
					if (hTKKCMCTableInfo.Header != null && hTKKCMCTableInfo.Header.StaticCells != null && hTKKCMCTableInfo.Header.StaticCells.Count > 0)
					{
						num = location.TableTemplateRowStart;
						num2 = num + (hTKKCMCTableInfo.TableLength - 1) + hTKKCMCTableInfo.RowsValues.Count;
					}
					else
					{
						num2 = num + hTKKCMCTableInfo.RowsValues.Count;
					}
					if (row >= num && row <= num2)
					{
						return i;
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					return -1;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return -1;
	}

	public static int DetectIndexOfSheet(List<HTKKCMCSheetInfo> sInfos, string formula)
	{
		try
		{
			string[] array = formula.Split(new char[1] { '\'' });
			for (int i = 0; i < sInfos.Count; i++)
			{
				if (sInfos[i] != null)
				{
					string sheetName = sInfos[i].SheetName;
					if (!string.IsNullOrEmpty(sheetName) && formula.Contains(sheetName))
					{
						return i;
					}
					if (formula.Contains("PL_GiamThue_GTGT_23_24") && formula.Contains(sInfos[i].SheetName))
					{
						return i;
					}
				}
			}
			return -1;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1;
		}
	}

	public static void CopyPaseTableToTable(HTKKCMCTableInfo tableFrom, HTKKCMCTableInfo tableTo, SheetView sViewFrom, SheetView sViewTo, Dictionary<int, int> colsCanTinhFromAndTo, int colMaFrom, int colMaTo, int colTenFrom, int colTenTo, string keyDefault, HTKKCMCSheetInfo sInfo, int colCompare = -1, string valueCompare = "", bool usingCellType = false)
	{
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			Dictionary<string, Dictionary<int, CellTypeKHBS>> dictionary2 = new Dictionary<string, Dictionary<int, CellTypeKHBS>>();
			int dataLocationTemplateRowStart = tableFrom.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < tableFrom.RowsValues.Count; i++)
			{
				if (colCompare >= 0)
				{
					string text = sViewFrom.Cells[dataLocationTemplateRowStart + i - 1, colCompare - 1].Text.Trim();
					if (text != valueCompare)
					{
						continue;
					}
				}
				string text2 = sViewFrom.Cells[dataLocationTemplateRowStart + i - 1, colMaFrom - 1].Text.Trim();
				if (text2 == keyDefault)
				{
					continue;
				}
				string value = sViewFrom.Cells[dataLocationTemplateRowStart + i - 1, colTenFrom - 1].Text.Trim();
				if (!dictionary.ContainsKey(text2))
				{
					dictionary.Add(text2, value);
				}
				else
				{
					dictionary[text2] = value;
				}
				foreach (KeyValuePair<int, int> item in colsCanTinhFromAndTo)
				{
					double num = Convert.ToDouble(sViewFrom.Cells[dataLocationTemplateRowStart + i - 1, item.Key - 1].Value);
					Dictionary<int, CellTypeKHBS> dictionary3;
					CellTypeKHBS cellTypeKHBS;
					if (!dictionary2.ContainsKey(text2))
					{
						dictionary3 = new Dictionary<int, CellTypeKHBS>();
						cellTypeKHBS = new CellTypeKHBS();
						cellTypeKHBS.value = num;
						cellTypeKHBS.iCellType = sViewFrom.Cells[dataLocationTemplateRowStart + i - 1, item.Key - 1].CellType;
						dictionary3.Add(item.Value, cellTypeKHBS);
						dictionary2.Add(text2, dictionary3);
						continue;
					}
					dictionary3 = new Dictionary<int, CellTypeKHBS>();
					cellTypeKHBS = new CellTypeKHBS();
					cellTypeKHBS.value = num;
					cellTypeKHBS.iCellType = sViewFrom.Cells[dataLocationTemplateRowStart + i - 1, item.Key - 1].CellType;
					if (!dictionary2[text2].ContainsKey(item.Key))
					{
						dictionary2[text2].Add(item.Value, cellTypeKHBS);
					}
					else
					{
						dictionary2[text2][item.Key].value = Convert.ToDouble(dictionary2[text2][item.Key].value) + num;
					}
				}
			}
			int dataLocationTemplateRowStart2 = tableTo.Location.DataLocationTemplateRowStart;
			int num2 = DetectIndexOfTable(sInfo, dataLocationTemplateRowStart2);
			if (num2 < 0)
			{
				return;
			}
			if (sInfo.Tables[num2].RowsValues.Count > 1)
			{
				RemoveDataRows(sViewTo, num2, dataLocationTemplateRowStart2 - 1, sInfo.Tables[num2].RowsValues.Count - 1, sInfo);
			}
			foreach (HTKKCMCTableCellInfo cell in sInfo.Tables[num2].RowsInfo.Cells)
			{
				sViewTo.Cells[dataLocationTemplateRowStart2 - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
			}
			int num3 = 0;
			foreach (KeyValuePair<string, string> item2 in dictionary)
			{
				int num4 = DetectRowByValue(sViewTo, tableTo, colMaTo, item2.Key);
				if (num4 < 0)
				{
					if (num3 > 0)
					{
						num4 = dataLocationTemplateRowStart2 + num3;
						InsertTableRow(sViewTo, num4 - 1, -1, num4 - 1, -1, -1, sInfo, num2);
					}
					else
					{
						num4 = dataLocationTemplateRowStart2;
					}
				}
				sViewTo.Cells[num4 - 1, colMaTo - 1].Text = item2.Key;
				sViewTo.Cells[num4 - 1, colTenTo - 1].Text = item2.Value;
				foreach (KeyValuePair<int, CellTypeKHBS> item3 in dictionary2[item2.Key])
				{
					sViewTo.Cells[num4 - 1, item3.Key - 1].Value = item3.Value.value;
					try
					{
						if (usingCellType)
						{
							sViewTo.Cells[num4 - 1, item3.Key - 1].CellType = item3.Value.iCellType;
						}
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					}
				}
				num3++;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static int DetectRowByValue(SheetView sView, HTKKCMCTableInfo table, int colCheck, string valueSoSanh)
	{
		try
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				string text = sView.Cells[dataLocationTemplateRowStart + i - 1, colCheck - 1].Text.Trim();
				if (string.IsNullOrEmpty(text))
				{
					return -1;
				}
				if (text == valueSoSanh)
				{
					return dataLocationTemplateRowStart + i;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return -1;
	}

	public static List<int> GetRowIndexInSViewByCompare(SheetView sView, HTKKCMCTableInfo table, int colCheck, string valueCompare)
	{
		List<int> list = new List<int>();
		try
		{
			if (table == null)
			{
				return list;
			}
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count && dataLocationTemplateRowStart + i <= sView.RowCount; i++)
			{
				if (sView.Cells[dataLocationTemplateRowStart + i - 1, colCheck - 1].Value != null && sView.Cells[dataLocationTemplateRowStart + i - 1, colCheck - 1].Value.ToString() == valueCompare)
				{
					list.Add(dataLocationTemplateRowStart + i);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	public static List<int> GetRowIndexInSViewByCompareHeader(SheetView sView, List<HTKKCMCTableInfo> tables, int colCheck, string valueCompare)
	{
		List<int> list = new List<int>();
		try
		{
			foreach (HTKKCMCTableInfo table in tables)
			{
				if (table.Header != null && table.Header.StaticCells != null)
				{
					int rowIndex = table.Header.StaticCells[0].CellTemplate.RowIndex;
					if (sView.Cells[rowIndex - 1, colCheck - 1].Value != null && sView.Cells[rowIndex - 1, colCheck - 1].Value.ToString() == valueCompare)
					{
						list.Add(rowIndex);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	public static List<int> GetRowIndexInSViewByCompares(SheetView sView, HTKKCMCTableInfo table, List<int> colChecks, List<string> valueCompares)
	{
		List<int> list = new List<int>();
		try
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			MaskCellType maskCellType = new MaskCellType();
			NumberCellType numberCellType = new NumberCellType();
			CurrencyCellType currencyCellType = new CurrencyCellType();
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				int num = 0;
				for (int j = 0; j < colChecks.Count; j++)
				{
					ICellType cellType = sView.Cells[dataLocationTemplateRowStart + i - 1, colChecks[j] - 1].CellType;
					object value = sView.Cells[dataLocationTemplateRowStart + i - 1, colChecks[j] - 1].Value;
					if ((object)cellType.GetType() == maskCellType.GetType())
					{
						if (value != null && value.ToString() == valueCompares[j].Replace("/", ""))
						{
							num++;
						}
						if (sView.Cells[dataLocationTemplateRowStart + i - 1, colChecks[j] - 1].Value == null && "" == valueCompares[j].Replace("/", ""))
						{
							num++;
						}
					}
					else if ((object)cellType.GetType() == numberCellType.GetType() || (object)cellType.GetType() == currencyCellType.GetType())
					{
						if (value != null && !string.IsNullOrEmpty(valueCompares[j]) && Convert.ToDouble(value) == Convert.ToDouble(valueCompares[j].Replace(",", ".")))
						{
							num++;
						}
						if (value == null && "" == valueCompares[j])
						{
							num++;
						}
					}
					else
					{
						if (value != null && value.ToString() == valueCompares[j])
						{
							num++;
						}
						if (value == null && "" == valueCompares[j])
						{
							num++;
						}
					}
				}
				if (num == colChecks.Count)
				{
					list.Add(dataLocationTemplateRowStart + i);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	public static List<int> GetRowIndexInSViewByCompares2(SheetView sView, HTKKCMCTableInfo table, List<int> colChecks, string valueCompares)
	{
		List<int> list = new List<int>();
		try
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			MaskCellType maskCellType = new MaskCellType();
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				int num = 0;
				for (int j = 0; j < colChecks.Count; j++)
				{
					if (sView.Cells[dataLocationTemplateRowStart + i - 1, colChecks[j] - 1].Value != null && valueCompares.IndexOf(sView.Cells[dataLocationTemplateRowStart + i - 1, colChecks[j] - 1].Value.ToString()) >= 0)
					{
						num++;
					}
				}
				if (num == colChecks.Count)
				{
					list.Add(dataLocationTemplateRowStart + i);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	public static List<int> GetValuesExistsCompare(List<int> list1, List<int> list2)
	{
		List<int> list3 = new List<int>();
		int i;
		for (i = 0; i < list1.Count; i++)
		{
			IEnumerable<int> source = list2.Where((int x) => x == list1[i]);
			if (source.Count() == 1)
			{
				list3.Add(list1[i]);
			}
		}
		return list3;
	}

	public static void TongCongTheoTungTinh(HTKKCMCSheetInfo toSInfo, SheetView fromSView, HTKKCMCTableInfo fromTable, int fromColMaTinh, int formColTenTinh, SheetView ToSView, HTKKCMCTableInfo toTable, int toTableIndex, int toColMaTinh, int toColTenTinh, Dictionary<int, int> colsSum)
	{
		try
		{
			int dataLocationTemplateRowStart = fromTable.Location.DataLocationTemplateRowStart;
			int dataLocationTemplateRowStart2 = toTable.Location.DataLocationTemplateRowStart;
			if (toTable.RowsValues.Count > 1)
			{
				RemoveDataRows(ToSView, toTableIndex, dataLocationTemplateRowStart2 - 1, toTable.RowsValues.Count - 1, toSInfo);
			}
			foreach (KeyValuePair<int, int> item in colsSum)
			{
				ToSView.Cells[dataLocationTemplateRowStart2 - 1, item.Value - 1].Value = 0;
			}
			for (int i = 0; i < fromTable.RowsValues.Count; i++)
			{
				string text = fromSView.Cells[dataLocationTemplateRowStart + i - 1, fromColMaTinh - 1].Text.Trim();
				string text2 = fromSView.Cells[dataLocationTemplateRowStart + i - 1, formColTenTinh - 1].Text.Trim();
				if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(text))
				{
					continue;
				}
				int num = DetectRowByValue(ToSView, toTable, toColMaTinh, text);
				if (num < 0)
				{
					if (i == 0)
					{
						num = dataLocationTemplateRowStart2;
					}
					else
					{
						num = dataLocationTemplateRowStart2 + toTable.RowsValues.Count;
						InsertTableRow(ToSView, num - 1, -1, num - 1, -1, -1, toSInfo, toTableIndex);
					}
				}
				ToSView.Cells[num - 1, toColMaTinh - 1].Text = text;
				ToSView.Cells[num - 1, toColTenTinh - 1].Text = text2;
				foreach (KeyValuePair<int, int> item2 in colsSum)
				{
					double num2 = Convert.ToDouble(ToSView.Cells[num - 1, item2.Value - 1].Value);
					double num3 = Convert.ToDouble(fromSView.Cells[dataLocationTemplateRowStart + i - 1, item2.Value - 1].Value);
					ToSView.Cells[num - 1, item2.Value - 1].Value = num2 + num3;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void checkLockUnLockCells(HTKKCMCToKhaiInfo ToKhaiInfo, FpSpread spread, int iRowCheck = -1, int tableIndexCheck = -1)
	{
		try
		{
			int index = spread.Sheets.Count - 1;
			SheetView sheetView = spread.Sheets[index];
			int[] array = ArrayCellId("BA_1");
			if (sheetView.Cells[array[1], array[0]].Text.Trim() != "1")
			{
				return;
			}
			sheetView.RecalculateAll();
			sheetView.AutoCalculation = true;
			sheetView.LoadFormulas(recalculate: true);
			int num = HTKKFile.ColumnLettersToIndex("AK");
			int num2 = HTKKFile.ColumnLettersToIndex("AL");
			int num3 = HTKKFile.ColumnLettersToIndex("AM");
			int num4 = HTKKFile.ColumnLettersToIndex("AN");
			int column = HTKKFile.ColumnLettersToIndex("BA");
			for (int i = 1; i < sheetView.NonEmptyRowCount; i++)
			{
				string formula = sheetView.Cells[i, num - 1].Formula;
				if (string.IsNullOrEmpty(formula))
				{
					break;
				}
				string[] array2 = GetCellValild(formula).Split(new char[1] { '#' });
				if (array2.Length != 3)
				{
					continue;
				}
				int num5 = -1;
				int num6 = Convert.ToInt32(array2[1]);
				int num7 = DetectIndexOfSheet(ToKhaiInfo.SheetInfos, formula);
				if (num7 < 0)
				{
					continue;
				}
				num5 = DetectIndexOfTable(ToKhaiInfo.SheetInfos[num7], num6);
				SheetView sheetView2;
				string formula3;
				string text;
				string[] array3;
				if (num5 >= 0 && (tableIndexCheck < 0 || tableIndexCheck == num5))
				{
					sheetView2 = spread.Sheets[num7];
					sheetView2.RecalculateAll();
					string formula2 = sheetView.Cells[i, num2 - 1].Formula;
					if (string.IsNullOrEmpty(formula2))
					{
						continue;
					}
					formula3 = sheetView.Cells[i, num3 - 1].Formula;
					text = sheetView.Cells[i, num4 - 1].Text;
					HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[num7].Tables[num5];
					int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
					for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
					{
						try
						{
							string formula4 = FormulaToTemplate(formula2, num6);
							string formula5 = SetFormulaNextRow(formula4, dataLocationTemplateRowStart + j);
							spread.Sheets[index].Cells[i, column].Formula = formula5;
							string text2 = spread.Sheets[index].Cells[i, column].Text;
							string formula6 = FormulaToTemplate(formula3, num6);
							string text3 = SetFormulaNextRow(formula6, dataLocationTemplateRowStart + j);
							array3 = text3.Split(new char[1] { '&' });
							foreach (string cellFomular in array3)
							{
								string[] array4 = GetCellValild(cellFomular).Split(new char[1] { '#' });
								LockUnLockCol(sheetView2, int.Parse(array4[1]), HTKKFile.ColumnLettersToIndex(array4[2]), text2 == "0", text == "1");
							}
						}
						catch (Exception ex)
						{
							HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
						}
					}
					continue;
				}
				sheetView2 = spread.Sheets[num7];
				sheetView2.RecalculateAll();
				formula3 = sheetView.Cells[i, num3 - 1].Formula;
				if (string.IsNullOrEmpty(formula3))
				{
					break;
				}
				text = sheetView.Cells[i, num4 - 1].Text;
				object value = sheetView.Cells[i, num2 - 1].Value;
				array3 = formula3.Split(new char[1] { '&' });
				foreach (string cellFomular in array3)
				{
					string[] array4 = GetCellValild(cellFomular).Split(new char[1] { '#' });
					if (iRowCheck < 0 || iRowCheck == int.Parse(array4[1]))
					{
						LockUnLockCol(sheetView2, int.Parse(array4[1]), HTKKFile.ColumnLettersToIndex(array4[2]), Convert.ToInt32(value) == 0, text == "1");
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertTableF7(SheetView sView, HTKKCMCSheetInfo sInfo, int iCurTable)
	{
		InsertTable(sView, iCurTable, sInfo);
		sView.AutoCalculation = true;
		sView.RecalculateAll();
		fpSpReindexingSheetTemplate(sView, sInfo);
	}

	public static void InsertTable(SheetView sView, int itable, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			string tableID = sInfo.Tables[itable].TableID;
			tableID = Convert.ToString(int.Parse(tableID) + 1);
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[itable];
			sView.Rows.Add(hTKKCMCTableInfo.Location.TableTemplateRowStart + hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count - 1, hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count);
			sView.CopyRange(hTKKCMCTableInfo.Location.TableTemplateRowStart - 1, -1, hTKKCMCTableInfo.Location.TableTemplateRowStart + hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count - 1, -1, hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count, -1, dataOnly: false);
			HTKKCMCTableInfo hTKKCMCTableInfo2 = new HTKKCMCTableInfo(hTKKCMCTableInfo, 1);
			sInfo.Tables.Insert(itable + 1, hTKKCMCTableInfo2);
			if (string.IsNullOrEmpty(sInfo.Tables[itable + 1].TableID))
			{
				sInfo.Tables[itable + 1].TableID = tableID;
			}
			if (sInfo.Tables[itable + 1].CellTableSTT != null)
			{
				sInfo.Tables[itable + 1].CellTableSTT.Value = tableID;
			}
			if (sInfo.Tables[itable + 1].CellTableReportSTT != null)
			{
				sInfo.Tables[itable + 1].CellTableReportSTT.Value = tableID;
			}
			hTKKCMCTableInfo2.Increment(hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count);
			if (sInfo.Tables[itable + 1].RowsInfo.RowDefaultDisplay == "0")
			{
				sView.Rows[sInfo.Tables[itable + 1].Location.DataLocationTemplateRowStart - 1].Visible = false;
				sInfo.Tables[itable + 1].RowsValues[0].RowDisplay = "0";
			}
			if (hTKKCMCTableInfo.RowsValues.Count > 1)
			{
				sView.Rows.Remove(sInfo.Tables[itable + 1].Location.DataLocationTemplateRowStart, hTKKCMCTableInfo.RowsValues.Count - 1);
			}
			string tableName = sInfo.Tables[itable + 1].TableName;
			for (int i = itable + 2; i < sInfo.Tables.Count; i++)
			{
				string tableName2 = sInfo.Tables[i].TableName;
				if (sInfo.Tables[i].CellTableSTT != null && tableName2 == tableName)
				{
					string tableID2 = sInfo.Tables[i - 1].TableID;
					sInfo.Tables[i].TableID = string.Concat(ConvertHelpers.ConvertToInt(tableID2) + 1);
				}
				if (tableName2 == tableName && sInfo.Tables[i].CellTableSTT != null)
				{
					string tableID2 = sInfo.Tables[i - 1].TableID;
					sInfo.Tables[i].CellTableSTT.Value = string.Concat(ConvertHelpers.ConvertToInt(tableID2) + 1);
				}
				sInfo.Tables[i].Increment(hTKKCMCTableInfo2.TableLength + hTKKCMCTableInfo2.RowsValues.Count);
			}
			sInfo.IncrementStaticCell(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, hTKKCMCTableInfo2.TableLength + hTKKCMCTableInfo2.RowsValues.Count);
			if (sInfo.Tables[itable + 1].Header != null)
			{
				for (int j = 0; j < sInfo.Tables[itable + 1].Header.StaticCells.Count; j++)
				{
					if (sInfo.Tables[itable + 1].Header.StaticCells[j].DefaultValue == "NA")
					{
						sInfo.Tables[itable + 1].Header.StaticCells[j].Value = (ConvertHelpers.ConvertToDouble(sInfo.Tables[itable + 1].Header.StaticCells[0].Value) + 1.0).ToString();
						continue;
					}
					sInfo.Tables[itable + 1].Header.StaticCells[j].Value = sInfo.Tables[itable + 1].Header.StaticCells[j].DefaultValue;
					fpSpBindDataCell(sView, sInfo.Tables[itable + 1].Header.StaticCells[j].CellTemplate.RowIndex, sInfo.Tables[itable + 1].Header.StaticCells[j].CellTemplate.ColIndex, sInfo.Tables[itable + 1].Header.StaticCells[j].Controltype, sInfo.Tables[itable + 1].Header.StaticCells[j].DefaultValue ?? "");
				}
			}
			for (int i = 0; i < sInfo.Tables[itable + 1].RowsInfo.Cells.Count; i++)
			{
				fpSpBindDataCell(sView, sInfo.Tables[itable + 1].Location.TableTemplateRowStart + 1, sInfo.Tables[itable + 1].RowsInfo.Cells[i].CellTemplate.ColIndex, sInfo.Tables[itable + 1].RowsInfo.Cells[i].Controltype, sInfo.Tables[itable + 1].RowsInfo.Cells[i].DefaultValue);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertTableNoSView(int itable, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			string tableID = sInfo.Tables[itable].TableID;
			tableID = Convert.ToString(int.Parse(tableID) + 1);
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[itable];
			HTKKCMCTableInfo hTKKCMCTableInfo2 = new HTKKCMCTableInfo(hTKKCMCTableInfo, 1);
			sInfo.Tables.Insert(itable + 1, hTKKCMCTableInfo2);
			if (string.IsNullOrEmpty(sInfo.Tables[itable + 1].TableID))
			{
				sInfo.Tables[itable + 1].TableID = tableID;
			}
			if (sInfo.Tables[itable + 1].CellTableSTT != null)
			{
				sInfo.Tables[itable + 1].CellTableSTT.Value = tableID;
			}
			if (sInfo.Tables[itable + 1].CellTableReportSTT != null)
			{
				sInfo.Tables[itable + 1].CellTableReportSTT.Value = tableID;
			}
			hTKKCMCTableInfo2.Increment(hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count);
			if (sInfo.Tables[itable + 1].RowsInfo.RowDefaultDisplay == "0")
			{
				sInfo.Tables[itable + 1].RowsValues[0].RowDisplay = "0";
			}
			string tableName = sInfo.Tables[itable + 1].TableName;
			for (int i = itable + 2; i < sInfo.Tables.Count; i++)
			{
				string tableName2 = sInfo.Tables[i].TableName;
				if (sInfo.Tables[i].CellTableSTT != null && tableName2 == tableName)
				{
					string tableID2 = sInfo.Tables[i - 1].TableID;
					sInfo.Tables[i].TableID = string.Concat(ConvertHelpers.ConvertToInt(tableID2) + 1);
				}
				if (tableName2 == tableName && sInfo.Tables[i].CellTableSTT != null)
				{
					string tableID2 = sInfo.Tables[i - 1].TableID;
					sInfo.Tables[i].CellTableSTT.Value = string.Concat(ConvertHelpers.ConvertToInt(tableID2) + 1);
				}
				sInfo.Tables[i].Increment(hTKKCMCTableInfo2.TableLength + hTKKCMCTableInfo2.RowsValues.Count);
			}
			sInfo.IncrementStaticCell(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, hTKKCMCTableInfo2.TableLength + hTKKCMCTableInfo2.RowsValues.Count);
			if (sInfo.Tables[itable + 1].Header != null)
			{
				for (int j = 0; j < sInfo.Tables[itable + 1].Header.StaticCells.Count; j++)
				{
					if (sInfo.Tables[itable + 1].Header.StaticCells[j].DefaultValue == "NA")
					{
						sInfo.Tables[itable + 1].Header.StaticCells[j].Value = (ConvertHelpers.ConvertToDouble(sInfo.Tables[itable + 1].Header.StaticCells[0].Value) + 1.0).ToString();
					}
					else
					{
						sInfo.Tables[itable + 1].Header.StaticCells[j].Value = sInfo.Tables[itable + 1].Header.StaticCells[j].DefaultValue;
					}
				}
			}
			for (int i = 0; i < sInfo.Tables[itable + 1].RowsInfo.Cells.Count; i++)
			{
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpReindexingSheetTemplate(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			sView.FpSpread.SuspendLayout();
			if (sInfo != null)
			{
				for (int i = 0; i < sInfo.Tables.Count; i++)
				{
					if (sInfo.Tables[i].CellTableSTT != null)
					{
						sView.Cells[sInfo.Tables[i].CellTableSTT.RowIndex - 1, sInfo.Tables[i].CellTableSTT.ColIndex - 1].Value = sInfo.Tables[i].TableID;
						sInfo.Tables[i].CellTableSTT.Value = sInfo.Tables[i].TableID;
						fpSpReindexingTableTemplate(sView, sInfo.Tables[i]);
					}
				}
			}
			sView.FpSpread.ResumeLayout();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpReindexingSheetTemplate2(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				fpSpReindexingTableTemplate(sView, sInfo.Tables[i]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool DeleteTableF8(SheetView sView, HTKKCMCSheetInfo sInfo, int iCurTable, int iRow, int iCol, bool batBuocRemove = false)
	{
		try
		{
			if (iCurTable == 0 && !batBuocRemove)
			{
				if (sInfo.Tables.Count((HTKKCMCTableInfo p) => p.TableName == sInfo.Tables[iCurTable].TableName) == 1)
				{
					return false;
				}
			}
			else if (iCurTable > 0 && !batBuocRemove && sInfo.Tables[iCurTable].TableName != sInfo.Tables[iCurTable - 1].TableName)
			{
				return false;
			}
			sView.SetActiveCell(0, iCol);
			if (iCurTable == sInfo.Tables.Count - 1 && iCurTable > 0)
			{
				if (sInfo.Tables[iCurTable - 1].RowsValues.Count > 0)
				{
					iRow = sInfo.Tables[iCurTable - 1].Location.TableTemplateRowStart - 1;
				}
				else
				{
					int num = iRow + 1 - sInfo.Tables[iCurTable].Location.TableTemplateRowStart;
					iRow = sInfo.Tables[iCurTable - 1].Location.TableTemplateRowStart - 1 + num;
				}
			}
			else
			{
				iRow = sInfo.Tables[iCurTable].Location.TableTemplateRowStart - 1;
			}
			RemoveTable(sView, iCurTable, sInfo);
			fpSpReindexingSheetTemplate(sView, sInfo);
			sView.SetActiveCell(iRow, iCol);
			sView.AutoCalculation = true;
			sView.RecalculateAll();
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool RemoveTableF8New(SheetView sView, HTKKCMCSheetInfo sInfo, int iCurTable, bool setActiveCell)
	{
		try
		{
			string tableName = sInfo.Tables[iCurTable].TableName;
			List<HTKKCMCTableInfo> list = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == tableName).ToList();
			if (list.Count == 1)
			{
				return false;
			}
			if (setActiveCell)
			{
				int index = sView.ActiveRow.Index;
				int index2 = sView.ActiveColumn.Index;
				sView.SetActiveCell(0, index2);
				if (iCurTable == sInfo.Tables.Count - 1 && iCurTable > 0)
				{
					if (sInfo.Tables[iCurTable - 1].RowsValues.Count > 0)
					{
						index = sInfo.Tables[iCurTable - 1].Location.TableTemplateRowStart - 1;
					}
					else
					{
						int num = sView.ActiveRow.Index + 1 - sInfo.Tables[iCurTable].Location.TableTemplateRowStart;
						index = sInfo.Tables[iCurTable - 1].Location.TableTemplateRowStart - 1 + num;
					}
				}
				else
				{
					index = sInfo.Tables[iCurTable].Location.TableTemplateRowStart - 1;
				}
				sView.SetActiveCell(index, index2);
			}
			RemoveTableNew(sView, iCurTable, sInfo);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static void RemoveTableNew(SheetView sView, int itable, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[itable];
			string tableName = hTKKCMCTableInfo.TableName;
			sView.RemoveRows(hTKKCMCTableInfo.Location.TableTemplateRowStart - 1, hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count);
			List<HTKKCMCTableInfo> tables = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == tableName).ToList();
			int num = sInfo.Tables.FindIndex((HTKKCMCTableInfo x) => x.TableName == tableName && x.TableID == tables[0].TableID);
			for (int num2 = itable + 1; num2 < sInfo.Tables.Count; num2++)
			{
				int num3 = ((num2 - 1 == num || (num2 - 1 >= 0 && sInfo.Tables[num2].TableName != sInfo.Tables[num2 - 1].TableName)) ? 1 : (Convert.ToInt32(sInfo.Tables[num2 - 2].TableID) + 1));
				sInfo.Tables[num2].TableID = num3.ToString();
				if (sInfo.Tables[num2].Header.StaticCells.Count > 0 && ConvertHelpers.ConvertToString(sInfo.Tables[num2].Header.StaticCells[0].DefaultValue).Equals("NA"))
				{
					sInfo.Tables[num2].Header.StaticCells[0].Value = num3.ToString();
				}
				if (sInfo.Tables[num2].CellTableSTT != null)
				{
					sInfo.Tables[num2].CellTableSTT.Value = num3.ToString();
				}
				sInfo.Tables[num2].Increment((hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count) * -1);
			}
			sInfo.IncrementStaticCell(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, (hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count) * -1);
			sInfo.Tables.RemoveAt(itable);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void RemoveTable(SheetView sView, int itable, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[itable];
			sView.Rows.Remove(hTKKCMCTableInfo.Location.TableTemplateRowStart - 1, hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count);
			for (int i = itable + 1; i < sInfo.Tables.Count; i++)
			{
				if (sInfo.Tables[i].TableID != "1")
				{
					sInfo.Tables[i].TableID = i.ToString();
				}
				if (sInfo.Tables[i].Header.StaticCells.Count > 0 && ConvertHelpers.ConvertToString(sInfo.Tables[i].Header.StaticCells[0].DefaultValue).Equals("NA"))
				{
					sInfo.Tables[i].Header.StaticCells[0].Value = sInfo.Tables[i].TableID;
				}
				if (sInfo.Tables[i].CellTableSTT != null)
				{
					sInfo.Tables[i].CellTableSTT.Value = i.ToString();
				}
				sInfo.Tables[i].Increment((hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count) * -1);
			}
			sInfo.IncrementStaticCell(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, (hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count) * -1);
			sInfo.Tables.RemoveAt(itable);
			ReindexTableId(sInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertNumberCell(SheetView sView, int iRow, int iCol, string value, HTKKCMCValidate validate, string defaultValue)
	{
		try
		{
			int num = 0;
			if (iRow == 23 && iCol == HTKKFile.ColumnLettersToIndex("D"))
			{
				num = iRow;
			}
			if ((!HTKKFormater.isNullString(validate.MaxValue) && validate.MaxValue.ToNumber() > 999999999999999m) || (!HTKKFormater.isNullString(validate.MinValue) && validate.MinValue.ToNumber() < -999999999999999m))
			{
				InsertCurrencyCell(sView, iRow, iCol, value, validate, defaultValue);
				return;
			}
			int decimalPlaces = 0;
			NumberCellType numberCellType = new NumberCellType();
			numberCellType.ShowSeparator = true;
			numberCellType.NegativeFormat = NegativeFormat.Parentheses;
			if (HTKKFormater.isNullString(defaultValue) || defaultValue == "0")
			{
				decimalPlaces = 0;
				numberCellType.DecimalSeparator = ",";
				numberCellType.Separator = ".";
			}
			else if (defaultValue.IndexOf(".") > -1)
			{
				decimalPlaces = defaultValue.Substring(defaultValue.LastIndexOf(".") + 1).Length;
				numberCellType.DecimalSeparator = ".";
				numberCellType.Separator = ",";
			}
			else if (defaultValue.IndexOf(",") > -1)
			{
				decimalPlaces = defaultValue.Substring(defaultValue.LastIndexOf(",") + 1).Length;
				numberCellType.DecimalSeparator = ",";
				numberCellType.Separator = ".";
			}
			numberCellType.MinimumValue = getMinNumber(validate.MinValue);
			numberCellType.MaximumValue = getMaxNumber(validate.MaxValue);
			numberCellType.DecimalPlaces = decimalPlaces;
			sView.Cells[iRow - 1, iCol - 1].CellType = numberCellType;
			if (!HTKKFormater.isNullString(sView.Cells[iRow - 1, iCol - 1].Formula))
			{
				if (!HTKKFormater.isNullString(value))
				{
					sView.Cells[iRow - 1, iCol - 1].Value = ConvertHelpers.ConvertToDouble1(value).ToString();
				}
				else if (!HTKKFormater.isNullString(defaultValue))
				{
					sView.Cells[iRow - 1, iCol - 1].Value = ConvertHelpers.ConvertToDouble1(defaultValue).ToString();
				}
				else
				{
					sView.Cells[iRow - 1, iCol - 1].Value = "0";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertCurrencyCell(SheetView sView, int iRow, int iCol, string value, HTKKCMCValidate validate, string defaultValue)
	{
		try
		{
			int num = 0;
			if (iRow == 23 && iCol == HTKKFile.ColumnLettersToIndex("D"))
			{
				num = iRow;
			}
			int decimalPlaces = 0;
			CurrencyCellType currencyCellType = new CurrencyCellType();
			currencyCellType.ShowSeparator = true;
			currencyCellType.ShowCurrencySymbol = false;
			if (HTKKFormater.isNullString(defaultValue) || defaultValue == "0")
			{
				decimalPlaces = 0;
				currencyCellType.DecimalSeparator = ",";
				currencyCellType.Separator = ".";
			}
			else if (defaultValue.IndexOf(".") > -1)
			{
				decimalPlaces = defaultValue.Substring(defaultValue.LastIndexOf(".") + 1).Length;
				currencyCellType.DecimalSeparator = ".";
				currencyCellType.Separator = ",";
			}
			else if (defaultValue.IndexOf(",") > -1)
			{
				decimalPlaces = defaultValue.Substring(defaultValue.LastIndexOf(",") + 1).Length;
				currencyCellType.DecimalSeparator = ",";
				currencyCellType.Separator = ".";
			}
			currencyCellType.MinimumValue = getMinNumber(validate.MinValue);
			currencyCellType.MaximumValue = getMaxNumber(validate.MaxValue);
			currencyCellType.DecimalPlaces = decimalPlaces;
			sView.Cells[iRow - 1, iCol - 1].CellType = currencyCellType;
			if (!HTKKFormater.isNullString(sView.Cells[iRow - 1, iCol - 1].Formula))
			{
				if (!HTKKFormater.isNullString(value))
				{
					sView.Cells[iRow - 1, iCol - 1].Value = ConvertHelpers.ConvertToDouble1(value).ToString();
				}
				else if (!HTKKFormater.isNullString(defaultValue))
				{
					sView.Cells[iRow - 1, iCol - 1].Value = ConvertHelpers.ConvertToDouble1(defaultValue).ToString();
				}
				else
				{
					sView.Cells[iRow - 1, iCol - 1].Value = "0";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool IsNumeric(string input)
	{
		if (!decimal.TryParse(input, out var _))
		{
			return false;
		}
		return true;
	}

	public static void SetRowHeight(SheetView sView, int row1)
	{
		Row row2 = sView.Cells[row1 - 1, 0].Row;
		float preferredHeight = row2.GetPreferredHeight();
		if (row2.Height < preferredHeight - 5f || row2.Height > preferredHeight + 5f)
		{
			row2.Height = preferredHeight;
		}
	}

	public static void SetRowHeight(SheetView sView, int row1, float height)
	{
		Row row2 = sView.Cells[row1 - 1, 0].Row;
		row2.Height = height;
	}

	public static Dictionary<string, string> GetDMFromCombobox(SheetView sView, int row, int col)
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		try
		{
			ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
			ICellType cellType = sView.Cells[row - 1, col - 1].CellType;
			if (cellType == null)
			{
				return dictionary;
			}
			Type type = cellType.GetType();
			if (cellType != null && type.Name.ToLower().IndexOf("combobox") >= 0)
			{
				ComboBoxCellType comboBoxCellType2 = (ComboBoxCellType)cellType;
				if (comboBoxCellType2.Items.Count() == 0)
				{
					return dictionary;
				}
				int num = comboBoxCellType2.ItemData.Count();
				for (int i = 0; i < comboBoxCellType2.Items.Length; i++)
				{
					string text = comboBoxCellType2.Items[i];
					string text2 = ((num == 0) ? text : comboBoxCellType2.ItemData[i]);
					if (!string.IsNullOrEmpty(text2) && !dictionary.ContainsKey(text2))
					{
						dictionary.Add(text2, text);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return dictionary;
	}

	public static void changFormatCell(SheetView sView, string cellFormat, string maLoaiTien)
	{
		try
		{
			string text = "###";
			int num = 0;
			string[] array = cellFormat.Split(new string[1] { text }, StringSplitOptions.None);
			num = (maLoaiTien.Equals("USD") ? 2 : 0);
			if (array.Length <= 0)
			{
				return;
			}
			for (int i = 0; i < array.Length; i++)
			{
				int[] array2 = ArrayCellId(array[i]);
				if (array2.Length == 2)
				{
					ICellType cellType = sView.Cells[array2[1], array2[0]].CellType;
					if (cellType.ToString() == "NumberCellType")
					{
						NumberCellType numberCellType = (NumberCellType)sView.Cells[array2[1], array2[0]].CellType;
						numberCellType.DecimalPlaces = num;
						sView.Cells[array2[1], array2[0]].CellType = numberCellType;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static List<string> SplitString(string text, string splitBy)
	{
		string[] separator = new string[1] { splitBy };
		string[] source = text.Split(separator, StringSplitOptions.RemoveEmptyEntries);
		return source.ToList();
	}

	public static List<string> SplitString2(string text, string splitBy)
	{
		string[] separator = new string[1] { splitBy };
		string[] source = text.Split(separator, StringSplitOptions.None);
		return source.ToList();
	}

	public static List<HTKKCMCRowValue> GetRowValueInTable(HTKKCMCTableInfo table, List<string> paths, List<string> values)
	{
		try
		{
			if (table == null)
			{
				return null;
			}
			List<int> indexs = new List<int>();
			foreach (string path in paths)
			{
				List<HTKKCMCTableCellInfo> cells = table.RowsInfo.Cells;
				Predicate<HTKKCMCTableCellInfo> match = (HTKKCMCTableCellInfo x) => x.Path == path;
				int num = cells.FindIndex(match);
				if (num < 0)
				{
					break;
				}
				indexs.Add(num);
			}
			if (indexs.Count == 0)
			{
				return null;
			}
			for (int num2 = 0; num2 < values.Count; num2++)
			{
				int num = indexs[num2];
				if (num >= 0)
				{
					string controltype = table.RowsInfo.Cells[num].Controltype;
					if (controltype == "14" && values[num2].Contains("/"))
					{
						values[num2] = values[num2].Replace("/", "").Trim();
					}
				}
			}
			List<HTKKCMCRowValue> list = null;
			List<HTKKCMCRowValue> source = table.RowsValues.ToList();
			List<HTKKCMCRowValue> list2 = source.Where(delegate(HTKKCMCRowValue xx)
			{
				int num3 = 0;
				for (int i = 0; i < indexs.Count; i++)
				{
					if (xx.RowValue[indexs[i]].Trim().Equals(values[i].Trim()))
					{
						num3++;
					}
					if (num3 == indexs.Count)
					{
						return true;
					}
				}
				return false;
			}).ToList();
			list = list2;
			for (int num2 = 0; num2 < values.Count; num2++)
			{
				int num = indexs[num2];
				string controltype = table.RowsInfo.Cells[num].Controltype;
				if (controltype == "14" && !string.IsNullOrEmpty(values[num2].Trim()))
				{
					values[num2] = FormartDateStr(values[num2].Trim());
				}
			}
			if (list.Count > 0)
			{
				return list;
			}
			return null;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static List<HTKKCMCTableInfo> GetTableByKeyBody(List<HTKKCMCTableInfo> tables, List<string> paths, List<string> values, string tablePath)
	{
		try
		{
			List<HTKKCMCTableInfo> list = new List<HTKKCMCTableInfo>();
			if (tables == null)
			{
				return null;
			}
			int num = 0;
			List<HTKKCMCTableInfo> list2 = tables.Where((HTKKCMCTableInfo x) => x.TablePath == tablePath).ToList();
			if (list2 != null && list2.Count() == 1 && list2[0].Header != null && list2[0].Header.Length == 0)
			{
				return list2.ToList();
			}
			foreach (HTKKCMCTableInfo table in tables)
			{
				if (!string.IsNullOrEmpty(tablePath) && table.TablePath != tablePath)
				{
					num++;
					continue;
				}
				int num2 = 0;
				int i;
				for (i = 0; i < paths.Count; i++)
				{
					if (!string.IsNullOrEmpty(values[i]))
					{
						int num3 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == paths[i]);
						if (num3 < 0)
						{
							continue;
						}
						string controltype = table.RowsInfo.Cells[num3].Controltype;
						for (int num4 = 0; num4 < table.RowsValues.Count; num4++)
						{
							if (table.RowsValues[num4].RowValue[num3].Trim() == ((controltype != "14") ? values[i].Trim() : values[i].Replace("/", "").Trim()))
							{
								num2++;
								break;
							}
						}
					}
					else
					{
						num2++;
					}
				}
				if (num2 > 0 && num2 == paths.Count && num >= 0 && num < tables.Count)
				{
					list.Add(tables[num]);
				}
				num++;
			}
			if (list.Count > 0)
			{
				return list;
			}
			return null;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static HTKKCMCTableInfo GetTableByHeader(List<HTKKCMCTableInfo> tables, List<string> paths, List<string> values)
	{
		try
		{
			if (tables == null)
			{
				return null;
			}
			int num = 0;
			foreach (HTKKCMCTableInfo table in tables)
			{
				int num2 = 0;
				int i;
				for (i = 0; i < paths.Count; i++)
				{
					if (!string.IsNullOrEmpty(paths[i]))
					{
						int num3 = table.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == paths[i] && x.Value == values[i]);
						if (num3 >= 0)
						{
							num2++;
						}
					}
					else
					{
						num2++;
					}
				}
				if (num2 > 0 && num2 == paths.Count)
				{
					break;
				}
				num++;
			}
			if (num >= 0 && num < tables.Count)
			{
				return tables[num];
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return null;
	}

	public static string FormartDateStr(string dateStr)
	{
		if (!string.IsNullOrEmpty(dateStr) && dateStr.Length == 8 && dateStr.IndexOf("/") < 0)
		{
			return dateStr.Substring(0, 2) + "/" + dateStr.Substring(2, 2) + "/" + dateStr.Substring(4, 4);
		}
		return dateStr;
	}

	public static string ListToString(List<string> list, string formatKey)
	{
		return string.Join(formatKey, list.ToArray());
	}

	public static string[] StringToList(string text, string fomartKey)
	{
		string[] separator = new string[1] { fomartKey };
		return text.Split(separator, StringSplitOptions.RemoveEmptyEntries);
	}

	public static string FormatKey(string str1, string str2)
	{
		return str1 + "___" + str2;
	}

	public static string FormatKey2(string str1, string str2, string fomartKey)
	{
		if (!string.IsNullOrEmpty(str1) && !string.IsNullOrEmpty(str2))
		{
			return str1 + fomartKey + str2;
		}
		if (string.IsNullOrEmpty(str1) && !string.IsNullOrEmpty(str2))
		{
			return str2;
		}
		if (!string.IsNullOrEmpty(str1) && string.IsNullOrEmpty(str2))
		{
			return str1;
		}
		return "";
	}

	public static string FormatKey3(string str1, string str2, string fomartKey)
	{
		return str1 + fomartKey + str2;
	}

	public static string GetMaTieuMucByMaHH(string maHH, string loaiDM)
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0047: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Expected O, but got Unknown
		//IL_0094: Unknown result type (might be due to invalid IL or missing references)
		//IL_009b: Expected O, but got Unknown
		try
		{
			string result = "";
			string text = HTKKConfig.INTERFACE_INI + "\\DM_MAP_MAHH_MA_TM.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Items");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Type"]).Value;
				if (!(value == loaiDM))
				{
					continue;
				}
				foreach (XmlNode childNode in val2.ChildNodes)
				{
					XmlNode val3 = childNode;
					string value2 = ((XmlNode)val3.Attributes["MaHH"]).Value;
					string value3 = ((XmlNode)val3.Attributes["MaTM"]).Value;
					if (value2 == maHH)
					{
						result = value3;
						break;
					}
				}
				break;
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static string GetTenTieuMucTuMa(string maTieuMuc, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			if (string.IsNullOrEmpty(maTieuMuc))
			{
				return "";
			}
			DateTime ngayDauKyKeKhai = GetNgayDauKyKeKhai(TTinChung, PeriodInfo);
			return GetTenTmByMaTm(maTieuMuc, ngayDauKyKeKhai);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static string GetTenTmByMaTm(string maTieuMuc, DateTime dateKyKK)
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_004e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0055: Expected O, but got Unknown
		try
		{
			string result = "";
			string text = HTKKConfig.INTERFACE_INI + "\\DM_MaTM.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			List<string[]> list = new List<string[]>();
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["MaTM"]).Value;
				string value2 = ((XmlNode)val2.Attributes["TenTM"]).Value;
				string value3 = ((XmlNode)val2.Attributes["HLTu"]).Value;
				DateTime dateTime = ConvertHelpers.ConvertToDateTime(value3, "dd/MM/yyyy");
				string value4 = ((XmlNode)val2.Attributes["HLDen"]).Value;
				DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(value4, "dd/MM/yyyy");
				if (maTieuMuc == value && dateTime <= dateKyKK && dateTime2 >= dateKyKK)
				{
					list.Add(new string[2] { value2, value3 });
				}
			}
			if (list != null)
			{
				result = list[0][0];
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static string[] GetMucByMaTm(string maTieuMuc, DateTime dateKyKK)
	{
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_001f: Expected O, but got Unknown
		//IL_004f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0056: Expected O, but got Unknown
		try
		{
			string[] result = new string[2];
			string text = HTKKConfig.INTERFACE_INI + "\\DM_MaTM.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			List<string[]> list = new List<string[]>();
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["MaTM"]).Value;
				string value2 = ((XmlNode)val2.Attributes["MaMuc"]).Value;
				string value3 = ((XmlNode)val2.Attributes["TenMuc"]).Value;
				string value4 = ((XmlNode)val2.Attributes["HLTu"]).Value;
				DateTime dateTime = ConvertHelpers.ConvertToDateTime(value4, "dd/MM/yyyy");
				string value5 = ((XmlNode)val2.Attributes["HLDen"]).Value;
				DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(value5, "dd/MM/yyyy");
				if (maTieuMuc == value && dateTime <= dateKyKK && dateTime2 >= dateKyKK)
				{
					list.Add(new string[2] { value2, value3 });
				}
			}
			if (list != null)
			{
				result = list[0];
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static double ConvertByDefault(object valueObj, string defaultValue)
	{
		double num = 0.0;
		try
		{
			string text = valueObj.ToString();
			int num2 = 0;
			if (text.ToString().IndexOf(".") >= 0)
			{
				string text2 = text.Substring(text.LastIndexOf(".") + 1, text.Length - (text.LastIndexOf(".") + 1));
				num2 = text2.Length;
			}
			if (num2 == 0 && defaultValue.IndexOf(",") >= 0)
			{
				string text2 = defaultValue.Substring(defaultValue.LastIndexOf(",") + 1, defaultValue.Length - (defaultValue.LastIndexOf(",") + 1));
				num2 = text2.Length;
			}
			num = ((text.IndexOf(")") >= 0) ? (-1.0 * Convert.ToDouble(text.Replace(",", "").Replace(".", "").Replace("(", "")
				.Replace(")", ""))) : Convert.ToDouble(text.Replace(",", "").Replace(".", "")));
			return num2 switch
			{
				1 => Math.Round(num / 10.0, num2), 
				2 => Math.Round(num / 100.0, num2), 
				3 => Math.Round(num / 1000.0, num2), 
				4 => Math.Round(num / 10000.0, num2), 
				5 => Math.Round(num / 100000.0, num2), 
				6 => Math.Round(num / 1000000.0, num2), 
				7 => Math.Round(num / 10000000.0, num2), 
				8 => Math.Round(num / 100000000.0, num2), 
				9 => Math.Round(num / 1000000000.0, num2), 
				10 => Math.Round(num / 10000000000.0, num2), 
				11 => Math.Round(num / 100000000000.0, num2), 
				12 => Math.Round(num / 1000000000000.0, num2), 
				13 => Math.Round(num / 10000000000000.0, num2), 
				14 => Math.Round(num / 100000000000000.0, num2), 
				15 => Math.Round(num / 1000000000000000.0, num2), 
				_ => num, 
			};
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return num;
		}
	}

	public static DateTime GetNgayDauKyKeKhai(HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo TTPeriodInfo)
	{
		string s = "";
		if (TTinChung.kieuKy == "Y")
		{
			s = ((!string.IsNullOrEmpty(TTinChung.kyKKhaiTuThang)) ? ("01/" + TTinChung.kyKKhaiTuThang) : (string.IsNullOrEmpty(TTinChung.kyKKhaiTuNgay) ? ("01/01/" + TTinChung.nam) : TTinChung.kyKKhaiTuNgay));
		}
		if (TTinChung.kieuKy == "M" || TTinChung.kieuKy == "Q")
		{
			return NgayDauKyTinhThue(TTPeriodInfo);
		}
		if (TTinChung.kieuKy == "D" && !string.IsNullOrEmpty(TTinChung.ngay) && !string.IsNullOrEmpty(TTinChung.thang) && !string.IsNullOrEmpty(TTinChung.nam))
		{
			s = TTinChung.ngay + "/" + TTinChung.thang + "/" + TTinChung.nam;
		}
		if (!string.IsNullOrEmpty(TTinChung.kyKKhaiTuNgay) && (TTinChung.kieuKy == "D_D" || TTinChung.kieuKy == "K"))
		{
			s = TTinChung.kyKKhaiTuNgay;
		}
		return DateTime.ParseExact(s, "dd/MM/yyyy", null);
	}

	public static DateTime GetNgayCuoiKyKeKhai(HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo TTPeriodInfo)
	{
		if (TTinChung.kieuKy == "Y")
		{
			string text = "";
			text = ((!string.IsNullOrEmpty(TTinChung.kyKKhaiDenThang)) ? ("01/" + TTinChung.kyKKhaiTuThang) : (string.IsNullOrEmpty(TTinChung.kyKKhaiDenNgay) ? ("01/01/" + TTinChung.nam) : TTinChung.kyKKhaiTuNgay));
			return new DateTime(DateTime.ParseExact(text, "dd/MM/yyyy", null).Year + 1, 1, 1).AddDays(-1.0);
		}
		if (TTinChung.kieuKy == "M")
		{
			return new DateTime(Convert.ToInt32(TTPeriodInfo.Nam), Convert.ToInt32(TTPeriodInfo.Thang), DateTime.DaysInMonth(Convert.ToInt32(TTPeriodInfo.Nam), Convert.ToInt32(TTPeriodInfo.Thang)));
		}
		if (TTinChung.kieuKy == "Q")
		{
			return GetNgayCuoiQuy(Convert.ToInt32(TTPeriodInfo.Quy), Convert.ToInt32(TTPeriodInfo.Nam), 1, 1);
		}
		if (TTinChung.kieuKy == "D" && !string.IsNullOrEmpty(TTinChung.kyKKhaiDenNgay))
		{
			string text = TTinChung.kyKKhaiDenNgay;
			return DateTime.ParseExact(text, "dd/MM/yyyy", null);
		}
		if ((!string.IsNullOrEmpty(TTinChung.kyKKhaiDenNgay) && TTinChung.kieuKy == "D_D") || TTinChung.kieuKy == "K")
		{
			string text = TTinChung.kyKKhaiDenNgay;
			return DateTime.ParseExact(text, "dd/MM/yyyy", null);
		}
		return DateTime.Now;
	}

	public static string ReplaceFromStart(string text, string txtStart)
	{
		if (string.IsNullOrEmpty(text) || text.Length - txtStart.Length <= 0)
		{
			return text;
		}
		return text.Substring(txtStart.Length, text.Length - txtStart.Length);
	}

	public static string getThuSuatByMa(string ma, string controlType)
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0047: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Expected O, but got Unknown
		//IL_0090: Unknown result type (might be due to invalid IL or missing references)
		//IL_0097: Expected O, but got Unknown
		try
		{
			string result = "";
			string text = HTKKConfig.INTERFACE_INI + "\\DM_DUNG_CHUNG.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Map");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				if (!(((XmlNode)val2.Attributes["Controltype"]).Value == controlType))
				{
					continue;
				}
				foreach (XmlNode childNode in val2.ChildNodes)
				{
					XmlNode val3 = childNode;
					string value = ((XmlNode)val3.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					string text2 = array[0];
					if (ma == text2)
					{
						result = array[1];
					}
				}
				break;
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static HTKKCMCTTinChungInfo GetHeaderInfo(string mst, bool removeMaCQT = true)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		HTKKCMCTTinChungInfo hTKKCMCTTinChungInfo = new HTKKCMCTTinChungInfo();
		try
		{
			XmlDocument val = new XmlDocument();
			val.Load("../DataFiles/" + mst + "/HEADER.xml");
			hTKKCMCTTinChungInfo.tenCQTCapCuc = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p28']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.maCQTCapCuc = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p32']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.tenCQTNoiNop = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p29']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.maCQTNoiNop = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p33']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.tenTinhNNT = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p18']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.tenHuyenNNT = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p17']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.maTinhNNT = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p50']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.maHuyenNNT = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p51']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.tenTinhDLyThue = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p41']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.tenHuyenDLyThue = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p40']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.maTinhDLyThue = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p52']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.maHuyenDLyThue = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p53']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.tenXaPhuongNNT = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p54']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.maXaPhuongNNT = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p56']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.tenXaPhuongDLyThue = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p55']").Attributes["Value"]).InnerText;
			hTKKCMCTTinChungInfo.maXaPhuongDLyThue = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p57']").Attributes["Value"]).InnerText;
			if (removeMaCQT)
			{
				List<string> list = SplitString(hTKKCMCTTinChungInfo.tenCQTCapCuc, " - ");
				hTKKCMCTTinChungInfo.tenCQTCapCuc = ((list.Count <= 2) ? list[list.Count - 1] : GetSubstringFromArray(list, 1, " - "));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return hTKKCMCTTinChungInfo;
	}

	public static string GetSubstringFromArray(List<string> array, int from, string key)
	{
		string text = "";
		for (int i = from; i < array.Count; i++)
		{
			text = text + array[i] + key;
		}
		return text.Substring(0, text.Length - key.Length);
	}

	public static string GetMaCQTNoiNopHeader(string mst)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		XmlDocument val = new XmlDocument();
		val.Load(HTKKConfig.DATAFILES + "\\" + mst + "/HEADER.xml");
		return ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p33']").Attributes["Value"]).InnerText;
	}

	public static bool CheckCellHaveNotification(SheetView sView, int row, int col)
	{
		try
		{
			if (sView.Cells[row - 1, col - 1].NoteStyle != NoteStyle.Hidden)
			{
				return true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return false;
	}

	public static void ThayDoiTinhTheoCQT(string maCQTCapCuc_TSC, string maCQTNoiNop_TSC, SheetView sView, HTKKPeriodInfo PeriodInfo, int[] cellTenTinh, int[] cellMaTinh, int[] cellTenHuyen, int[] cellMaHuyen, int[] cellTenXa, int[] cellMaXa, int[] cellDiaChi, int[] cellTenDonVi, int[] cellMaSoDonVi, bool PhuThuocTruSoChinh = true)
	{
		try
		{
			if (PeriodInfo.maCQTCapCuc == "82500")
			{
				sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Locked = false;
				sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked = false;
			}
			else
			{
				if (!string.IsNullOrEmpty(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text.Trim()) && !string.IsNullOrEmpty(sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Text.Trim()) && cellTenXa != null && !string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim()))
				{
					return;
				}
				if (!string.IsNullOrEmpty(PeriodInfo.MaTinh))
				{
					sView.Cells[cellTenTinh[1], cellTenTinh[0]].Value = PeriodInfo.TenTinh;
					sView.Cells[cellMaTinh[1], cellMaTinh[0]].Value = PeriodInfo.MaTinh;
					sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked = true;
					Dictionary<string, string> dictionary = HTKKWinDictionaryDiaDanh.dicQuanHuyen.Where((KeyValuePair<string, string> x) => x.Key.StartsWith(PeriodInfo.MaTinh)).ToDictionary((KeyValuePair<string, string> x) => x.Key, (KeyValuePair<string, string> x) => x.Value);
					if (dictionary != null)
					{
						if (string.IsNullOrEmpty(sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Text.Trim()))
						{
							sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Value = "";
							sView.Cells[cellMaHuyen[1], cellMaHuyen[0]].Value = "";
						}
						InsertCombobox(sView, cellTenHuyen[1] + 1, cellTenHuyen[0] + 1, dictionary);
						if (cellTenXa != null && string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim()))
						{
							sView.Cells[cellTenXa[1], cellTenXa[0]].Value = "";
							sView.Cells[cellMaXa[1], cellMaXa[0]].Value = "";
						}
						if (cellDiaChi != null && cellDiaChi.Length > 0 && string.IsNullOrEmpty(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Text.Trim()))
						{
							sView.Cells[cellDiaChi[1], cellDiaChi[0]].Value = "";
						}
					}
					return;
				}
				string maTinh = (HTKKWinDictionaryDiaDanh.dicTinhTP.ContainsValue(PeriodInfo.maCQTCapCuc.Substring(0, 3)) ? PeriodInfo.maCQTCapCuc.Substring(0, 3) : "");
				if (PhuThuocTruSoChinh && PeriodInfo.maCQTCapCuc == maCQTCapCuc_TSC)
				{
					if (cellTenXa != null)
					{
						sView.Cells[cellTenXa[1], cellTenXa[0]].Locked = true;
					}
					sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Locked = true;
					sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked = true;
					if (cellDiaChi != null)
					{
						sView.Cells[cellDiaChi[1], cellDiaChi[0]].Locked = true;
					}
					if (cellTenDonVi != null)
					{
						sView.Cells[cellTenDonVi[1], cellTenDonVi[0]].Locked = true;
					}
					if (cellMaSoDonVi != null)
					{
						sView.Cells[cellMaSoDonVi[1], cellMaSoDonVi[0]].Locked = true;
					}
				}
				else if (string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy) && !string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc) && !string.IsNullOrEmpty(maTinh))
				{
					sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text = HTKKWinDictionaryDiaDanh.dicTinhTP[maTinh];
					sView.Cells[cellMaTinh[1], cellMaTinh[0]].Text = maTinh;
					sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked = true;
					Dictionary<string, string> dictionary = HTKKWinDictionaryDiaDanh.dicQuanHuyen.Where((KeyValuePair<string, string> x) => x.Key.StartsWith(maTinh)).ToDictionary((KeyValuePair<string, string> x) => x.Key, (KeyValuePair<string, string> x) => x.Value);
					if (dictionary != null)
					{
						if (string.IsNullOrEmpty(sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Text.Trim()))
						{
							sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Value = "";
							sView.Cells[cellMaHuyen[1], cellMaHuyen[0]].Value = "";
						}
						InsertCombobox(sView, cellTenHuyen[1] + 1, cellTenHuyen[0] + 1, dictionary);
						if (cellTenXa != null && string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim()))
						{
							sView.Cells[cellTenXa[1], cellTenXa[0]].Value = "";
							sView.Cells[cellMaXa[1], cellMaXa[0]].Value = "";
						}
						if (cellDiaChi != null && string.IsNullOrEmpty(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Text.Trim()))
						{
							sView.Cells[cellDiaChi[1], cellDiaChi[0]].Value = "";
						}
					}
				}
				else
				{
					if (string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy) || string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc))
					{
						return;
					}
					string maHuyen = (HTKKWinDictionaryDiaDanh.dicQuanHuyen.ContainsKey(PeriodInfo.maCQTQuanLy) ? PeriodInfo.maCQTQuanLy : "");
					if (!string.IsNullOrEmpty(maHuyen) && !string.IsNullOrEmpty(maTinh))
					{
						KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
						sView.Cells[cellTenTinh[1], cellTenTinh[0]].Value = keyValuePair.Key;
						sView.Cells[cellMaTinh[1], cellMaTinh[0]].Value = maTinh;
						sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked = true;
						sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Value = HTKKWinDictionaryDiaDanh.dicQuanHuyen[maHuyen];
						sView.Cells[cellMaHuyen[1], cellMaHuyen[0]].Value = maHuyen;
						sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Locked = true;
						if (cellTenXa != null)
						{
							Dictionary<string, string> dic = HTKKWinDictionaryDiaDanh.dicPhuongXa.Where((KeyValuePair<string, string> x) => x.Key.Substring(0, 5) == maHuyen).ToDictionary((KeyValuePair<string, string> x) => x.Key, (KeyValuePair<string, string> x) => x.Value);
							if (string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim()))
							{
								sView.Cells[cellTenXa[1], cellTenXa[0]].Value = "";
								sView.Cells[cellMaXa[1], cellMaXa[0]].Value = "";
							}
							InsertCombobox(sView, cellTenXa[1] + 1, cellTenXa[0] + 1, dic);
						}
						if (cellDiaChi != null && string.IsNullOrEmpty(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Text.Trim()))
						{
							sView.Cells[cellDiaChi[1], cellDiaChi[0]].Value = "";
						}
					}
					if (string.IsNullOrEmpty(maHuyen) && !string.IsNullOrEmpty(maTinh))
					{
						KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
						sView.Cells[cellTenTinh[1], cellTenTinh[0]].Value = keyValuePair.Key;
						sView.Cells[cellMaTinh[1], cellMaTinh[0]].Value = maTinh;
						sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked = true;
						if (string.IsNullOrEmpty(sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Text.Trim()))
						{
							sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Value = "";
							sView.Cells[cellMaHuyen[1], cellMaHuyen[0]].Value = "";
							sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Locked = false;
						}
						if (cellTenXa != null && string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim()))
						{
							sView.Cells[cellTenXa[1], cellTenXa[0]].Value = "";
							sView.Cells[cellMaXa[1], cellMaXa[0]].Value = "";
						}
						if (cellDiaChi != null && string.IsNullOrEmpty(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Text.Trim()))
						{
							sView.Cells[cellDiaChi[1], cellDiaChi[0]].Value = "";
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ThayDoiTinhTheoCQTDiaBan(string maCQTCapCuc_TSC, string maCQTNoiNop_TSC, SheetView sView, HTKKPeriodInfo PeriodInfo, int[] cellTenTinh, int[] cellMaTinh, int[] cellTenXa, int[] cellMaXa, int[] cellDiaChi, int[] cellTenDonVi, int[] cellMaSoDonVi, bool PhuThuocTruSoChinh = true)
	{
		try
		{
			if (PeriodInfo.maCQTCapCuc == "82500")
			{
				sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked = false;
			}
			else
			{
				if (!string.IsNullOrEmpty(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text.Trim()) && cellTenXa != null && !string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim()))
				{
					return;
				}
				string maTinhCu = (HTKKWinDictionaryDiaDanh.dicTinhTPMoi.ContainsValue(PeriodInfo.maCQTCapCuc.Substring(0, 3)) ? PeriodInfo.maCQTCapCuc.Substring(0, 3) : "");
				string maTinh = GetMaTinhMoiByMaTinhCu(maTinhCu);
				if (PhuThuocTruSoChinh && PeriodInfo.maCQTCapCuc == maCQTCapCuc_TSC)
				{
					if (cellTenXa != null)
					{
						sView.Cells[cellTenXa[1], cellTenXa[0]].Locked = true;
					}
					sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked = true;
					if (cellDiaChi != null)
					{
						sView.Cells[cellDiaChi[1], cellDiaChi[0]].Locked = true;
					}
					if (cellTenDonVi != null)
					{
						sView.Cells[cellTenDonVi[1], cellTenDonVi[0]].Locked = true;
					}
					if (cellMaSoDonVi != null)
					{
						sView.Cells[cellMaSoDonVi[1], cellMaSoDonVi[0]].Locked = true;
					}
				}
				else
				{
					if (string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc) || string.IsNullOrEmpty(maTinh))
					{
						return;
					}
					KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTPMoi.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
					sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text = keyValuePair.Key;
					sView.Cells[cellMaTinh[1], cellMaTinh[0]].Text = maTinh;
					sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked = true;
					string[] danhMucPhuongXaMoi = HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi(maTinh);
					if (danhMucPhuongXaMoi != null)
					{
						InsertCombobox(sView, cellMaXa[1] + 1, cellMaXa[0] + 1, danhMucPhuongXaMoi);
						if (cellTenXa != null && string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim()))
						{
							sView.Cells[cellTenXa[1], cellTenXa[0]].Value = "";
							sView.Cells[cellMaXa[1], cellMaXa[0]].Value = "";
						}
						if (cellDiaChi != null && string.IsNullOrEmpty(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Text.Trim()))
						{
							sView.Cells[cellDiaChi[1], cellDiaChi[0]].Value = "";
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ThayDoiTinhTheoCQT2Cap(HTKKCMCTTinChungInfo header, SheetView sView, HTKKPeriodInfo PeriodInfo, HTKKCMCSheetInfo sInfo, int startRow, int endRow, int[] cellBoQua1, int[] cellBoQua2)
	{
		try
		{
			IEnumerable<HTKKCMCCellInfo> enumerable = sInfo.StaticCells.Where((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex >= startRow && x.CellTemplate.RowIndex <= endRow);
			foreach (HTKKCMCCellInfo item in enumerable)
			{
				if ((cellBoQua1 != null && item.CellTemplate.RowIndex == cellBoQua1[1] + 1 && item.CellTemplate.ColIndex == cellBoQua1[0] + 1) || (cellBoQua2 != null && item.CellTemplate.RowIndex == cellBoQua2[1] + 1 && item.CellTemplate.ColIndex == cellBoQua2[0] + 1))
				{
					sView.Cells[item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1].Locked = true;
				}
				else
				{
					sView.Cells[item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1].Locked = ((header.maTinhNNT == PeriodInfo.MaTinh) ? true : false);
				}
			}
			if (header.maTinhNNT != PeriodInfo.MaTinh)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Controltype == "100");
				if (hTKKCMCCellInfo != null)
				{
					sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = PeriodInfo.TenTinh;
					sView.Cells[hTKKCMCCellInfo.CellSelectedValue.RowIndex - 1, hTKKCMCCellInfo.CellSelectedValue.ColIndex - 1].Value = PeriodInfo.MaTinh;
					sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Locked = true;
				}
				HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Controltype == "101");
				if (hTKKCMCCellInfo2 != null)
				{
					sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = PeriodInfo.TenXa;
					sView.Cells[hTKKCMCCellInfo2.CellSelectedValue.RowIndex - 1, hTKKCMCCellInfo2.CellSelectedValue.ColIndex - 1].Value = PeriodInfo.MaXa;
					sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Locked = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ThayDoiTinhTheoCQT2Cap(string maCQTCapCuc_TSC, string maCQTNoiNop_TSC, SheetView sView, HTKKPeriodInfo PeriodInfo, int[] cellTenTinh, int[] cellMaTinh, int[] cellTenHuyen, int[] cellMaHuyen, int[] cellTenXa, int[] cellMaXa, int[] cellDiaChi, int[] cellTenDonVi, int[] cellMaSoDonVi, bool PhuThuocTruSoChinh = true)
	{
		try
		{
			if (PeriodInfo.maCQTCapCuc == "82500")
			{
				sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked = false;
			}
			else
			{
				if (!string.IsNullOrEmpty(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text.Trim()) && cellTenXa != null && !string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim()))
				{
					return;
				}
				string maTinhCu = (HTKKWinDictionaryDiaDanh.dicTinhTP.ContainsValue(PeriodInfo.maCQTCapCuc.Substring(0, 3)) ? PeriodInfo.maCQTCapCuc.Substring(0, 3) : "");
				string maTinh = GetMaTinhMoiByMaTinhCu(maTinhCu);
				if (PhuThuocTruSoChinh && PeriodInfo.maCQTCapCuc == maCQTCapCuc_TSC)
				{
					if (cellTenXa != null)
					{
						sView.Cells[cellTenXa[1], cellTenXa[0]].Locked = true;
					}
					sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked = true;
					if (cellDiaChi != null)
					{
						sView.Cells[cellDiaChi[1], cellDiaChi[0]].Locked = true;
					}
					if (cellTenDonVi != null)
					{
						sView.Cells[cellTenDonVi[1], cellTenDonVi[0]].Locked = true;
					}
					if (cellMaSoDonVi != null)
					{
						sView.Cells[cellMaSoDonVi[1], cellMaSoDonVi[0]].Locked = true;
					}
				}
				else if (string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy) && !string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc) && !string.IsNullOrEmpty(maTinh))
				{
					sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text = HTKKWinDictionaryDiaDanh.dicTinhTPMoi[maTinh];
					sView.Cells[cellMaTinh[1], cellMaTinh[0]].Text = maTinh;
					sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked = true;
				}
				else if (!string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy) && !string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc) && !string.IsNullOrEmpty(maTinh))
				{
					KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTPMoi.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
					sView.Cells[cellTenTinh[1], cellTenTinh[0]].Value = keyValuePair.Key;
					sView.Cells[cellMaTinh[1], cellMaTinh[0]].Value = maTinh;
					sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool CheckTinhTheoCQT(string maCQTCapCuc_TSC, string maCQTNoiNop_TSC, SheetView sView, HTKKPeriodInfo PeriodInfo, int[] cellTenTinh, int[] cellMaTinh, int[] cellTenHuyen, int[] cellMaHuyen, int[] cellTenXa, int[] cellMaXa, int[] cellDiaChi, int[] cellTenDonVi = null, int[] cellMaSoDonVi = null, bool PhuThuocTruSoChinh = true, bool isBackGroudColor = false)
	{
		try
		{
			if (PeriodInfo.maCQTCapCuc == "82500")
			{
				return false;
			}
			XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
			string message = xMLMessage.GetMessage("0596");
			string message2 = xMLMessage.GetMessage("0047");
			bool result = false;
			if (Convert.ToBoolean(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked) && Convert.ToBoolean(sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Locked) && Convert.ToBoolean(sView.Cells[cellTenXa[1], cellTenXa[0]].Locked && cellDiaChi != null && Convert.ToBoolean(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Locked)))
			{
				return result;
			}
			if (isBackGroudColor)
			{
				if (cellDiaChi != null)
				{
					HTKKFile.RemoveValidateBackGround1(sView, cellDiaChi[1], cellDiaChi[0]);
				}
				HTKKFile.RemoveValidateBackGround1(sView, cellTenXa[1], cellTenXa[0]);
				HTKKFile.RemoveValidateBackGround1(sView, cellTenHuyen[1], cellTenHuyen[0]);
				HTKKFile.RemoveValidateBackGround1(sView, cellTenTinh[1], cellTenTinh[0]);
			}
			else
			{
				if (cellDiaChi != null)
				{
					HTKKFile.RemoveValidate(sView, cellDiaChi[1], cellDiaChi[0]);
				}
				HTKKFile.RemoveValidate(sView, cellTenXa[1], cellTenXa[0]);
				HTKKFile.RemoveValidate(sView, cellTenHuyen[1], cellTenHuyen[0]);
				HTKKFile.RemoveValidate(sView, cellTenTinh[1], cellTenTinh[0]);
			}
			if (cellTenDonVi != null)
			{
				HTKKFile.RemoveValidate(sView, cellTenDonVi[1], cellTenDonVi[0]);
			}
			if (cellMaSoDonVi != null)
			{
				HTKKFile.RemoveValidate(sView, cellMaSoDonVi[1], cellMaSoDonVi[0]);
			}
			if (!string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy) && PeriodInfo.maCQTQuanLy != maCQTNoiNop_TSC)
			{
				if (cellDiaChi != null && string.IsNullOrEmpty(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Text.Trim()) && !sView.Cells[cellDiaChi[1], cellDiaChi[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellDiaChi[1], cellDiaChi[0], message);
					result = true;
				}
				if (cellTenXa != null && string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim()) && !sView.Cells[cellTenXa[1], cellTenXa[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenXa[1], cellTenXa[0], message);
					result = true;
				}
				if (string.IsNullOrEmpty(sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Text.Trim()) && !sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenHuyen[1], cellTenHuyen[0], message);
					result = true;
				}
				if (string.IsNullOrEmpty(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text.Trim()) && !sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenTinh[1], cellTenTinh[0], message);
					result = true;
				}
			}
			if (!string.IsNullOrEmpty(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text.Trim()) && (string.IsNullOrEmpty(sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Text.Trim()) || (cellTenXa != null && string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim())) || (cellDiaChi != null && string.IsNullOrEmpty(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Text.Trim()))))
			{
				if (cellDiaChi != null && string.IsNullOrEmpty(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Text.Trim()) && !sView.Cells[cellDiaChi[1], cellDiaChi[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellDiaChi[1], cellDiaChi[0], message);
					result = true;
				}
				if (cellTenXa != null && string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim()) && !sView.Cells[cellTenXa[1], cellTenXa[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenXa[1], cellTenXa[0], message);
					result = true;
				}
				if (string.IsNullOrEmpty(sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Text.Trim()) && !sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenHuyen[1], cellTenHuyen[0], message);
					result = true;
				}
				if (string.IsNullOrEmpty(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text.Trim()) && !sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenTinh[1], cellTenTinh[0], message);
					result = true;
				}
			}
			if (PhuThuocTruSoChinh && cellTenDonVi != null && cellMaSoDonVi != null)
			{
				string text = sView.Cells[cellMaSoDonVi[1], cellMaSoDonVi[0]].Text.Trim();
				if (string.IsNullOrEmpty(text) && !sView.Cells[cellMaSoDonVi[1], cellMaSoDonVi[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellMaSoDonVi[1], cellMaSoDonVi[0], message);
					result = true;
				}
				if (!string.IsNullOrEmpty(text) && text.Length != 5 && !ValidateBase.CheckMaSoThue(text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellMaSoDonVi[1], cellMaSoDonVi[0], message2);
					result = true;
				}
				else if (!string.IsNullOrEmpty(text) && text.Length == 5 && !ValidateBase.CheckMaSoThue5So(text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellMaSoDonVi[1], cellMaSoDonVi[0], message2);
					result = true;
				}
				string value = sView.Cells[cellTenDonVi[1], cellTenDonVi[0]].Text.Trim();
				if (string.IsNullOrEmpty(value) && !sView.Cells[cellTenDonVi[1], cellTenDonVi[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenDonVi[1], cellTenDonVi[0], message);
					result = true;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return true;
		}
	}

	public static bool CheckTinhTheoCQTDiaBan(string maCQTCapCuc_TSC, string maCQTNoiNop_TSC, SheetView sView, HTKKPeriodInfo PeriodInfo, int[] cellTenTinh, int[] cellMaTinh, int[] cellTenXa, int[] cellMaXa, int[] cellDiaChi, int[] cellTenDonVi = null, int[] cellMaSoDonVi = null, bool PhuThuocTruSoChinh = true, bool isBackGroudColor = false)
	{
		try
		{
			if (PeriodInfo.maCQTCapCuc == "82500")
			{
				return false;
			}
			XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
			string message = xMLMessage.GetMessage("0596");
			string message2 = xMLMessage.GetMessage("0047");
			bool result = false;
			if (Convert.ToBoolean(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked) && Convert.ToBoolean(sView.Cells[cellTenXa[1], cellTenXa[0]].Locked && cellDiaChi != null && Convert.ToBoolean(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Locked)))
			{
				return result;
			}
			if (isBackGroudColor)
			{
				if (cellDiaChi != null)
				{
					HTKKFile.RemoveValidateBackGround1(sView, cellDiaChi[1], cellDiaChi[0]);
				}
				HTKKFile.RemoveValidateBackGround1(sView, cellTenXa[1], cellTenXa[0]);
				HTKKFile.RemoveValidateBackGround1(sView, cellTenTinh[1], cellTenTinh[0]);
			}
			else
			{
				if (cellDiaChi != null)
				{
					HTKKFile.RemoveValidate(sView, cellDiaChi[1], cellDiaChi[0]);
				}
				HTKKFile.RemoveValidate(sView, cellTenXa[1], cellTenXa[0]);
				HTKKFile.RemoveValidate(sView, cellTenTinh[1], cellTenTinh[0]);
			}
			if (cellTenDonVi != null)
			{
				HTKKFile.RemoveValidate(sView, cellTenDonVi[1], cellTenDonVi[0]);
			}
			if (cellMaSoDonVi != null)
			{
				HTKKFile.RemoveValidate(sView, cellMaSoDonVi[1], cellMaSoDonVi[0]);
			}
			if (!string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy) && PeriodInfo.maCQTQuanLy != maCQTNoiNop_TSC)
			{
				if (cellDiaChi != null && string.IsNullOrEmpty(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Text.Trim()) && !sView.Cells[cellDiaChi[1], cellDiaChi[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellDiaChi[1], cellDiaChi[0], message);
					result = true;
				}
				if (cellTenXa != null && string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim()) && !sView.Cells[cellTenXa[1], cellTenXa[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenXa[1], cellTenXa[0], message);
					result = true;
				}
				if (string.IsNullOrEmpty(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text.Trim()) && !sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenTinh[1], cellTenTinh[0], message);
					result = true;
				}
			}
			if (!string.IsNullOrEmpty(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text.Trim()) || (cellTenXa != null && string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim())) || (cellDiaChi != null && string.IsNullOrEmpty(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Text.Trim())))
			{
				if (cellDiaChi != null && string.IsNullOrEmpty(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Text.Trim()) && !sView.Cells[cellDiaChi[1], cellDiaChi[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellDiaChi[1], cellDiaChi[0], message);
					result = true;
				}
				if (cellTenXa != null && string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim()) && !sView.Cells[cellTenXa[1], cellTenXa[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenXa[1], cellTenXa[0], message);
					result = true;
				}
				if (string.IsNullOrEmpty(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text.Trim()) && !sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenTinh[1], cellTenTinh[0], message);
					result = true;
				}
			}
			if (PhuThuocTruSoChinh && cellTenDonVi != null && cellMaSoDonVi != null)
			{
				string text = sView.Cells[cellMaSoDonVi[1], cellMaSoDonVi[0]].Text.Trim();
				if (string.IsNullOrEmpty(text) && !sView.Cells[cellMaSoDonVi[1], cellMaSoDonVi[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellMaSoDonVi[1], cellMaSoDonVi[0], message);
					result = true;
				}
				if (!string.IsNullOrEmpty(text) && text.Length != 5 && !ValidateBase.CheckMaSoThue(text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellMaSoDonVi[1], cellMaSoDonVi[0], message2);
					result = true;
				}
				else if (!string.IsNullOrEmpty(text) && text.Length == 5 && !ValidateBase.CheckMaSoThue5So(text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellMaSoDonVi[1], cellMaSoDonVi[0], message2);
					result = true;
				}
				string value = sView.Cells[cellTenDonVi[1], cellTenDonVi[0]].Text.Trim();
				if (string.IsNullOrEmpty(value) && !sView.Cells[cellTenDonVi[1], cellTenDonVi[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenDonVi[1], cellTenDonVi[0], message);
					result = true;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return true;
		}
	}

	public static bool CheckTinhTheoCQTDbhc(HTKKCMCTTinChungInfo header, SheetView sView, HTKKPeriodInfo PeriodInfo, HTKKCMCSheetInfo sInfo, int startRow, int endRow)
	{
		try
		{
			XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
			string message = xMLMessage.GetMessage("0596");
			bool result = false;
			IEnumerable<HTKKCMCCellInfo> enumerable = sInfo.StaticCells.Where((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex >= startRow && x.CellTemplate.RowIndex <= endRow);
			foreach (HTKKCMCCellInfo item in enumerable)
			{
				HTKKFile.RemoveValidateBackGround1(sView, item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1);
			}
			foreach (HTKKCMCCellInfo item2 in enumerable)
			{
				Cell cell = sView.Cells[item2.CellTemplate.RowIndex - 1, item2.CellTemplate.ColIndex - 1];
				if (!cell.Locked && string.IsNullOrEmpty(cell.Text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, item2.CellTemplate.RowIndex - 1, item2.CellTemplate.ColIndex - 1, message);
					result = true;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return true;
		}
	}

	public static bool CheckTinhTheoCQTDbhc(string maCQTCapCuc_TSC, string maCQTNoiNop_TSC, SheetView sView, HTKKPeriodInfo PeriodInfo, int[] cellTenTinh, int[] cellMaTinh, int[] cellTenHuyen, int[] cellMaHuyen, int[] cellTenXa, int[] cellMaXa, int[] cellDiaChi, int[] cellTenDonVi = null, int[] cellMaSoDonVi = null, bool PhuThuocTruSoChinh = true, bool isBackGroudColor = false)
	{
		try
		{
			if (PeriodInfo.maCQTCapCuc == "82500")
			{
				return false;
			}
			XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
			string message = xMLMessage.GetMessage("0596");
			string message2 = xMLMessage.GetMessage("0047");
			bool result = false;
			if (Convert.ToBoolean(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked) && Convert.ToBoolean(sView.Cells[cellTenXa[1], cellTenXa[0]].Locked && cellDiaChi != null && Convert.ToBoolean(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Locked)))
			{
				return result;
			}
			if (isBackGroudColor)
			{
				if (cellDiaChi != null)
				{
					HTKKFile.RemoveValidateBackGround1(sView, cellDiaChi[1], cellDiaChi[0]);
				}
				HTKKFile.RemoveValidateBackGround1(sView, cellTenXa[1], cellTenXa[0]);
				HTKKFile.RemoveValidateBackGround1(sView, cellTenTinh[1], cellTenTinh[0]);
			}
			else
			{
				if (cellDiaChi != null)
				{
					HTKKFile.RemoveValidate(sView, cellDiaChi[1], cellDiaChi[0]);
				}
				HTKKFile.RemoveValidate(sView, cellTenXa[1], cellTenXa[0]);
				HTKKFile.RemoveValidate(sView, cellTenTinh[1], cellTenTinh[0]);
			}
			if (cellTenDonVi != null)
			{
				HTKKFile.RemoveValidate(sView, cellTenDonVi[1], cellTenDonVi[0]);
			}
			if (cellMaSoDonVi != null)
			{
				HTKKFile.RemoveValidate(sView, cellMaSoDonVi[1], cellMaSoDonVi[0]);
			}
			if (!string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy) && PeriodInfo.maCQTQuanLy != maCQTNoiNop_TSC)
			{
				if (cellDiaChi != null && string.IsNullOrEmpty(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Text.Trim()) && !sView.Cells[cellDiaChi[1], cellDiaChi[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellDiaChi[1], cellDiaChi[0], message);
					result = true;
				}
				if (cellTenXa != null && string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim()) && !sView.Cells[cellTenXa[1], cellTenXa[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenXa[1], cellTenXa[0], message);
					result = true;
				}
				if (string.IsNullOrEmpty(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text.Trim()) && !sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenTinh[1], cellTenTinh[0], message);
					result = true;
				}
			}
			if (!string.IsNullOrEmpty(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text.Trim()) && ((cellTenXa != null && string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim())) || (cellDiaChi != null && string.IsNullOrEmpty(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Text.Trim()))))
			{
				if (cellDiaChi != null && string.IsNullOrEmpty(sView.Cells[cellDiaChi[1], cellDiaChi[0]].Text.Trim()) && !sView.Cells[cellDiaChi[1], cellDiaChi[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellDiaChi[1], cellDiaChi[0], message);
					result = true;
				}
				if (cellTenXa != null && string.IsNullOrEmpty(sView.Cells[cellTenXa[1], cellTenXa[0]].Text.Trim()) && !sView.Cells[cellTenXa[1], cellTenXa[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenXa[1], cellTenXa[0], message);
					result = true;
				}
				if (string.IsNullOrEmpty(sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text.Trim()) && !sView.Cells[cellTenTinh[1], cellTenTinh[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenTinh[1], cellTenTinh[0], message);
					result = true;
				}
			}
			if (PhuThuocTruSoChinh && cellTenDonVi != null && cellMaSoDonVi != null)
			{
				string text = sView.Cells[cellMaSoDonVi[1], cellMaSoDonVi[0]].Text.Trim();
				if (string.IsNullOrEmpty(text) && !sView.Cells[cellMaSoDonVi[1], cellMaSoDonVi[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellMaSoDonVi[1], cellMaSoDonVi[0], message);
					result = true;
				}
				if (!string.IsNullOrEmpty(text) && text.Length != 5 && !ValidateBase.CheckMaSoThue(text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellMaSoDonVi[1], cellMaSoDonVi[0], message2);
					result = true;
				}
				else if (!string.IsNullOrEmpty(text) && text.Length == 5 && !ValidateBase.CheckMaSoThue5So(text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellMaSoDonVi[1], cellMaSoDonVi[0], message2);
					result = true;
				}
				string value = sView.Cells[cellTenDonVi[1], cellTenDonVi[0]].Text.Trim();
				if (string.IsNullOrEmpty(value) && !sView.Cells[cellTenDonVi[1], cellTenDonVi[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cellTenDonVi[1], cellTenDonVi[0], message);
					result = true;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return true;
		}
	}

	public static bool CheckTinhTheoCQTDangBangKhac_TSC(string maCQTCapCuc_TSC, string maCQTNoiNop_TSC, SheetView sView, int colTenTinh, int colMaTinh, int colTenHuyen, int colMaHuyen, int colTenXa, int colMaXa, int colDiaChi, int iRow)
	{
		try
		{
			if (maCQTCapCuc_TSC == "82500")
			{
				return false;
			}
			XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
			string message = xMLMessage.GetMessage("0596");
			string message2 = xMLMessage.GetMessage("0574");
			bool result = false;
			if (Convert.ToBoolean(sView.Cells[iRow - 1, colTenTinh - 1].Locked) && Convert.ToBoolean(sView.Cells[iRow - 1, colTenHuyen - 1].Locked) && colTenXa >= 0 && Convert.ToBoolean(sView.Cells[iRow - 1, colTenXa - 1].Locked && colDiaChi >= 0 && Convert.ToBoolean(sView.Cells[iRow - 1, colDiaChi - 1].Locked)))
			{
				return result;
			}
			if (colDiaChi >= 0)
			{
				HTKKFile.RemoveValidate(sView, iRow - 1, colDiaChi - 1);
			}
			if (colTenXa >= 0)
			{
				HTKKFile.RemoveValidate(sView, iRow - 1, colTenXa - 1);
			}
			HTKKFile.RemoveValidate(sView, iRow - 1, colTenHuyen - 1);
			HTKKFile.RemoveValidate(sView, iRow - 1, colTenTinh - 1);
			string text = sView.Cells[iRow - 1, colMaTinh - 1].Text.Trim();
			string text2 = (HTKKWinDictionaryDiaDanh.dicTinhTP.ContainsValue(maCQTCapCuc_TSC.Substring(0, 3)) ? maCQTCapCuc_TSC.Substring(0, 3) : "");
			if (text == text2)
			{
				HTKKFile.SetNodeBackGroundPink(sView, iRow - 1, colTenTinh - 1, message2);
				result = true;
			}
			if (!string.IsNullOrEmpty(sView.Cells[iRow - 1, colTenTinh - 1].Text.Trim()) && (string.IsNullOrEmpty(sView.Cells[iRow - 1, colTenHuyen - 1].Text.Trim()) || (colTenXa >= 0 && string.IsNullOrEmpty(sView.Cells[iRow - 1, colTenXa - 1].Text.Trim())) || (colDiaChi >= 0 && string.IsNullOrEmpty(sView.Cells[iRow - 1, colDiaChi - 1].Text.Trim()))))
			{
				if (colDiaChi >= 0 && string.IsNullOrEmpty(sView.Cells[iRow - 1, colDiaChi - 1].Text.Trim()) && !sView.Cells[iRow - 1, colDiaChi - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, iRow - 1, colDiaChi - 1, message);
					result = true;
				}
				if (colTenXa >= 0 && string.IsNullOrEmpty(sView.Cells[iRow - 1, colTenXa - 1].Text.Trim()) && !sView.Cells[iRow - 1, colTenXa - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, iRow - 1, colTenXa - 1, message);
					result = true;
				}
				if (string.IsNullOrEmpty(sView.Cells[iRow - 1, colTenHuyen - 1].Text.Trim()) && !sView.Cells[iRow - 1, colTenHuyen - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, iRow - 1, colTenHuyen - 1, message);
					result = true;
				}
				if (string.IsNullOrEmpty(sView.Cells[iRow - 1, colTenTinh - 1].Text.Trim()) && !sView.Cells[iRow - 1, colTenTinh - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sView, iRow - 1, colTenTinh - 1, message);
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

	public static void SaveAllCellTypeOfSheet(SheetView sView, HTKKCMCSheetInfo sInfo, int sheetIndex)
	{
		try
		{
			if (sInfo == null)
			{
				return;
			}
			List<HTKKCMCTableInfo> tables = sInfo.Tables;
			foreach (HTKKCMCTableInfo item in tables)
			{
				int dataLocationTemplateRowStart = item.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < item.RowsValues.Count; i++)
				{
					foreach (HTKKCMCTableCellInfo cell in item.RowsInfo.Cells)
					{
						UpsertCellType(sheetIndex, dataLocationTemplateRowStart + i, cell.CellTemplate.ColIndex, sView, dataLocationTemplateRowStart + i, cell.CellTemplate.ColIndex);
					}
				}
			}
			foreach (HTKKCMCCellInfo staticCell in sInfo.StaticCells)
			{
				if (!UpsertCellType(sInfo.SheetIndex, staticCell.CellTemplate.RowIndex, staticCell.CellTemplate.ColIndex, sView, staticCell.CellTemplate.RowIndex, staticCell.CellTemplate.ColIndex))
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name + " loi KHBS: xem lai path: " + staticCell.Path);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool UpsertCellType(int sheetIndex_KHBS, int row_KHBS, int col_KHBS, SheetView sViewToDieuChinh, int rowToDieuChinh, int colDieuChinh)
	{
		try
		{
			if (rowToDieuChinh < 1)
			{
				return true;
			}
			if (ConvertHelpers._dicCellType_01_1_KHBS == null)
			{
				ConvertHelpers._dicCellType_01_1_KHBS = new Dictionary<string, CellTypeKHBS>();
			}
			CellTypeKHBS cellTypeKHBS = new CellTypeKHBS();
			ICellType cellType = sViewToDieuChinh.Cells[rowToDieuChinh - 1, colDieuChinh - 1].CellType;
			cellTypeKHBS.row = row_KHBS;
			cellTypeKHBS.col = col_KHBS;
			cellTypeKHBS.sheetIndex = sheetIndex_KHBS;
			cellTypeKHBS.iCellType = cellType;
			List<string> list = new List<string>();
			list.Add(sheetIndex_KHBS.ToString());
			list.Add(row_KHBS.ToString());
			list.Add(col_KHBS.ToString());
			string key = FormatKey(list, "___");
			if (!ConvertHelpers._dicCellType_01_1_KHBS.ContainsKey(key))
			{
				ConvertHelpers._dicCellType_01_1_KHBS.Add(key, cellTypeKHBS);
			}
			else
			{
				ConvertHelpers._dicCellType_01_1_KHBS[key] = cellTypeKHBS;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
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

	public static void ThayDoiTinhTheoCQTDangBang(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKPeriodInfo PeriodInfo, int tableIndex, int colTenTinh, int colMaTinh, int colTenHuyen, int colMaHuyen, int colTenXa, int colMaXa, int iRow = -1, bool isLock = true)
	{
		try
		{
			if (string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc))
			{
				return;
			}
			string maTinh = (HTKKWinDictionaryDiaDanh.dicTinhTP.ContainsValue(PeriodInfo.maCQTCapCuc.Substring(0, 3)) ? PeriodInfo.maCQTCapCuc.Substring(0, 3) : "");
			KeyValuePair<string, string> keyValuePair;
			if (!string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy) && !string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc) && PeriodInfo.maCQTQuanLy != PeriodInfo.maCQTCapCuc)
			{
				string maHuyen = (HTKKWinDictionaryDiaDanh.dicQuanHuyen.ContainsKey(PeriodInfo.maCQTQuanLy) ? PeriodInfo.maCQTQuanLy : "");
				if (string.IsNullOrEmpty(maHuyen) || string.IsNullOrEmpty(maTinh))
				{
					return;
				}
				keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
				{
					foreach (HTKKCMCTableInfo table in sInfo.Tables)
					{
						if (tableIndex >= 0 && (tableIndex < 0 || !(table.TableName == sInfo.Tables[tableIndex].TableName)))
						{
							continue;
						}
						int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
						if (iRow < 0)
						{
							for (int num = 0; num < table.RowsValues.Count; num++)
							{
								sView.Cells[dataLocationTemplateRowStart + num - 1, colTenTinh - 1].Text = keyValuePair.Key;
								sView.Cells[dataLocationTemplateRowStart + num - 1, colMaTinh - 1].Text = maTinh;
								sView.Cells[dataLocationTemplateRowStart + num - 1, colTenHuyen - 1].Text = HTKKWinDictionaryDiaDanh.dicQuanHuyen[maHuyen];
								sView.Cells[dataLocationTemplateRowStart + num - 1, colMaHuyen - 1].Text = maHuyen;
								if (colTenXa >= 0)
								{
									Dictionary<string, string> dic = HTKKWinDictionaryDiaDanh.dicPhuongXa.Where((KeyValuePair<string, string> x) => x.Key.Substring(0, 5) == maHuyen).ToDictionary((KeyValuePair<string, string> x) => x.Key, (KeyValuePair<string, string> x) => x.Value);
									InsertCombobox(sView, dataLocationTemplateRowStart + num, colTenXa, dic, deleteText: false);
								}
								LockUnLockCol(sView, dataLocationTemplateRowStart + num, colTenTinh, isLock, deleteData: false);
								LockUnLockCol(sView, dataLocationTemplateRowStart + num, colTenHuyen, isLock, deleteData: false);
							}
							continue;
						}
						sView.Cells[iRow - 1, colTenTinh - 1].Text = keyValuePair.Key;
						sView.Cells[iRow - 1, colMaTinh - 1].Text = maTinh;
						sView.Cells[iRow - 1, colTenHuyen - 1].Text = HTKKWinDictionaryDiaDanh.dicQuanHuyen[maHuyen];
						sView.Cells[iRow - 1, colMaHuyen - 1].Text = maHuyen;
						if (colTenXa >= 0)
						{
							Dictionary<string, string> dic = HTKKWinDictionaryDiaDanh.dicPhuongXa.Where((KeyValuePair<string, string> x) => x.Key.Substring(0, 5) == maHuyen).ToDictionary((KeyValuePair<string, string> x) => x.Key, (KeyValuePair<string, string> x) => x.Value);
							InsertCombobox(sView, iRow, colTenXa, dic, deleteText: false);
						}
						LockUnLockCol(sView, iRow, colTenTinh, isLock, deleteData: false);
						LockUnLockCol(sView, iRow, colTenHuyen, isLock, deleteData: false);
					}
					return;
				}
			}
			if (string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc))
			{
				return;
			}
			keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
			foreach (HTKKCMCTableInfo table2 in sInfo.Tables)
			{
				if (tableIndex >= 0 && (tableIndex < 0 || !(table2.TableName == sInfo.Tables[tableIndex].TableName)))
				{
					continue;
				}
				if (iRow < 0)
				{
					int dataLocationTemplateRowStart = table2.Location.DataLocationTemplateRowStart;
					for (int num = 0; num < table2.RowsValues.Count; num++)
					{
						sView.Cells[dataLocationTemplateRowStart + num - 1, colTenTinh - 1].Text = (string.IsNullOrEmpty(maTinh) ? "" : keyValuePair.Key);
						sView.Cells[dataLocationTemplateRowStart + num - 1, colMaTinh - 1].Text = maTinh;
						if (!string.IsNullOrEmpty(maTinh))
						{
							Dictionary<string, string> dictionary = HTKKWinDictionaryDiaDanh.dicQuanHuyen.Where((KeyValuePair<string, string> x) => x.Key.StartsWith(maTinh)).ToDictionary((KeyValuePair<string, string> x) => x.Key, (KeyValuePair<string, string> x) => x.Value);
							if (dictionary != null)
							{
								if (string.IsNullOrEmpty(sView.Cells[dataLocationTemplateRowStart + num - 1, colTenHuyen - 1].Text.Trim()))
								{
									sView.Cells[dataLocationTemplateRowStart + num - 1, colTenHuyen - 1].Value = "";
									sView.Cells[dataLocationTemplateRowStart + num - 1, colMaHuyen - 1].Value = "";
								}
								InsertCombobox(sView, dataLocationTemplateRowStart + num, colTenHuyen, dictionary);
							}
						}
						LockUnLockCol(sView, dataLocationTemplateRowStart + num, colTenTinh, !string.IsNullOrEmpty(maTinh), deleteData: false);
						LockUnLockCol(sView, dataLocationTemplateRowStart + num, colTenHuyen, locked: false, deleteData: false);
					}
					continue;
				}
				sView.Cells[iRow - 1, colTenTinh - 1].Text = (string.IsNullOrEmpty(maTinh) ? "" : keyValuePair.Key);
				sView.Cells[iRow - 1, colMaTinh - 1].Text = maTinh;
				if (!string.IsNullOrEmpty(maTinh))
				{
					Dictionary<string, string> dictionary = HTKKWinDictionaryDiaDanh.dicQuanHuyen.Where((KeyValuePair<string, string> x) => x.Key.StartsWith(maTinh)).ToDictionary((KeyValuePair<string, string> x) => x.Key, (KeyValuePair<string, string> x) => x.Value);
					if (dictionary != null)
					{
						if (string.IsNullOrEmpty(sView.Cells[iRow - 1, colTenHuyen - 1].Text.Trim()))
						{
							sView.Cells[iRow - 1, colTenHuyen - 1].Value = "";
							sView.Cells[iRow - 1, colMaHuyen - 1].Value = "";
						}
						InsertCombobox(sView, iRow, colTenHuyen, dictionary);
					}
				}
				LockUnLockCol(sView, iRow, colTenTinh, !string.IsNullOrEmpty(maTinh), deleteData: false);
				LockUnLockCol(sView, iRow, colTenHuyen, locked: false, deleteData: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ThayDoiTinhTheoCQTDangBangDiaBan(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKPeriodInfo PeriodInfo, int tableIndex, int colTenTinh, int colMaTinh, int colTenHuyen, int colMaHuyen, int colTenXa, int colMaXa, int iRow = -1, bool isLock = true)
	{
		try
		{
			if (string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc))
			{
				return;
			}
			string maTinh = (HTKKWinDictionaryDiaDanh.dicTinhTPMoi.ContainsValue(PeriodInfo.maCQTCapCuc.Substring(0, 3)) ? PeriodInfo.maCQTCapCuc.Substring(0, 3) : "");
			if (string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc))
			{
				return;
			}
			KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTPMoi.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (tableIndex >= 0 && (tableIndex < 0 || !(table.TableName == sInfo.Tables[tableIndex].TableName)))
				{
					continue;
				}
				if (iRow < 0)
				{
					int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
					for (int num = 0; num < table.RowsValues.Count; num++)
					{
						sView.Cells[dataLocationTemplateRowStart + num - 1, colTenTinh - 1].Text = (string.IsNullOrEmpty(maTinh) ? "" : keyValuePair.Key);
						sView.Cells[dataLocationTemplateRowStart + num - 1, colMaTinh - 1].Text = maTinh;
						if (!string.IsNullOrEmpty(maTinh))
						{
							string[] danhMucPhuongXaMoi = HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi(maTinh);
							if (danhMucPhuongXaMoi != null)
							{
								if (string.IsNullOrEmpty(sView.Cells[dataLocationTemplateRowStart + num - 1, colTenHuyen - 1].Text.Trim()))
								{
									sView.Cells[dataLocationTemplateRowStart + num - 1, colTenHuyen - 1].Value = "";
									sView.Cells[dataLocationTemplateRowStart + num - 1, colMaHuyen - 1].Value = "";
								}
								InsertCombobox(sView, dataLocationTemplateRowStart + num, colTenHuyen, danhMucPhuongXaMoi);
							}
						}
						LockUnLockCol(sView, dataLocationTemplateRowStart + num, colTenTinh, !string.IsNullOrEmpty(maTinh), deleteData: false);
						LockUnLockCol(sView, dataLocationTemplateRowStart + num, colTenHuyen, locked: false, deleteData: false);
					}
					continue;
				}
				sView.Cells[iRow - 1, colTenTinh - 1].Text = (string.IsNullOrEmpty(maTinh) ? "" : keyValuePair.Key);
				sView.Cells[iRow - 1, colMaTinh - 1].Text = maTinh;
				if (!string.IsNullOrEmpty(maTinh))
				{
					string[] danhMucPhuongXaMoi = HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi(maTinh);
					if (danhMucPhuongXaMoi != null)
					{
						if (string.IsNullOrEmpty(sView.Cells[iRow - 1, colTenHuyen - 1].Text.Trim()))
						{
							sView.Cells[iRow - 1, colTenHuyen - 1].Value = "";
							sView.Cells[iRow - 1, colMaHuyen - 1].Value = "";
						}
						InsertCombobox(sView, iRow, colTenHuyen, danhMucPhuongXaMoi);
					}
				}
				LockUnLockCol(sView, iRow, colTenTinh, !string.IsNullOrEmpty(maTinh), deleteData: false);
				LockUnLockCol(sView, iRow, colTenHuyen, locked: false, deleteData: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ThayDoiTinhTheoCQTDangBang2Cap(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKPeriodInfo PeriodInfo, int tableIndex, int colTenTinh, int colMaTinh, int colTenHuyen, int colMaHuyen, int colTenXa, int colMaXa, int iRow = -1, bool isLock = true)
	{
		try
		{
			if (string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc))
			{
				return;
			}
			string maTinh = PeriodInfo.MaTinh;
			KeyValuePair<string, string> keyValuePair;
			if (!string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy) && !string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc) && PeriodInfo.maCQTQuanLy != PeriodInfo.maCQTCapCuc)
			{
				if (string.IsNullOrEmpty(maTinh))
				{
					return;
				}
				keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTPMoi.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
				{
					foreach (HTKKCMCTableInfo table in sInfo.Tables)
					{
						if (tableIndex >= 0 && (tableIndex < 0 || !(table.TableName == sInfo.Tables[tableIndex].TableName)))
						{
							continue;
						}
						int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
						if (iRow < 0)
						{
							for (int num = 0; num < table.RowsValues.Count; num++)
							{
								sView.Cells[dataLocationTemplateRowStart + num - 1, colTenTinh - 1].Text = keyValuePair.Key;
								sView.Cells[dataLocationTemplateRowStart + num - 1, colMaTinh - 1].Text = maTinh;
								if (colTenXa >= 0)
								{
									Dictionary<string, string> danhMucPhuongXaMoi = HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi1(maTinh);
									InsertCombobox(sView, dataLocationTemplateRowStart + num, colTenXa, danhMucPhuongXaMoi, deleteText: false);
								}
								LockUnLockCol(sView, dataLocationTemplateRowStart + num, colTenTinh, isLock, deleteData: false);
							}
						}
						else
						{
							sView.Cells[iRow - 1, colTenTinh - 1].Text = keyValuePair.Key;
							sView.Cells[iRow - 1, colMaTinh - 1].Text = maTinh;
							if (colTenXa >= 0)
							{
								Dictionary<string, string> danhMucPhuongXaMoi = HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi1(maTinh);
								InsertCombobox(sView, iRow, colTenXa, danhMucPhuongXaMoi, deleteText: false);
							}
							LockUnLockCol(sView, iRow, colTenTinh, isLock, deleteData: false);
						}
					}
					return;
				}
			}
			if (string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc))
			{
				return;
			}
			keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTPMoi.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
			foreach (HTKKCMCTableInfo table2 in sInfo.Tables)
			{
				if (tableIndex >= 0 && (tableIndex < 0 || !(table2.TableName == sInfo.Tables[tableIndex].TableName)))
				{
					continue;
				}
				if (iRow < 0)
				{
					int dataLocationTemplateRowStart = table2.Location.DataLocationTemplateRowStart;
					for (int num = 0; num < table2.RowsValues.Count; num++)
					{
						sView.Cells[dataLocationTemplateRowStart + num - 1, colTenTinh - 1].Text = (string.IsNullOrEmpty(maTinh) ? "" : keyValuePair.Key);
						sView.Cells[dataLocationTemplateRowStart + num - 1, colMaTinh - 1].Text = maTinh;
						if (!string.IsNullOrEmpty(maTinh) && colTenXa >= 0)
						{
							Dictionary<string, string> danhMucPhuongXaMoi = HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi1(maTinh);
							InsertCombobox(sView, dataLocationTemplateRowStart + num, colTenXa, danhMucPhuongXaMoi, deleteText: false);
						}
						LockUnLockCol(sView, dataLocationTemplateRowStart + num, colTenTinh, !string.IsNullOrEmpty(maTinh), deleteData: false);
					}
				}
				else
				{
					sView.Cells[iRow - 1, colTenTinh - 1].Text = (string.IsNullOrEmpty(maTinh) ? "" : keyValuePair.Key);
					sView.Cells[iRow - 1, colMaTinh - 1].Text = maTinh;
					if (!string.IsNullOrEmpty(maTinh) && colTenXa >= 0)
					{
						Dictionary<string, string> danhMucPhuongXaMoi = HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi1(maTinh);
						InsertCombobox(sView, iRow, colTenXa, danhMucPhuongXaMoi, deleteText: false);
					}
					LockUnLockCol(sView, iRow, colTenTinh, !string.IsNullOrEmpty(maTinh), deleteData: false);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ThayDoiTinhTheoCQTDangBang2(string maCQTCapCuc_TSC, string maCQTNoiNop_TSC, SheetView sView, int colTenTinh, int colMaTinh, int colTenHuyen, int colMaHuyen, int iRow, bool awaySet)
	{
		try
		{
			if (string.IsNullOrEmpty(maCQTCapCuc_TSC) || (!awaySet && !string.IsNullOrEmpty(sView.Cells[iRow - 1, colTenTinh - 1].Text.Trim()) && !string.IsNullOrEmpty(sView.Cells[iRow - 1, colTenHuyen - 1].Text.Trim())))
			{
				return;
			}
			string maTinh = maCQTCapCuc_TSC.Substring(0, 3);
			if (!string.IsNullOrEmpty(maCQTNoiNop_TSC) && !string.IsNullOrEmpty(maCQTCapCuc_TSC))
			{
				string text = (HTKKWinDictionaryDiaDanh.dicQuanHuyen.ContainsKey(maCQTNoiNop_TSC) ? maCQTNoiNop_TSC : "");
				if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(maTinh))
				{
					KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
					if (!string.IsNullOrEmpty(keyValuePair.Key))
					{
						sView.Cells[iRow - 1, colTenTinh - 1].Value = keyValuePair.Key;
						sView.Cells[iRow - 1, colMaTinh - 1].Value = maTinh;
						sView.Cells[iRow - 1, colTenTinh - 1].Locked = true;
						sView.Cells[iRow - 1, colTenHuyen - 1].Value = HTKKWinDictionaryDiaDanh.dicQuanHuyen[text];
						sView.Cells[iRow - 1, colMaHuyen - 1].Value = text;
						sView.Cells[iRow - 1, colTenHuyen - 1].Locked = true;
					}
				}
				if (string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(maTinh))
				{
					KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
					if (!string.IsNullOrEmpty(keyValuePair.Key))
					{
						sView.Cells[iRow - 1, colTenTinh - 1].Value = keyValuePair.Key;
						sView.Cells[iRow - 1, colMaTinh - 1].Value = maTinh;
						sView.Cells[iRow - 1, colTenTinh - 1].Locked = true;
						sView.Cells[iRow - 1, colTenHuyen - 1].Value = "";
						sView.Cells[iRow - 1, colMaHuyen - 1].Value = "";
						sView.Cells[iRow - 1, colTenHuyen - 1].Locked = false;
					}
				}
			}
			if (string.IsNullOrEmpty(maCQTNoiNop_TSC) && !string.IsNullOrEmpty(maCQTCapCuc_TSC) && !string.IsNullOrEmpty(maTinh))
			{
				KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
				if (!string.IsNullOrEmpty(keyValuePair.Key))
				{
					sView.Cells[iRow - 1, colTenTinh - 1].Value = keyValuePair.Key;
					sView.Cells[iRow - 1, colMaTinh - 1].Value = maTinh;
					sView.Cells[iRow - 1, colTenTinh - 1].Locked = true;
					sView.Cells[iRow - 1, colTenHuyen - 1].Value = "";
					sView.Cells[iRow - 1, colMaHuyen - 1].Value = "";
					sView.Cells[iRow - 1, colTenHuyen - 1].Locked = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool Compare2RowValueSame(SheetView sView, int row1, int row2, int colCompare)
	{
		try
		{
			string text = sView.Cells[row1 - 1, colCompare - 1].Value.ToString().Trim();
			string text2 = sView.Cells[row2 - 1, colCompare - 1].Value.ToString().Trim();
			if (text == text2)
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

	public static void SetAllValueInRowSame_Table(SheetView sView, HTKKCMCTableInfo table, List<int> cols, bool isLock, bool isDelete)
	{
		try
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 1; i < table.RowsValues.Count; i++)
			{
				foreach (int col in cols)
				{
					if (!sView.Cells[dataLocationTemplateRowStart - 1, col - 1].Locked)
					{
						LockUnLockCol(sView, dataLocationTemplateRowStart, col, locked: false, deleteData: false);
					}
					object value = sView.Cells[dataLocationTemplateRowStart - 1, col - 1].Value;
					ICellType cellType = sView.Cells[dataLocationTemplateRowStart - 1, col - 1].CellType;
					sView.Cells[dataLocationTemplateRowStart + i - 1, col - 1].Value = value;
					sView.Cells[dataLocationTemplateRowStart + i - 1, col - 1].CellType = cellType;
					LockUnLockCol(sView, dataLocationTemplateRowStart + i, col, isLock, isDelete);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CreateTableGroup(SheetView sView, HTKKCMCSheetInfo sInfo, string fromTableName, string toTableName, Dictionary<int, int> dicHeaderGroup, Dictionary<int, int> dicInfoHeader)
	{
		try
		{
			List<HTKKCMCTableInfo> list = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == fromTableName).ToList();
			List<HTKKCMCTableInfo> toTables = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == toTableName).ToList();
			int num = toTables[0].Header.StaticCells[0].CellTemplate.RowIndex;
			foreach (HTKKCMCTableInfo item in list)
			{
				foreach (int key in dicHeaderGroup.Keys)
				{
					int num2 = dicHeaderGroup[key];
					int rowIndex = item.Header.StaticCells[0].CellTemplate.RowIndex;
					foreach (int key2 in dicInfoHeader.Keys)
					{
						int num3 = dicInfoHeader[key2];
						string text = sView.Cells[rowIndex - 1, key2 - 1].Text.Trim();
						string value = sView.Cells[rowIndex - 1, key - 1].Text.Trim();
						if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(value))
						{
							continue;
						}
						toTables = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == toTableName).ToList();
						List<int> rowIndexInSViewByCompareHeader = GetRowIndexInSViewByCompareHeader(sView, toTables, num3, text.ToString().Trim());
						if (rowIndexInSViewByCompareHeader.Count >= 1)
						{
							continue;
						}
						if (num == toTables[0].Header.StaticCells[0].CellTemplate.RowIndex)
						{
							sView.Cells[num - 1, num3 - 1].Value = text;
							sView.Cells[num - 1, num2 - 1].Value = value;
							num++;
							continue;
						}
						int num4 = sInfo.Tables.FindIndex((HTKKCMCTableInfo x) => x.TableName == toTableName && x.TableID == toTables[toTables.Count - 1].TableID);
						InsertTableF7(sView, sInfo, num4);
						num = sInfo.Tables[num4 + 1].Header.StaticCells[0].CellTemplate.RowIndex;
						sView.Cells[num - 1, num3 - 1].Value = text;
						sView.Cells[num - 1, num2 - 1].Value = value;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CopyTableToTableGroup_F7_F7(SheetView sView, HTKKCMCSheetInfo sInfo, string fromTableName, string toTableName, Dictionary<int, int> dicHeaderGroup, Dictionary<int, int> dicInfoHeader, Dictionary<int, int> dicInfo, bool usingCellType = false)
	{
		try
		{
			DeleteTableGroup(sView, sInfo, toTableName);
			CreateTableGroup(sView, sInfo, fromTableName, toTableName, dicHeaderGroup, dicInfoHeader);
			List<HTKKCMCTableInfo> list = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == fromTableName).ToList();
			List<HTKKCMCTableInfo> list2 = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == toTableName).ToList();
			Dictionary<string, int> checkExists = new Dictionary<string, int>();
			foreach (HTKKCMCTableInfo item in list)
			{
				int rowIndex = item.Header.StaticCells[0].CellTemplate.RowIndex;
				foreach (int key in dicHeaderGroup.Keys)
				{
					int colCheck = dicHeaderGroup[key];
					object value = sView.Cells[rowIndex - 1, key - 1].Value;
					if (value != null)
					{
						int tableIndexInSViewByCompareHeader = GetTableIndexInSViewByCompareHeader(sView, sInfo, colCheck, value.ToString().Trim(), toTableName);
						CopyTableToTable2(sView, sInfo, item, tableIndexInSViewByCompareHeader, dicInfo, deleteDataOld: false, ref checkExists, usingCellType);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static int GetTableIndexInSViewByCompareHeader(SheetView sView, HTKKCMCSheetInfo sInfo, int colCheck, string valueCompare, string tableName)
	{
		try
		{
			int result = -1;
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
				if (hTKKCMCTableInfo.TableName == tableName)
				{
					int rowIndex = hTKKCMCTableInfo.Header.StaticCells[0].CellTemplate.RowIndex;
					if (sView.Cells[rowIndex - 1, colCheck - 1].Value != null && sView.Cells[rowIndex - 1, colCheck - 1].Value.ToString() == valueCompare)
					{
						result = i;
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

	public static void CopyTableToTable2(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKCMCTableInfo fromTable, int toTableIndex, Dictionary<int, int> dicInfo, bool deleteDataOld, ref Dictionary<string, int> checkExists, bool usingCellType = false)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[toTableIndex];
			int dataLocationTemplateRowStart = fromTable.Location.DataLocationTemplateRowStart;
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			if (deleteDataOld)
			{
				if (hTKKCMCTableInfo.RowsValues.Count > 1)
				{
					RemoveDataRows(sView, toTableIndex, num - 1, hTKKCMCTableInfo.RowsValues.Count - 1, sInfo);
				}
				foreach (HTKKCMCTableCellInfo cell2 in hTKKCMCTableInfo.RowsInfo.Cells)
				{
					sView.Cells[num - 1, cell2.CellTemplate.ColIndex - 1].Value = cell2.DefaultValue;
				}
			}
			NumberCellType numberCellType = new NumberCellType();
			CurrencyCellType currencyCellType = new CurrencyCellType();
			TextCellType textCellType = new TextCellType();
			GeneralCellType generalCellType = new GeneralCellType();
			for (int i = 0; i < fromTable.RowsValues.Count; i++)
			{
				string text = "";
				foreach (int key in dicInfo.Keys)
				{
					int num2 = dicInfo[key];
					Cell cell = sView.Cells[dataLocationTemplateRowStart + i - 1, key - 1];
					ICellType cellType = sView.Cells[dataLocationTemplateRowStart + i - 1, key - 1].CellType;
					if (cell != null && (object)cellType.GetType() != numberCellType.GetType() && (object)cellType.GetType() != currencyCellType.GetType())
					{
						if (string.IsNullOrEmpty(cell.Text.Trim()))
						{
							text = "";
							break;
						}
						text += cell.Text.Trim();
					}
				}
				if (string.IsNullOrEmpty(text))
				{
					continue;
				}
				text = hTKKCMCTableInfo.TableName + hTKKCMCTableInfo.TableID + text;
				if (checkExists.ContainsKey(text))
				{
					num = checkExists[text];
				}
				else
				{
					object value = sView.Cells[num - 1, dicInfo.Values.ToArray()[0] - 1].Value;
					if (value != null && !string.IsNullOrEmpty(value.ToString().Trim()))
					{
						num = checkExists.Last().Value;
						checkExists.Add(text, num + 1);
						InsertTableRow(sView, num, -1, num, -1, -1, sInfo, toTableIndex);
						num++;
					}
					else
					{
						checkExists.Add(text, num);
					}
				}
				foreach (int key2 in dicInfo.Keys)
				{
					int num2 = dicInfo[key2];
					object value2 = sView.Cells[dataLocationTemplateRowStart + i - 1, key2 - 1].Value;
					ICellType cellType = sView.Cells[dataLocationTemplateRowStart + i - 1, key2 - 1].CellType;
					if (value2 == null)
					{
						continue;
					}
					if ((object)cellType.GetType() == numberCellType.GetType() || (object)cellType.GetType() == currencyCellType.GetType())
					{
						if (usingCellType)
						{
							sView.Cells[num - 1, num2 - 1].CellType = cellType;
						}
						sView.Cells[num - 1, num2 - 1].Value = Convert.ToDouble(sView.Cells[num - 1, num2 - 1].Value) + Convert.ToDouble(value2);
					}
					else
					{
						sView.Cells[num - 1, num2 - 1].Value = value2;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void DeleteTableGroup(SheetView sView, HTKKCMCSheetInfo sInfo, string tableName)
	{
		try
		{
			List<HTKKCMCTableInfo> tables = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == tableName).ToList();
			int num = sInfo.Tables.FindIndex((HTKKCMCTableInfo x) => x.TableName == tableName && x.TableID == tables[0].TableID);
			int num2 = sInfo.Tables.FindIndex((HTKKCMCTableInfo x) => x.TableName == tableName && x.TableID == tables[tables.Count - 1].TableID);
			for (int num3 = num2; num3 > num; num3--)
			{
				if (num3 < sInfo.Tables.Count - 1 && sInfo.Tables[num3] != null)
				{
					int tableTemplateRowStart = sInfo.Tables[num3].Location.TableTemplateRowStart;
					RemoveTableF8New(sView, sInfo, num3, setActiveCell: false);
				}
			}
			if (sInfo.Tables[num].RowsValues.Count > 1)
			{
				int dataLocationTemplateRowStart = sInfo.Tables[num].Location.DataLocationTemplateRowStart;
				RemoveDataRows(sView, num, dataLocationTemplateRowStart - 1, sInfo.Tables[num].RowsValues.Count - 1, sInfo);
			}
			if (sInfo.Tables[num].RowsValues.Count > 0)
			{
				int dataLocationTemplateRowStart = sInfo.Tables[num].Location.DataLocationTemplateRowStart;
				foreach (HTKKCMCTableCellInfo cell in sInfo.Tables[num].RowsInfo.Cells)
				{
					sView.Cells[dataLocationTemplateRowStart - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
				}
			}
			if (sInfo.Tables[num].Header == null || sInfo.Tables[num].Header.StaticCells == null)
			{
				return;
			}
			int rowIndex = sInfo.Tables[num].Header.StaticCells[0].CellTemplate.RowIndex;
			foreach (HTKKCMCCellInfo staticCell in sInfo.Tables[num].Header.StaticCells)
			{
				sView.Cells[rowIndex - 1, staticCell.CellTemplate.ColIndex - 1].Value = staticCell.DefaultValue;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CopyTableToTableGroup_F7_F5(SheetView sView, HTKKCMCSheetInfo sInfo, string fromTableName, string toTableName, Dictionary<int, int> dicGroup, Dictionary<int, int> dicInfo, bool usingCellType = false)
	{
		try
		{
			List<HTKKCMCTableInfo> list = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == fromTableName).ToList();
			List<HTKKCMCTableInfo> list2 = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == toTableName).ToList();
			int num = sInfo.Tables.FindIndex((HTKKCMCTableInfo x) => x.TableName == toTableName);
			int dataLocationTemplateRowStart = list2[0].Location.DataLocationTemplateRowStart;
			if (list2[0].RowsValues.Count > 1)
			{
				RemoveDataRows(sView, num, dataLocationTemplateRowStart - 1, list2[0].RowsValues.Count - 1, sInfo);
			}
			foreach (HTKKCMCTableCellInfo cell in list2[0].RowsInfo.Cells)
			{
				sView.Cells[dataLocationTemplateRowStart - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
			}
			Dictionary<string, int> checkExists = new Dictionary<string, int>();
			foreach (HTKKCMCTableInfo item in list)
			{
				CopyTableToTable2(sView, sInfo, item, num, dicInfo, deleteDataOld: false, ref checkExists, usingCellType);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ThayDoiCQTTheoCQTOKyTinhThueDangBang(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKPeriodInfo PeriodInfo, int tableIndex, int colTenCQTCuc, int colMaCQTCuc, int colTenCQTChiCuc, int colMaCQTChiCuc, int iRow = -1, bool isLock = true)
	{
		try
		{
			if (string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc))
			{
				return;
			}
			string tenCoQuanThueCapCuc = HTKKWinDictionaryDiaDanh.GetTenCoQuanThueCapCuc(PeriodInfo.maCQTCapCuc, isGetFullName: false);
			if (string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy) || string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc) || !(PeriodInfo.maCQTQuanLy != PeriodInfo.maCQTCapCuc))
			{
				return;
			}
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (tableIndex >= 0 && (tableIndex < 0 || !(table.TableName == sInfo.Tables[tableIndex].TableName)))
				{
					continue;
				}
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				if (iRow < 0)
				{
					for (int i = 0; i < table.RowsValues.Count; i++)
					{
						sView.Cells[dataLocationTemplateRowStart + i - 1, colTenCQTCuc - 1].Text = tenCoQuanThueCapCuc;
						sView.Cells[dataLocationTemplateRowStart + i - 1, colMaCQTCuc - 1].Text = PeriodInfo.maCQTCapCuc;
						sView.Cells[dataLocationTemplateRowStart + i - 1, colTenCQTChiCuc - 1].Text = PeriodInfo.tenCQTQuanLy;
						sView.Cells[dataLocationTemplateRowStart + i - 1, colMaCQTChiCuc - 1].Text = PeriodInfo.maCQTQuanLy;
						LockUnLockCol(sView, dataLocationTemplateRowStart + i, colTenCQTCuc, isLock, deleteData: false);
						LockUnLockCol(sView, dataLocationTemplateRowStart + i, colTenCQTChiCuc, isLock, deleteData: false);
					}
				}
				else
				{
					sView.Cells[iRow - 1, colTenCQTCuc - 1].Text = tenCoQuanThueCapCuc;
					sView.Cells[iRow - 1, colMaCQTCuc - 1].Text = PeriodInfo.maCQTCapCuc;
					sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Text = PeriodInfo.tenCQTQuanLy;
					sView.Cells[iRow - 1, colMaCQTChiCuc - 1].Text = PeriodInfo.maCQTQuanLy;
					LockUnLockCol(sView, iRow, colTenCQTCuc, isLock, deleteData: false);
					LockUnLockCol(sView, iRow, colTenCQTChiCuc, isLock, deleteData: false);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetLaiDongCon(SheetView sView, HTKKCMCSheetInfo sInfo, int tableIndex, int iRow, int colHeader, int colChilden, int indexRowHeader = 0)
	{
		try
		{
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (table == null || table.Header == null || table.Header.StaticCells == null || table.Header.StaticCells.Count == 0 || (tableIndex >= 0 && (tableIndex < 0 || !(table.TableName == sInfo.Tables[tableIndex].TableName) || !(table.TableID == sInfo.Tables[tableIndex].TableID))))
				{
					continue;
				}
				int rowIndex = table.Header.StaticCells[indexRowHeader].CellTemplate.RowIndex;
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				if (iRow < 0)
				{
					for (int i = 0; i < table.RowsValues.Count; i++)
					{
						sView.Cells[dataLocationTemplateRowStart + i - 1, colChilden - 1].Value = sView.Cells[rowIndex - 1, colHeader - 1].Value;
					}
				}
				else
				{
					sView.Cells[iRow - 1, colChilden - 1].Value = sView.Cells[rowIndex - 1, colHeader - 1].Value;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ThayDoiCQTTheoTinhDynamic(SheetView sView, string maTinh, int iRow, int colTenCQTCuc, int colMaCQTCuc, int colTenCQTChiCuc, int colMaCQTChiCuc, bool isLock = true)
	{
		try
		{
			if (string.IsNullOrEmpty(maTinh))
			{
				sView.Cells[iRow - 1, colTenCQTCuc - 1].Value = "";
				sView.Cells[iRow - 1, colMaCQTCuc - 1].Value = "";
				sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Value = "";
				sView.Cells[iRow - 1, colMaCQTChiCuc - 1].Value = "";
			}
			else if (HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.ContainsValue(maTinh + "00"))
			{
				string maCQTCapCuc = maTinh + "00";
				Dictionary<string, string> coQuanThueQuanLybyMaCoQuanThueCapCuc = HTKKWinDictionaryDiaDanh.getCoQuanThueQuanLybyMaCoQuanThueCapCuc(maCQTCapCuc);
				KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maCQTCapCuc);
				string tenCoQuanThueCapCuc = HTKKWinDictionaryDiaDanh.GetTenCoQuanThueCapCuc(maCQTCapCuc, isGetFullName: false);
				string text = maCQTCapCuc;
				string text2 = coQuanThueQuanLybyMaCoQuanThueCapCuc[text];
				int num = keyValuePair.Key.IndexOf(" - ");
				sView.Cells[iRow - 1, colTenCQTCuc - 1].Text = tenCoQuanThueCapCuc;
				sView.Cells[iRow - 1, colMaCQTCuc - 1].Text = maCQTCapCuc;
				InsertCombobox(sView, iRow, colTenCQTChiCuc, coQuanThueQuanLybyMaCoQuanThueCapCuc);
				sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Text = text2;
				sView.Cells[iRow - 1, colMaCQTChiCuc - 1].Text = text;
				LockUnLockCol(sView, iRow, colTenCQTCuc, isLock, deleteData: false);
				LockUnLockCol(sView, iRow, colTenCQTChiCuc, isLock, deleteData: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ThayDoiCQTTheoTinhDynamic2(SheetView sView, string tenCQTCapCuc, string maCQTCapCuc, string tenCQTChiCuc, string maCQTChiCuc, int iRow, int colTenCQTCuc, int colMaCQTCuc, int colTenCQTChiCuc, int colMaCQTChiCuc)
	{
		try
		{
			if (string.IsNullOrEmpty(maCQTCapCuc))
			{
				sView.Cells[iRow - 1, colTenCQTCuc - 1].Value = "";
				sView.Cells[iRow - 1, colMaCQTCuc - 1].Value = "";
				sView.Cells[iRow - 1, colTenCQTCuc - 1].Locked = false;
				sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Value = "";
				sView.Cells[iRow - 1, colMaCQTChiCuc - 1].Value = "";
				sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Locked = false;
				return;
			}
			sView.Cells[iRow - 1, colTenCQTCuc - 1].Value = "";
			sView.Cells[iRow - 1, colMaCQTCuc - 1].Value = "";
			sView.Cells[iRow - 1, colTenCQTCuc - 1].Locked = false;
			sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Value = "";
			sView.Cells[iRow - 1, colMaCQTChiCuc - 1].Value = "";
			sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Locked = false;
			if (string.IsNullOrEmpty(maCQTCapCuc))
			{
				return;
			}
			CQTDiaBan cQTDiaBan = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == maCQTCapCuc);
			CQTDiaBan cQTDiaBan2 = null;
			if (!string.IsNullOrEmpty(maCQTChiCuc))
			{
				cQTDiaBan2 = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == maCQTChiCuc);
			}
			Dictionary<string, string> dic = HTKKWinDictionaryDiaDanh._cqtDiaBan.Where((CQTDiaBan x) => x.MaCQT == maCQTCapCuc).ToDictionary((CQTDiaBan x) => x.MaCQT, (CQTDiaBan x) => x.TenCQT);
			tenCQTCapCuc = cQTDiaBan.TenCQT;
			tenCQTChiCuc = ((cQTDiaBan2 == null) ? "" : cQTDiaBan2.TenCQT);
			sView.Cells[iRow - 1, colTenCQTCuc - 1].Value = tenCQTCapCuc;
			sView.Cells[iRow - 1, colMaCQTCuc - 1].Value = maCQTCapCuc;
			InsertCombobox(sView, iRow, colTenCQTChiCuc, dic);
			sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Value = tenCQTChiCuc;
			sView.Cells[iRow - 1, colMaCQTChiCuc - 1].Value = maCQTChiCuc;
			LockUnLockCol(sView, iRow, colTenCQTCuc, !string.IsNullOrEmpty(tenCQTCapCuc), deleteData: false);
			LockUnLockCol(sView, iRow, colTenCQTChiCuc, !string.IsNullOrEmpty(tenCQTChiCuc), deleteData: false);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ThayDoiCQTTheoTinhStatic(SheetView sView, string maTinh, int[] cellTenCQTCuc, int[] cellMaCQTCuc, int[] cellTenCQTChiCuc, int[] cellMaCQTChiCuc, bool isLock = true)
	{
		try
		{
			if (string.IsNullOrEmpty(maTinh))
			{
				sView.Cells[cellTenCQTCuc[1], cellTenCQTCuc[0]].Value = "";
				sView.Cells[cellMaCQTCuc[1], cellMaCQTCuc[0]].Value = "";
				sView.Cells[cellTenCQTChiCuc[1], cellTenCQTChiCuc[0]].Value = "";
				sView.Cells[cellMaCQTChiCuc[1], cellMaCQTChiCuc[0]].Value = "";
			}
			else if (HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.ContainsValue(maTinh + "00"))
			{
				string maCQTCapCuc = maTinh + "00";
				Dictionary<string, string> coQuanThueQuanLybyMaCoQuanThueCapCuc = HTKKWinDictionaryDiaDanh.getCoQuanThueQuanLybyMaCoQuanThueCapCuc(maCQTCapCuc);
				KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maCQTCapCuc);
				string text = maCQTCapCuc;
				string text2 = coQuanThueQuanLybyMaCoQuanThueCapCuc[text];
				int num = keyValuePair.Key.IndexOf(" - ");
				string tenCoQuanThueCapCuc = HTKKWinDictionaryDiaDanh.GetTenCoQuanThueCapCuc(maCQTCapCuc, isGetFullName: false);
				sView.Cells[cellTenCQTCuc[1], cellTenCQTCuc[0]].Text = tenCoQuanThueCapCuc;
				sView.Cells[cellMaCQTCuc[1], cellMaCQTCuc[0]].Text = maCQTCapCuc;
				InsertCombobox(sView, cellTenCQTChiCuc[1] + 1, cellTenCQTChiCuc[0] + 1, coQuanThueQuanLybyMaCoQuanThueCapCuc);
				sView.Cells[cellTenCQTChiCuc[1], cellTenCQTChiCuc[0]].Text = text2;
				sView.Cells[cellMaCQTChiCuc[1], cellMaCQTChiCuc[0]].Text = text;
				LockUnLockCol(sView, cellTenCQTCuc[1] + 1, cellTenCQTCuc[0] + 1, isLock, deleteData: false);
				LockUnLockCol(sView, cellTenCQTChiCuc[1], cellTenCQTChiCuc[0], isLock, deleteData: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string GetNameNopTheoTinhCQT(HTKKCMCToKhaiInfo toKhaiInfo, HTKKPeriodInfo periodInfo)
	{
		try
		{
			if (!string.IsNullOrEmpty(toKhaiInfo.NopTheoTinhCQT))
			{
				if (toKhaiInfo.NopTheoTinhCQT == "1")
				{
					if (!string.IsNullOrEmpty(periodInfo.MaTinh))
					{
						return "TINH" + periodInfo.MaTinh;
					}
					if (!string.IsNullOrEmpty(periodInfo.maCQTQuanLy))
					{
						return "CQTQL" + periodInfo.maCQTQuanLy;
					}
				}
				if (toKhaiInfo.NopTheoTinhCQT == "2" && !string.IsNullOrEmpty(periodInfo.maCQTQuanLy))
				{
					return "CQTQL" + periodInfo.maCQTQuanLy;
				}
				if (toKhaiInfo.NopTheoTinhCQT == "3" && !string.IsNullOrEmpty(periodInfo.maCQTQuanLy))
				{
					return "CQTQL" + periodInfo.maCQTQuanLy;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return "";
	}

	public static Dictionary<string, double> SumValueInTableByKey(SheetView sView, HTKKCMCSheetInfo sInfo, string tableName, int colKey, int colValue)
	{
		Dictionary<string, double> dictionary = new Dictionary<string, double>();
		try
		{
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (!(table.TableName == tableName))
				{
					continue;
				}
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < table.RowsValues.Count; i++)
				{
					string text = sView.Cells[dataLocationTemplateRowStart + i - 1, colKey - 1].Text.Trim();
					double num = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + i - 1, colValue - 1].Value);
					if (!string.IsNullOrEmpty(text))
					{
						if (!dictionary.ContainsKey(text))
						{
							dictionary.Add(text, num);
						}
						else
						{
							dictionary[text] += num;
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return dictionary;
	}

	public static void SetAllRowSameRowFirst(SheetView sView, HTKKCMCTableInfo table, List<int> cols)
	{
		try
		{
			if (table.RowsValues.Count == 1)
			{
				return;
			}
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			TextCellType textCellType = new TextCellType();
			for (int i = 1; i < table.RowsValues.Count; i++)
			{
				foreach (int col in cols)
				{
					ICellType cellType = sView.Cells[dataLocationTemplateRowStart + i - 1, col - 1].CellType;
					object value = sView.Cells[dataLocationTemplateRowStart - 1, col - 1].Value;
					if ((object)cellType.GetType() != textCellType.GetType())
					{
						sView.Cells[dataLocationTemplateRowStart + i - 1, col - 1].Value = value;
					}
					else
					{
						sView.Cells[dataLocationTemplateRowStart + i - 1, col - 1].Text = value.ToString().Trim();
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static HTKKCMCCellInfo GetCellStaticByRowCol(HTKKCMCSheetInfo sInfo, int row, int col)
	{
		return sInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == row && x.CellTemplate.ColIndex == col);
	}

	public static string GetNgayTaiChinhDen(string ngayTCTu)
	{
		string result = "";
		switch (ngayTCTu)
		{
		case "01/01":
			result = "31/12";
			break;
		case "01/04":
			result = "31/03";
			break;
		case "01/07":
			result = "30/06";
			break;
		case "01/10":
			result = "30/09";
			break;
		}
		return result;
	}

	public static void UpdateCacHangSauGiongHangTruocDo(SheetView sView, HTKKCMCTableInfo table, List<int> cols, bool chiGiongHangLock)
	{
		try
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 1; i < table.RowsValues.Count; i++)
			{
				foreach (int col in cols)
				{
					if (chiGiongHangLock)
					{
						if (sView.Cells[dataLocationTemplateRowStart + i - 1, col - 1].Locked)
						{
							sView.Cells[dataLocationTemplateRowStart + i - 1, col - 1].Text = sView.Cells[dataLocationTemplateRowStart + i - 2, col - 1].Text;
						}
					}
					else
					{
						sView.Cells[dataLocationTemplateRowStart + i - 1, col - 1].Text = sView.Cells[dataLocationTemplateRowStart + i - 2, col - 1].Text;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetAllRowSame(SheetView sView, HTKKCMCTableInfo table, int col, object value)
	{
		try
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 1; i < table.RowsValues.Count; i++)
			{
				sView.Cells[dataLocationTemplateRowStart + i - 1, col - 1].Value = value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool CheckInListEmpty(List<string> list)
	{
		try
		{
			if (list == null || list.Count == 0)
			{
				return true;
			}
			int num = 0;
			foreach (string item in list)
			{
				if (!string.IsNullOrEmpty(item))
				{
					num++;
				}
			}
			if (num == 0)
			{
				return true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return false;
	}

	public static List<double> CalculateColSumInTable(SheetView sView, HTKKCMCRowInfo rowInfo, List<HTKKCMCRowValue> rowValues, List<int> colSums, int rowSum, bool isShow)
	{
		List<double> list = new List<double>();
		try
		{
			for (int i = 0; i < colSums.Count; i++)
			{
				list.Add(0.0);
			}
			List<int> list2 = new List<int>();
			foreach (int col in colSums)
			{
				List<HTKKCMCTableCellInfo> cells = rowInfo.Cells;
				Predicate<HTKKCMCTableCellInfo> match = (HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == col;
				int num = cells.FindIndex(match);
				if (num >= 0)
				{
					list2.Add(num);
				}
			}
			for (int num2 = 0; num2 < rowValues.Count; num2++)
			{
				int num3 = 0;
				foreach (int item in list2)
				{
					double num4 = (string.IsNullOrEmpty(rowValues[num2].RowValue[item]) ? 0.0 : Convert.ToDouble(rowValues[num2].RowValue[item]));
					list[num3] += num4;
					num3++;
				}
			}
			if (isShow)
			{
				int num5 = 0;
				foreach (int colSum in colSums)
				{
					sView.Cells[rowSum - 1, colSum - 1].Value = list[num5];
					num5++;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	public static List<double> CalculateColSumOnSview(SheetView sView, int rowStart, List<HTKKCMCRowValue> rowValues, List<int> colSums)
	{
		List<double> list = new List<double>();
		try
		{
			for (int i = 0; i < colSums.Count; i++)
			{
				list.Add(0.0);
			}
			for (int j = 0; j < rowValues.Count; j++)
			{
				int num = 0;
				foreach (int colSum in colSums)
				{
					double num2 = Convert.ToDouble(sView.Cells[rowStart + j - 1, colSum - 1].Value);
					list[num] += num2;
					num++;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	public static int GetRowIndexInSViewByCompare2(SheetView sView, int rowStart, int colCheck, string valueCompare)
	{
		int result = -1;
		try
		{
			for (int num = sView.Rows.Count - 1; num >= rowStart; num--)
			{
				if (sView.Cells[num - 1, colCheck - 1].Value != null && sView.Cells[num - 1, colCheck - 1].Value.ToString() == valueCompare)
				{
					result = num;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static void InsertRowOnSView(SheetView sView, int frRow, int frCol, int toRow, int toCol, int colCount, int rowCount, int row, int maxCopyRange)
	{
		try
		{
			sView.Rows.Add(toRow, rowCount);
			if (row < maxCopyRange)
			{
				sView.CopyRange(frRow - 1, frCol, toRow, toCol, rowCount, colCount, dataOnly: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void RemoveRowOnSView(SheetView sView, int frRow, int rowCount)
	{
		try
		{
			if (rowCount > 0)
			{
				sView.Rows.Remove(frRow, rowCount);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void BindDataToSviewFromTable(SheetView sView, HTKKCMCTableInfo table, List<int> colsFormualar, int colBoderRight)
	{
		try
		{
			HTKKCMCRowInfo rowsInfo = table.RowsInfo;
			List<HTKKCMCRowValue> rowsValues = table.RowsValues;
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			List<int> list = new List<int>();
			foreach (HTKKCMCTableCellInfo cell in table.RowsInfo.Cells)
			{
				if (!list.Contains(cell.CellTemplate.ColIndex))
				{
					list.Add(cell.CellTemplate.ColIndex);
				}
			}
			list.AddRange(colsFormualar);
			Dictionary<string, string> dicTeamplateFormula = CreateTeamplateFormulaAllCols(sView, rowsInfo, dataLocationTemplateRowStart, list);
			if (table.Header.StaticCells.Count > 0)
			{
				for (int i = 0; i < table.Header.StaticCells.Count; i++)
				{
					fpSpBindDataCell(sView, table.Header.StaticCells[i].CellTemplate.RowIndex, table.Header.StaticCells[i].CellTemplate.ColIndex, table.Header.StaticCells[i].Controltype, table.Header.StaticCells[i].Value);
				}
			}
			InsertTableRowsSpreadFormLargeData(sView, table, colsFormualar, colBoderRight, dicTeamplateFormula);
			BindDataLargeSetValueByCols(sView, table);
			if (table.Footer.StaticCells.Count > 0)
			{
				for (int i = 0; i < table.Footer.StaticCells.Count; i++)
				{
					fpSpBindDataCell(sView, table.Footer.StaticCells[i].CellTemplate.RowIndex, table.Footer.StaticCells[i].CellTemplate.ColIndex, table.Footer.StaticCells[i].Controltype, table.Footer.StaticCells[i].Value);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void UpdateFromSViewToSInfo(SheetView sView, int row, int col, HTKKCMCRowInfo rowInfo, HTKKCMCTableInfo table)
	{
		try
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			int num = row - dataLocationTemplateRowStart;
			if (num < 0)
			{
				return;
			}
			int num2 = 0;
			if (col >= 0)
			{
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == col);
				num2 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == col);
				if (hTKKCMCTableCellInfo == null || num2 < 0)
				{
					return;
				}
				object obj = sView.Cells[row - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Value;
				if (hTKKCMCTableCellInfo.Controltype == "16")
				{
					if (obj != null && !string.IsNullOrEmpty(obj.ToString()) && obj.ToString().EndsWith(".0"))
					{
						obj = obj.ToString().Substring(0, obj.ToString().Length - 2);
					}
					table.RowsValues[num].RowValue[num2] = ((obj == null) ? "0" : obj.ToString());
				}
				else if (hTKKCMCTableCellInfo.Controltype == "2")
				{
					table.RowsValues[num].RowValue[num2] = ((obj != null && Convert.ToBoolean(obj)) ? "1" : "0");
				}
				else
				{
					table.RowsValues[num].RowValue[num2] = ((obj != null) ? obj.ToString() : "");
				}
				return;
			}
			foreach (HTKKCMCTableCellInfo cell in table.RowsInfo.Cells)
			{
				object obj = sView.Cells[row - 1, cell.CellTemplate.ColIndex - 1].Value;
				if (cell.Controltype == "16")
				{
					if (obj != null && !string.IsNullOrEmpty(obj.ToString()) && obj.ToString().EndsWith(".0"))
					{
						obj = obj.ToString().Substring(0, obj.ToString().Length - 2);
					}
					table.RowsValues[num].RowValue[num2] = ((obj == null) ? "0" : obj.ToString());
				}
				else if (cell.Controltype == "2")
				{
					table.RowsValues[num].RowValue[num2] = ((obj != null && Convert.ToBoolean(obj)) ? "1" : "0");
				}
				else
				{
					table.RowsValues[num].RowValue[num2] = ((obj != null) ? obj.ToString() : "");
				}
				num2++;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SumColsInTable(SheetView sView, int row, List<int> cols, HTKKCMCTableInfo table)
	{
		try
		{
			foreach (int col in cols)
			{
				List<HTKKCMCTableCellInfo> cells = table.RowsInfo.Cells;
				Predicate<HTKKCMCTableCellInfo> match = (HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == col;
				int index = cells.FindIndex(match);
				if (index >= 0)
				{
					double num = table.RowsValues.Select((HTKKCMCRowValue x) => x).Sum((HTKKCMCRowValue x) => Convert.ToDouble(x.RowValue[index]));
					sView.Cells[row - 1, col - 1].Value = num;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void BindDataCQTInTable(SheetView sView, HTKKCMCTableInfo table, int colMaCQT, int colMaCucThue, int colTenCucThue)
	{
		try
		{
			int num = table.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				string text = sView.Cells[num + i, colMaCQT].Text;
				string maCucThue = "";
				string text2 = "";
				if (text.Length == 5)
				{
					maCucThue = text.Substring(0, 3) + "00";
				}
				KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maCucThue);
				int num2 = keyValuePair.Key.IndexOf(" - ");
				text2 = ((num2 > 0) ? keyValuePair.Key.Substring(num2 + 3, keyValuePair.Key.Length - num2 - 3) : keyValuePair.Key);
				sView.Cells[num + i, colMaCucThue].Text = maCucThue;
				sView.Cells[num + i, colTenCucThue].Text = text2;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void BindDataCQTInTableF7(SheetView sView, HTKKCMCSheetInfo sheetPLInfo, string pathTable, string rowPath, List<int> listCols, List<string> listHeaderPath, List<string> listFooterPath)
	{
		try
		{
			int column = 0;
			int column2 = 0;
			int column3 = 0;
			int row = 0;
			int row2 = 0;
			int row3 = 0;
			string text = "";
			string text2 = "";
			string text3 = "";
			if (sheetPLInfo == null)
			{
				return;
			}
			foreach (HTKKCMCTableInfo table in sheetPLInfo.Tables)
			{
				if (listHeaderPath != null && listHeaderPath.Count == 3)
				{
					text = listHeaderPath[0];
					text2 = listHeaderPath[1];
					text3 = listHeaderPath[2];
				}
				string text4;
				string text5;
				if (table.Header != null && table.Header.StaticCells.Count > 0)
				{
					foreach (HTKKCMCCellInfo staticCell in table.Header.StaticCells)
					{
						if (staticCell.Path == text)
						{
							column = staticCell.CellTemplate.ColIndex - 1;
							row = staticCell.CellTemplate.RowIndex - 1;
						}
						if (staticCell.Path == text2)
						{
							column2 = staticCell.CellTemplate.ColIndex - 1;
							row2 = staticCell.CellTemplate.RowIndex - 1;
						}
						if (staticCell.Path == text3)
						{
							column3 = staticCell.CellTemplate.ColIndex - 1;
							row3 = staticCell.CellTemplate.RowIndex - 1;
						}
					}
					text4 = sView.Cells[row, column].Text;
					string maCucThue = "";
					text5 = "";
					if (text4.Length == 5)
					{
						maCucThue = text4.Substring(0, 3) + "00";
					}
					if (!string.IsNullOrEmpty(maCucThue))
					{
						KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maCucThue);
						int num = keyValuePair.Key.IndexOf(" - ");
						text5 = ((num > 0) ? keyValuePair.Key.Substring(num + 3, keyValuePair.Key.Length - num - 3) : keyValuePair.Key);
						sView.Cells[row2, column2].Text = maCucThue;
						sView.Cells[row3, column3].Text = text5;
					}
				}
				if (table.TablePath == pathTable)
				{
					int num2 = table.Location.DataLocationTemplateRowStart - 1;
					if (table.RowsInfo.RowPath == rowPath)
					{
						for (int num3 = 0; num3 < table.RowsValues.Count; num3++)
						{
							text4 = sView.Cells[num2 + num3, column].Text;
							string maCucThue2 = "";
							text5 = "";
							if (text4.Length == 5)
							{
								maCucThue2 = text4.Substring(0, 3) + "00";
							}
							if (!string.IsNullOrEmpty(maCucThue2))
							{
								KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maCucThue2);
								int num = keyValuePair.Key.IndexOf(" - ");
								text5 = ((num > 0) ? keyValuePair.Key.Substring(num + 3, keyValuePair.Key.Length - num - 3) : keyValuePair.Key);
								sView.Cells[num2 + num3, column2].Text = maCucThue2;
								sView.Cells[num2 + num3, column3].Text = text5;
							}
						}
					}
				}
				if (listFooterPath != null && listFooterPath.Count == 3)
				{
					text = listFooterPath[0];
					text2 = listFooterPath[1];
					text3 = listFooterPath[2];
				}
				if (table.Footer == null || table.Footer.StaticCells.Count <= 0)
				{
					continue;
				}
				foreach (HTKKCMCCellInfo staticCell2 in table.Footer.StaticCells)
				{
					if (staticCell2.Path == text)
					{
						column = staticCell2.CellTemplate.ColIndex;
						row = staticCell2.CellTemplate.RowIndex;
					}
					if (staticCell2.Path == text2)
					{
						column2 = staticCell2.CellTemplate.ColIndex;
						row2 = staticCell2.CellTemplate.RowIndex;
					}
					if (staticCell2.Path == text3)
					{
						column3 = staticCell2.CellTemplate.ColIndex;
						row3 = staticCell2.CellTemplate.RowIndex;
					}
				}
				text4 = sView.Cells[row, column].Text;
				string maCucThue3 = "";
				text5 = "";
				if (text4.Length == 5)
				{
					maCucThue3 = text4.Substring(0, 3) + "00";
				}
				if (!string.IsNullOrEmpty(maCucThue3))
				{
					KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maCucThue3);
					int num = keyValuePair.Key.IndexOf(" - ");
					text5 = ((num > 0) ? keyValuePair.Key.Substring(num + 3, keyValuePair.Key.Length - num - 3) : keyValuePair.Key);
					sView.Cells[row2, column2].Text = maCucThue3;
					sView.Cells[row3, column3].Text = text5;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetCellTypeTheoKieuKy(SheetView sView, string typeKeKhai, int iRow, int iCol)
	{
		try
		{
			MaskCellType maskCellType = new MaskCellType();
			maskCellType.MaskChar = '.';
			switch (typeKeKhai)
			{
			case "Y":
				maskCellType.Mask = "0000";
				break;
			case "Q":
				maskCellType.Mask = "0/0000";
				break;
			case "M":
				maskCellType.Mask = "00/0000";
				break;
			case "D":
				maskCellType.Mask = "00/00/0000";
				break;
			default:
				maskCellType.Mask = "0/0000";
				break;
			}
			sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetCellTypeDataLarge(SheetView sView, HTKKCMCTableInfo table, List<int> colsFormula, int colBoderRight, Dictionary<string, string> dicTeamplateFormula)
	{
		try
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			int num = table.RowsValues.Count - 1;
			HTKKCMCRowInfo rowsInfo = table.RowsInfo;
			int num2 = dataLocationTemplateRowStart + num;
			CheckBoxCellType checkBoxCellType = new CheckBoxCellType();
			MaskCellType maskCellType = new MaskCellType();
			NumberCellType numberCellType = new NumberCellType();
			CurrencyCellType currencyCellType = new CurrencyCellType();
			GeneralCellType generalCellType = new GeneralCellType();
			for (int i = 0; i <= rowsInfo.Cells[0].CellTemplate.ColIndex - 1; i++)
			{
				sView.Cells[dataLocationTemplateRowStart, i, num2 - 1, i].Border = sView.Cells[dataLocationTemplateRowStart - 1, i].Border;
				sView.Cells[dataLocationTemplateRowStart, i, num2 - 1, i].BackColor = sView.Cells[dataLocationTemplateRowStart - 1, i].BackColor;
			}
			foreach (HTKKCMCTableCellInfo cell2 in rowsInfo.Cells)
			{
				Cell cell = sView.Cells[dataLocationTemplateRowStart - 1, cell2.CellTemplate.ColIndex - 1];
				ICellType cellType = cell.CellType;
				sView.Cells[dataLocationTemplateRowStart, cell2.CellTemplate.ColIndex - 1, num2 - 1, cell2.CellTemplate.ColIndex - 1].CellType = cell.CellType;
				sView.Cells[dataLocationTemplateRowStart, cell2.CellTemplate.ColIndex - 1, num2 - 1, cell2.CellTemplate.ColIndex - 1].Border = cell.Border;
				sView.Cells[dataLocationTemplateRowStart, cell2.CellTemplate.ColIndex - 1, num2 - 1, cell2.CellTemplate.ColIndex - 1].CellPadding = cell.CellPadding;
				sView.Cells[dataLocationTemplateRowStart, cell2.CellTemplate.ColIndex - 1, num2 - 1, cell2.CellTemplate.ColIndex - 1].BackColor = cell.BackColor;
				sView.Cells[dataLocationTemplateRowStart, cell2.CellTemplate.ColIndex - 1, num2 - 1, cell2.CellTemplate.ColIndex - 1].Font = cell.Font;
				sView.Cells[dataLocationTemplateRowStart, cell2.CellTemplate.ColIndex - 1, num2 - 1, cell2.CellTemplate.ColIndex - 1].ForeColor = cell.ForeColor;
				sView.Cells[dataLocationTemplateRowStart, cell2.CellTemplate.ColIndex - 1, num2 - 1, cell2.CellTemplate.ColIndex - 1].Locked = cell.Locked;
				sView.Cells[dataLocationTemplateRowStart, cell2.CellTemplate.ColIndex - 1, num2 - 1, cell2.CellTemplate.ColIndex - 1].Formatter = cell.Formatter;
				if (cellType != null && ((object)cellType.GetType() == numberCellType.GetType() || (object)cellType.GetType() == currencyCellType.GetType()))
				{
					sView.Cells[dataLocationTemplateRowStart, cell2.CellTemplate.ColIndex - 1, num2 - 1, cell2.CellTemplate.ColIndex - 1].HorizontalAlignment = CellHorizontalAlignment.Right;
				}
				else if (cellType != null && ((object)cellType.GetType() == maskCellType.GetType() || (object)cellType.GetType() == checkBoxCellType.GetType()))
				{
					sView.Cells[dataLocationTemplateRowStart, cell2.CellTemplate.ColIndex - 1, num2 - 1, cell2.CellTemplate.ColIndex - 1].HorizontalAlignment = CellHorizontalAlignment.Center;
				}
				else
				{
					sView.Cells[dataLocationTemplateRowStart, cell2.CellTemplate.ColIndex - 1, num2 - 1, cell2.CellTemplate.ColIndex - 1].HorizontalAlignment = CellHorizontalAlignment.Left;
				}
			}
			sView.Cells[dataLocationTemplateRowStart, colBoderRight - 1, num2 - 1, colBoderRight - 1].Border = sView.Cells[dataLocationTemplateRowStart - 1, colBoderRight - 1].Border;
			sView.Cells[dataLocationTemplateRowStart, colBoderRight - 1, num2 - 1, colBoderRight - 1].BackColor = sView.Cells[dataLocationTemplateRowStart - 1, colBoderRight - 1].BackColor;
			for (int j = 1; j < table.RowsValues.Count; j++)
			{
				num2 = dataLocationTemplateRowStart + j;
				foreach (HTKKCMCTableCellInfo cell3 in rowsInfo.Cells)
				{
					Cell cell = sView.Cells[dataLocationTemplateRowStart - 1, cell3.CellTemplate.ColIndex - 1];
					if (!string.IsNullOrEmpty(cell.Formula) && dicTeamplateFormula.ContainsKey(cell3.CellTemplate.ColIndex.ToString()))
					{
						string formula = dicTeamplateFormula[cell3.CellTemplate.ColIndex.ToString()];
						string formula2 = SetFormulaNextRow(formula, num2);
						sView.Cells[num2 - 1, cell3.CellTemplate.ColIndex - 1].Formula = formula2;
					}
				}
				if (colsFormula == null)
				{
					continue;
				}
				foreach (int item in colsFormula)
				{
					if (dicTeamplateFormula.ContainsKey(item.ToString()))
					{
						string formula = dicTeamplateFormula[item.ToString()];
						string formula2 = SetFormulaNextRow(formula, num2);
						sView.Cells[num2 - 1, item - 1].Formula = formula2;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static List<HTKKCMCRowValue> GetRowIndexInTableByCompare(HTKKCMCTableInfo table, int indexCol, string valueCompare)
	{
		List<HTKKCMCRowValue> result = new List<HTKKCMCRowValue>();
		if (table == null)
		{
			return result;
		}
		return table.RowsValues.Where((HTKKCMCRowValue x) => x.RowValue[indexCol] == valueCompare).ToList();
	}

	public static Dictionary<string, string> CreateTeamplateFormulaAllCols(SheetView sView, HTKKCMCRowInfo rowInfo, int iRow, List<int> cols)
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		try
		{
			foreach (int col in cols)
			{
				string formula = sView.Cells[iRow - 1, col - 1].Formula;
				if (!string.IsNullOrEmpty(formula))
				{
					string value = FormulaToTemplate(formula, iRow);
					if (!dictionary.ContainsKey(col.ToString()) && !string.IsNullOrEmpty(value))
					{
						dictionary.Add(col.ToString(), value);
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return dictionary;
		}
	}

	public static string FormulaToTemplate(string formula, int iRow)
	{
		string text = "";
		try
		{
			string[] separator = new string[1] { iRow.ToString() };
			string[] array = formula.Split(separator, StringSplitOptions.None);
			string text2 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
			int num = 0;
			string[] array2 = array;
			foreach (string text3 in array2)
			{
				if (text3.Length != 0)
				{
					string value = text3.Substring(text3.Length - 1, 1);
					text = ((text2.LastIndexOf(value) >= 0) ? (text + text3 + "###") : ((num != array.Length - 1) ? (text + text3 + iRow) : (text + text3)));
					num++;
				}
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return text;
		}
	}

	public static string SetFormulaNextRow(string formula, int newRow)
	{
		try
		{
			if (string.IsNullOrEmpty(formula))
			{
				return "";
			}
			return formula.Replace("###", newRow.ToString());
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return formula;
		}
	}

	public static string GetNumberInString(string str)
	{
		return Regex.Match(str, "\\d+").Value;
	}

	public static void InsertTableRowsSpreadFormLargeData(SheetView sView, HTKKCMCTableInfo table, List<int> colsFormula, int colBoderRight, Dictionary<string, string> dicTeamplateFormula)
	{
		if (table.RowsValues.Count > 1)
		{
			if (((Control)sView.FpSpread).InvokeRequired)
			{
				sView.AddRows(table.Location.DataLocationTemplateRowStart, table.RowsValues.Count - 1);
				SetCellTypeDataLarge(sView, table, colsFormula, colBoderRight, dicTeamplateFormula);
			}
			else
			{
				sView.AddRows(table.Location.DataLocationTemplateRowStart, table.RowsValues.Count - 1);
				SetCellTypeDataLarge(sView, table, colsFormula, colBoderRight, dicTeamplateFormula);
			}
		}
	}

	public static void ShowBangKe(string sFileName)
	{
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string text = Path.GetTempPath() + sFileName;
			File.Copy(HTKKConfig.INTERFACE_TEMPLATES + "\\excel\\" + sFileName, text, overwrite: true);
			Process.Start(text);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage = new frmMessage("0433", "3", "", "Thông báo");
			((Form)frmMessage).ShowDialog();
		}
	}

	public static int GetTableRowIndex(int startRow, int iRow)
	{
		return (iRow - startRow >= 0) ? (iRow - startRow) : (-1);
	}

	public static HTKKCMCRowValue GetRowValueDefault(HTKKCMCRowInfo rowInfo)
	{
		HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
		int num = 0;
		foreach (HTKKCMCTableCellInfo cell in rowInfo.Cells)
		{
			string defaultValue = cell.DefaultValue;
			hTKKCMCRowValue.RowValue.Add(defaultValue);
			num++;
		}
		return hTKKCMCRowValue;
	}

	public static bool checkHaveNewDataOfTable(HTKKCMCRowInfo rowInfo, List<string> rowData)
	{
		HTKKCMCRowValue rowValueDefault = GetRowValueDefault(rowInfo);
		if (rowData.SequenceEqual(rowValueDefault.RowValue))
		{
			return false;
		}
		return true;
	}

	public static void ReindexTableId(HTKKCMCSheetInfo sInfo)
	{
		try
		{
			string text = "";
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (!(table.TablePath != text))
				{
					continue;
				}
				List<HTKKCMCTableInfo> list = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TablePath == table.TablePath).ToList();
				int num = 1;
				foreach (HTKKCMCTableInfo item in list)
				{
					item.TableID = num.ToString();
					num++;
				}
				text = table.TablePath;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static int ValidateCombobox(SheetView sView, HTKKCMCTableInfo table, List<int> colsName, List<int> colsMa, string errorMessageCode, string errorMessageName, string errorMessageCodeNull = "")
	{
		try
		{
			int num = 0;
			ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				for (int j = 0; j < colsName.Count; j++)
				{
					Cell cell = sView.Cells[dataLocationTemplateRowStart + i - 1, colsName[j] - 1];
					string text = sView.Cells[dataLocationTemplateRowStart + i - 1, colsMa[j] - 1].Text;
					string text2 = cell.Text.Trim();
					ICellType cellType = cell.CellType;
					if ((object)cellType.GetType() != comboBoxCellType.GetType())
					{
						continue;
					}
					if (!string.IsNullOrEmpty(errorMessageCode))
					{
						Dictionary<string, string> dMFromCombobox = GetDMFromCombobox(sView, dataLocationTemplateRowStart + i, colsName[j]);
						if (dMFromCombobox.Count > 0 && !CheckKeySameValue(dMFromCombobox) && !dMFromCombobox.ContainsKey(text))
						{
							HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, colsName[j] - 1, errorMessageCode);
							num++;
						}
					}
					if (!string.IsNullOrEmpty(errorMessageName))
					{
						Dictionary<string, string> dMFromCombobox = GetDMFromCombobox(sView, dataLocationTemplateRowStart + i, colsName[j]);
						if (dMFromCombobox.Count > 0 && !CheckKeySameValue(dMFromCombobox) && dMFromCombobox.ContainsKey(text) && dMFromCombobox[text] != text2)
						{
							HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, colsName[j] - 1, errorMessageName);
							num++;
						}
					}
					if (!string.IsNullOrEmpty(errorMessageCodeNull) && !string.IsNullOrEmpty(text2) && string.IsNullOrEmpty(text))
					{
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, colsName[j] - 1, errorMessageCodeNull);
						num++;
					}
				}
			}
			return num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return 0;
		}
	}

	public static void TinhLaiCacGiaTriKhiThayDoiTuBangSangBang(SheetView sView, HTKKCMCSheetInfo sInfo, string fromTableName, string toTableName, int fromColKey, int toColKey, List<int> colsCalculator)
	{
		try
		{
			foreach (int item in colsCalculator)
			{
				Dictionary<string, double> dictionary = new Dictionary<string, double>();
				List<HTKKCMCTableInfo> list = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == fromTableName).ToList();
				foreach (HTKKCMCTableInfo item2 in list)
				{
					int dataLocationTemplateRowStart = item2.Location.DataLocationTemplateRowStart;
					for (int num = 0; num < item2.RowsValues.Count; num++)
					{
						string key = sView.Cells[dataLocationTemplateRowStart + num - 1, fromColKey - 1].Text.Trim();
						double num2 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num - 1, item - 1].Value);
						if (!dictionary.ContainsKey(key))
						{
							dictionary.Add(key, num2);
						}
						else
						{
							dictionary[key] += num2;
						}
					}
				}
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TableName == toTableName);
				int dataLocationTemplateRowStart2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int num = 0; num < hTKKCMCTableInfo.RowsValues.Count; num++)
				{
					string key2 = sView.Cells[dataLocationTemplateRowStart2 + num - 1, toColKey - 1].Text.Trim();
					sView.Cells[dataLocationTemplateRowStart2 + num - 1, item - 1].Value = (dictionary.ContainsKey(key2) ? dictionary[key2] : 0.0);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void TinhLaiCacGiaTriKhiThayDoiTuBangF7SangBangF7(SheetView sView, HTKKCMCSheetInfo sInfo, string fromTableName, string toTableName, int fromColHeaderKey, int fromColBodyKey, int toColHeaderKey, int toColBodyKey, List<int> colsCalculator)
	{
		try
		{
			foreach (int item in colsCalculator)
			{
				Dictionary<string, double> dictionary = new Dictionary<string, double>();
				List<HTKKCMCTableInfo> list = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == fromTableName).ToList();
				foreach (HTKKCMCTableInfo item2 in list)
				{
					int dataLocationTemplateRowStart = item2.Location.DataLocationTemplateRowStart;
					string text = sView.Cells[item2.Location.TableTemplateRowStart - 1, fromColHeaderKey - 1].Text.Trim();
					for (int num = 0; num < item2.RowsValues.Count; num++)
					{
						string text2 = sView.Cells[dataLocationTemplateRowStart + num - 1, fromColBodyKey - 1].Text.Trim();
						double num2 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num - 1, item - 1].Value);
						string key = text + text2;
						if (!dictionary.ContainsKey(key))
						{
							dictionary.Add(key, num2);
						}
						else
						{
							dictionary[key] += num2;
						}
					}
				}
				List<HTKKCMCTableInfo> list2 = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == toTableName).ToList();
				foreach (HTKKCMCTableInfo item3 in list2)
				{
					int dataLocationTemplateRowStart2 = item3.Location.DataLocationTemplateRowStart;
					string text3 = sView.Cells[item3.Location.TableTemplateRowStart - 1, toColHeaderKey - 1].Text.Trim();
					for (int num = 0; num < item3.RowsValues.Count; num++)
					{
						string text4 = sView.Cells[dataLocationTemplateRowStart2 + num - 1, toColBodyKey - 1].Text.Trim();
						string key = text3 + text4;
						sView.Cells[dataLocationTemplateRowStart2 + num - 1, item - 1].Value = (dictionary.ContainsKey(key) ? dictionary[key] : 0.0);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CopyTableToTable3(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKCMCTableInfo fromTable, int toTableIndex, Dictionary<int, int> dicInfo, int iRow, ref List<string> listKeyTo, ref List<string> listKeyOld, bool usingCellType = false)
	{
		try
		{
			NumberCellType numberCellType = new NumberCellType();
			CurrencyCellType currencyCellType = new CurrencyCellType();
			TextCellType textCellType = new TextCellType();
			GeneralCellType generalCellType = new GeneralCellType();
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[toTableIndex];
			int dataLocationTemplateRowStart = fromTable.Location.DataLocationTemplateRowStart;
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			if (listKeyTo.Count == 0 && fromTable.RowsValues.Count > 0 && hTKKCMCTableInfo.RowsValues.Count > 0)
			{
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					string text = "";
					foreach (int key in dicInfo.Keys)
					{
						int num2 = dicInfo[key];
						object value = sView.Cells[num + i - 1, num2 - 1].Value;
						ICellType cellType = sView.Cells[num + i - 1, num2 - 1].CellType;
						if (value != null && (object)cellType.GetType() != numberCellType.GetType() && (object)cellType.GetType() != currencyCellType.GetType() && value != null)
						{
							text += value.ToString();
						}
					}
					if (!string.IsNullOrEmpty(text))
					{
						text = hTKKCMCTableInfo.TableName + hTKKCMCTableInfo.TableID + text;
						if (!listKeyTo.Contains(text))
						{
							listKeyTo.Add(text);
						}
					}
				}
			}
			string keyTo = "";
			string keyFrom = "";
			List<int> list = new List<int>();
			foreach (int key2 in dicInfo.Keys)
			{
				int num2 = dicInfo[key2];
				object value = sView.Cells[iRow - 1, key2 - 1].Value;
				ICellType cellType = sView.Cells[iRow - 1, key2 - 1].CellType;
				if (value != null && (object)cellType.GetType() != numberCellType.GetType() && (object)cellType.GetType() != currencyCellType.GetType() && value != null)
				{
					keyTo += value.ToString();
					list.Add(num2);
				}
			}
			if (string.IsNullOrEmpty(keyTo))
			{
				return;
			}
			keyFrom = fromTable.TableName + fromTable.TableID + keyTo;
			keyTo = hTKKCMCTableInfo.TableName + hTKKCMCTableInfo.TableID + keyTo;
			bool flag = false;
			if (listKeyTo.Contains(keyTo))
			{
				num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + listKeyTo.FindIndex((string x) => x == keyTo);
			}
			else
			{
				object value2 = sView.Cells[num - 1, dicInfo.Values.ToArray()[0] - 1].Value;
				if (value2 != null && !string.IsNullOrEmpty(value2.ToString().Trim()))
				{
					listKeyTo.Add(keyTo);
					num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + ((listKeyTo.Count - 2 >= 0) ? (listKeyTo.Count - 2) : 0);
					InsertTableRow(sView, num, -1, num, -1, -1, sInfo, toTableIndex);
					num++;
				}
				else
				{
					listKeyTo.Add(keyTo);
				}
				flag = true;
			}
			foreach (int key3 in dicInfo.Keys)
			{
				int num2 = dicInfo[key3];
				object value = sView.Cells[iRow - 1, key3 - 1].Value;
				ICellType cellType = sView.Cells[iRow - 1, key3 - 1].CellType;
				if (value == null)
				{
					continue;
				}
				if ((object)cellType.GetType() == numberCellType.GetType() || (object)cellType.GetType() == currencyCellType.GetType())
				{
					if (usingCellType)
					{
						sView.Cells[num - 1, num2 - 1].CellType = cellType;
					}
					sView.Cells[num - 1, num2 - 1].Value = Convert.ToDouble(value);
				}
				else
				{
					sView.Cells[num - 1, num2 - 1].Value = value;
				}
			}
			int num3 = iRow - dataLocationTemplateRowStart;
			if (listKeyOld.Count > 0 && num3 >= 0 && listKeyOld.Count > num3)
			{
				string text2 = listKeyOld[num3];
				if (!flag && !string.IsNullOrEmpty(text2) && hTKKCMCTableInfo.RowsValues.Count > 1)
				{
					string keyNew = text2.Replace(fromTable.TableName + fromTable.TableID, hTKKCMCTableInfo.TableName + hTKKCMCTableInfo.TableID);
					if (listKeyTo.Contains(keyNew))
					{
						int num4 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + listKeyTo.FindIndex((string x) => x == keyNew);
						int num5 = 0;
						foreach (int item in list)
						{
							string value3 = sView.Cells[num - 1, item - 1].Text.Trim();
							if (text2.IndexOf(value3) <= 0)
							{
								num5++;
							}
						}
						if (num5 > 0 && num5 == list.Count)
						{
							RemoveDataRows(sView, toTableIndex, num4 - 1, 1, sInfo);
							listKeyTo.Remove(keyNew);
						}
					}
				}
			}
			if (listKeyOld.Contains(keyFrom))
			{
				int index = listKeyOld.FindIndex((string x) => x == keyFrom);
				listKeyOld[index] = keyFrom;
			}
			else
			{
				listKeyOld.Add(keyFrom);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static List<string> GetRowOnSViewCreateDictionary(SheetView sView, HTKKCMCTableInfo table, List<int> cols)
	{
		List<string> list = new List<string>();
		try
		{
			NumberCellType numberCellType = new NumberCellType();
			CurrencyCellType currencyCellType = new CurrencyCellType();
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				string text = "";
				foreach (int col in cols)
				{
					object value = sView.Cells[dataLocationTemplateRowStart + i - 1, col - 1].Value;
					ICellType cellType = sView.Cells[dataLocationTemplateRowStart + i - 1, col - 1].CellType;
					if (value != null && (object)cellType.GetType() != numberCellType.GetType() && (object)cellType.GetType() != currencyCellType.GetType() && value != null)
					{
						text += value.ToString();
					}
				}
				if (!string.IsNullOrEmpty(text))
				{
					text = table.TableName + table.TableID + text;
					if (!list.Contains(text))
					{
						list.Add(text);
					}
				}
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return list;
		}
	}

	public static bool ValidateDiaBan(string maTinh, string maHuyen, string maXa)
	{
		try
		{
			bool result = true;
			if (maTinh.Length != 0 && maHuyen.Length != 0 && maXa.Length != 0)
			{
				if (maTinh.Length != 3 || maHuyen.Length != 5 || maXa.Length != 7)
				{
					result = false;
				}
				else if (maHuyen.Substring(0, 3) != maTinh || maXa.Substring(0, 5) != maHuyen)
				{
					result = false;
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

	public static bool ValidateDiaBan2Cap(string maTinh, string maHuyen, string maXa)
	{
		try
		{
			bool result = true;
			if (maTinh.Length != 0 && maXa.Length != 0 && maTinh.Length == 3 && maXa.Length == 7 && maXa.Substring(0, 3) != maTinh)
			{
				result = false;
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static void BindDataLargeSetValueByCols(SheetView sView, HTKKCMCTableInfo table)
	{
		try
		{
			if (table == null || table.RowsValues.Count == 0)
			{
				return;
			}
			int count = table.RowsValues.Count;
			List<object[,]> list = new List<object[,]>();
			for (int i = 0; i < table.RowsInfo.Cells.Count; i++)
			{
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells[i];
				object[,] array = new object[count, 1];
				for (int j = 0; j < count; j++)
				{
					string text = table.RowsValues[j].RowValue[i].Trim();
					try
					{
						if (hTKKCMCTableCellInfo.Controltype == "16")
						{
							array[j, 0] = (string.IsNullOrEmpty(text) ? 0.0 : Convert.ToDouble(text.Replace(",", ".")));
						}
						else if (hTKKCMCTableCellInfo.Controltype == "2")
						{
							array[j, 0] = ((text.ToLower() == "1" || text.ToLower() == "true" || text.ToLower() == "x") ? true : false);
						}
						else
						{
							array[j, 0] = text;
						}
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					}
				}
				if (((Control)sView.FpSpread).InvokeRequired)
				{
					sView.SetArray(table.Location.DataLocationTemplateRowStart - 1, table.RowsInfo.Cells[i].CellTemplate.ColIndex - 1, array);
				}
				else
				{
					sView.SetArray(table.Location.DataLocationTemplateRowStart - 1, table.RowsInfo.Cells[i].CellTemplate.ColIndex - 1, array);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertTableRowNoSView(int frRow, int rowCount, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			sInfo.IncrementStaticCell(frRow, rowCount);
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (table.Location.DataLocationTemplateRowStart > frRow)
				{
					table.Increment(rowCount);
				}
				table.IncrementFooter(frRow, rowCount);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void TinhSoThueGTGT_03_HDDT(HTKKCMCTableInfo table, int rowIndexInTable, string maThueSuat)
	{
		try
		{
			string pathDoanhThu = "DoanhThu";
			string pathThueGTGT = "ThueGTGT";
			string pathGiaTriThueSuat = "path_13";
			int index = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pathDoanhThu);
			int index2 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pathThueGTGT);
			int index3 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pathGiaTriThueSuat);
			HTKKCMCRowValue hTKKCMCRowValue = table.RowsValues[rowIndexInTable];
			double num = (string.IsNullOrEmpty(hTKKCMCRowValue.RowValue[index]) ? 0.0 : Convert.ToDouble(hTKKCMCRowValue.RowValue[index]));
			string value = hTKKCMCRowValue.RowValue[index3];
			double num2 = 0.0;
			if (maThueSuat.Equals("0%") || maThueSuat.Equals("KCT") || maThueSuat.Equals("KKKNT"))
			{
				num2 = 0.0;
			}
			else if (maThueSuat.Equals("5%"))
			{
				num2 = Math.Round(0.05 * num, 0);
			}
			else if (maThueSuat.Equals("8%"))
			{
				num2 = Math.Round(0.08 * num, 0);
			}
			else if (maThueSuat.Equals("10%"))
			{
				num2 = Math.Round(0.1 * num, 0);
			}
			else if (maThueSuat.Equals("KHAC"))
			{
				if ("".Equals(value))
				{
					value = "0";
				}
				num2 = Math.Round(Convert.ToDouble(value) / 100.0 * num, 0);
			}
			else
			{
				num2 = 0.0;
			}
			hTKKCMCRowValue.RowValue[index2] = num2.ToString();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static int ValidateComboboxTableHeader(SheetView sView, HTKKCMCTableInfo table, List<string> pathsName, List<string> pathsCode, string errorMessageCode, string errorMessageName, string errorMessageCodeNull)
	{
		try
		{
			int num = 0;
			ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
			int p;
			for (p = 0; p < pathsName.Count; p++)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathsName[p]);
				HTKKCMCCellInfo hTKKCMCCellInfo2 = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathsCode[p]);
				Cell cell = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1];
				string text = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
				string text2 = cell.Text.Trim();
				ICellType cellType = cell.CellType;
				if ((object)cellType.GetType() != comboBoxCellType.GetType())
				{
					continue;
				}
				if (!string.IsNullOrEmpty(errorMessageCode))
				{
					Dictionary<string, string> dMFromCombobox = GetDMFromCombobox(sView, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex);
					if (dMFromCombobox.Count > 0 && !CheckKeySameValue(dMFromCombobox) && !dMFromCombobox.ContainsKey(text))
					{
						HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, errorMessageCode);
						num++;
					}
				}
				if (!string.IsNullOrEmpty(errorMessageName))
				{
					Dictionary<string, string> dMFromCombobox = GetDMFromCombobox(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
					if (dMFromCombobox.Count > 0 && !CheckKeySameValue(dMFromCombobox) && dMFromCombobox.ContainsKey(text) && dMFromCombobox[text] != text2)
					{
						HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, errorMessageName);
						num++;
					}
				}
				if (!string.IsNullOrEmpty(errorMessageCodeNull) && !string.IsNullOrEmpty(text2) && string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, errorMessageCodeNull);
					num++;
				}
			}
			return num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return 0;
		}
	}

	public static int ValidateComboboxStatic(SheetView sView, HTKKCMCSheetInfo sInfo, List<string> pathsName, List<string> pathsCode, string errorMessageCode, string errorMessageName, string errorMessageCodeNull)
	{
		try
		{
			int num = 0;
			ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
			int p;
			for (p = 0; p < pathsName.Count; p++)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathsName[p]);
				HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathsCode[p]);
				Cell cell = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1];
				string text = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
				string text2 = cell.Text.Trim();
				ICellType cellType = cell.CellType;
				if ((object)cellType.GetType() != comboBoxCellType.GetType())
				{
					continue;
				}
				if (!string.IsNullOrEmpty(errorMessageCode))
				{
					Dictionary<string, string> dMFromCombobox = GetDMFromCombobox(sView, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex);
					if (dMFromCombobox.Count > 0 && !CheckKeySameValue(dMFromCombobox) && !dMFromCombobox.ContainsKey(text))
					{
						HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, errorMessageCode);
						num++;
					}
				}
				if (!string.IsNullOrEmpty(errorMessageName))
				{
					Dictionary<string, string> dMFromCombobox = GetDMFromCombobox(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
					if (dMFromCombobox.Count > 0 && !CheckKeySameValue(dMFromCombobox) && dMFromCombobox.ContainsKey(text) && dMFromCombobox[text] != text2)
					{
						HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, errorMessageName);
						num++;
					}
				}
				if (!string.IsNullOrEmpty(errorMessageCodeNull) && !string.IsNullOrEmpty(text2) && string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, errorMessageCodeNull);
					num++;
				}
			}
			return num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return 0;
		}
	}

	public static void CheckComboboxEmptyCode(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, ref Dictionary<int, int> levelWarning)
	{
		try
		{
			if ("N7~N8~N5~T8".IndexOf(ToKhaiInfo.MenuID) < 0)
			{
				return;
			}
			ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
			int num = ((ToKhaiInfo.KieuKHBS == "3") ? (ToKhaiInfo.SheetInfos.Count - 2) : (ToKhaiInfo.SheetInfos.Count - 1));
			for (int i = 0; i < num; i++)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[i];
				if ((ToKhaiInfo.MenuID == "T8" && i == 0) || hTKKCMCSheetInfo == null || hTKKCMCSheetInfo.SheetName.ToLower().IndexOf("khbs") >= 0)
				{
					continue;
				}
				short sheetIndex = hTKKCMCSheetInfo.SheetIndex;
				SheetView sheetView = winFpSpread.Sheets[sheetIndex];
				foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
				{
					if (table == null || table.RowsValues.Count == 0)
					{
						continue;
					}
					IEnumerable<HTKKCMCTableCellInfo> enumerable = table.RowsInfo.Cells.Where((HTKKCMCTableCellInfo x) => x.Controltype != "0" && x.Controltype != "15" && x.Controltype != "1" && x.Controltype != "2" && x.Controltype != "14" && x.Controltype != "96" && x.Controltype != "94" && x.Controltype != "16" && x.Controltype != "17" && x.Controltype != "26" && x.Controltype != "27" && x.Controltype != "52");
					if (enumerable == null || enumerable.Count() <= 0)
					{
						continue;
					}
					int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
					for (int num2 = 0; num2 < table.RowsValues.Count; num2++)
					{
						foreach (HTKKCMCTableCellInfo item in enumerable)
						{
							if (item == null || string.IsNullOrEmpty(item.indexColumnSelectedValue))
							{
								continue;
							}
							string text = sheetView.Cells[dataLocationTemplateRowStart + num2 - 1, item.CellTemplate.ColIndex - 1].Text;
							string text2 = sheetView.Cells[dataLocationTemplateRowStart + num2 - 1, HTKKFile.ColumnLettersToIndex(item.indexColumnSelectedValue) - 1].Text;
							if (string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, item.CellTemplate.ColIndex - 1, HTKKFile.GetMessage("0629"), num2 + 1);
								sheetView.FpSpread.TabStrip[sheetIndex].BackColor = Color.Pink;
								if (levelWarning.ContainsKey(i))
								{
									levelWarning[i] = 2;
								}
								else
								{
									levelWarning.Add(i, 2);
								}
							}
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void GetParentDanhMuc(ref Dictionary<string, string> dic, ref Dictionary<string, string[]> temp, string maCha, List<item> items)
	{
		try
		{
			if (string.IsNullOrEmpty(maCha))
			{
				return;
			}
			string[] array;
			if (maCha.Length > 5)
			{
				item item = items.FirstOrDefault((item x) => x.ma == maCha.Substring(0, maCha.Length - 2));
				if (item != null && !string.IsNullOrEmpty(item.ma) && !dic.ContainsKey(maCha))
				{
					array = new string[7]
					{
						item.id,
						item.ma,
						"---" + item.ten,
						item.dvt,
						item.thuesuat,
						"0",
						"1"
					};
					if (array.Length >= 2)
					{
						string key = array[1];
						string text = array[2];
						if (!string.IsNullOrEmpty(item.thuesuat))
						{
							text = text + " (" + item.thuesuat + " %)";
						}
						if (!dic.ContainsKey(key))
						{
							dic.Add(key, text);
						}
						if (!temp.ContainsKey(text))
						{
							temp.Add(text, array);
						}
					}
				}
			}
			item item2 = items.FirstOrDefault((item x) => x.ma == maCha);
			if (item2 == null || string.IsNullOrEmpty(item2.ma) || dic.ContainsKey(maCha))
			{
				return;
			}
			array = new string[7]
			{
				item2.id,
				item2.ma,
				"-----" + item2.ten,
				item2.dvt,
				item2.thuesuat,
				"0",
				"1"
			};
			if (array.Length >= 2)
			{
				string key = array[1];
				string text = array[2];
				if (!string.IsNullOrEmpty(item2.thuesuat))
				{
					text = text + " (" + item2.thuesuat + " %)";
				}
				if (!dic.ContainsKey(key))
				{
					dic.Add(key, text);
				}
				if (!temp.ContainsKey(text))
				{
					temp.Add(text, array);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void GetCQTChiCucTuCQTCuc(string maCQTCapCuc, SheetView sView, int colTenCQTChiCuc, int colMaCQTChiCuc, int iRow)
	{
		try
		{
			if (!string.IsNullOrEmpty(maCQTCapCuc))
			{
				string text = maCQTCapCuc.Substring(0, 3) + "00";
				if (HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy.ContainsKey(text))
				{
					string text2 = HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy[text];
					sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Text = text2;
					sView.Cells[iRow - 1, colMaCQTChiCuc - 1].Text = text;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool CheckKeySameValue(Dictionary<string, string> dic)
	{
		bool result = false;
		foreach (KeyValuePair<string, string> item in dic)
		{
			if (item.Key == item.Value)
			{
				result = true;
				break;
			}
		}
		return result;
	}

	public static void FormatCol4567KHBSBySinfo(HTKKCMCTableCellInfo fromCell, SheetView sView_KHBS, int row_KHBS, int col4, int col5, int col6, int col7)
	{
		try
		{
			NumberCellType numberCellType = new NumberCellType();
			numberCellType.DecimalPlaces = 0;
			CurrencyCellType currencyCellType = new CurrencyCellType();
			TextCellType textCellType = new TextCellType();
			CheckBoxCellType checkBoxCellType = new CheckBoxCellType();
			MaskCellType maskCellType = new MaskCellType();
			object obj = FormatCelltypeFromSinfo(fromCell.Validate, fromCell.DefaultValue, fromCell.Controltype);
			ICellType cellType = sView_KHBS.Cells[0, 0].CellType;
			if (!(fromCell.Controltype == "16"))
			{
				cellType = ((fromCell.Controltype == "2") ? ((CheckBoxCellType)obj) : ((!(fromCell.Controltype == "14")) ? ((ICellType)(TextCellType)obj) : ((ICellType)(MaskCellType)obj)));
			}
			else
			{
				try
				{
					cellType = (NumberCellType)obj;
				}
				catch (Exception)
				{
					cellType = (CurrencyCellType)obj;
				}
			}
			object obj2 = sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value;
			object obj3 = sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value;
			if ((object)cellType.GetType() == numberCellType.GetType() || (object)cellType.GetType() == currencyCellType.GetType())
			{
				if (obj2 != null)
				{
					obj2 = obj2.ToString().Replace(",", ".");
				}
				if (obj3 != null)
				{
					obj3 = obj3.ToString().Replace(",", ".");
				}
				sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value = ((obj2 != null && !string.IsNullOrEmpty(obj2.ToString().Trim()) && IsNumeric(obj2.ToString())) ? Convert.ToDouble(obj2) : 0.0);
				sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value = ((obj3 != null && !string.IsNullOrEmpty(obj3.ToString().Trim()) && IsNumeric(obj3.ToString())) ? Convert.ToDouble(obj3) : 0.0);
				sView_KHBS.Cells[row_KHBS - 1, col4 - 1].CellType = cellType;
				sView_KHBS.Cells[row_KHBS - 1, col5 - 1].CellType = cellType;
				sView_KHBS.Cells[row_KHBS - 1, col6 - 1].CellType = cellType;
				sView_KHBS.Cells[row_KHBS - 1, col7 - 1].CellType = cellType;
				return;
			}
			sView_KHBS.Cells[row_KHBS - 1, col4 - 1].CellType = cellType;
			sView_KHBS.Cells[row_KHBS - 1, col5 - 1].CellType = cellType;
			if ((object)cellType.GetType() == checkBoxCellType.GetType() && (sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Text.Trim() == "1" || sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Text.Trim().ToLower() == "true"))
			{
				sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value = true;
			}
			if ((object)cellType.GetType() == checkBoxCellType.GetType() && (sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Text.Trim() == "1" || sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Text.Trim().ToLower() == "true"))
			{
				sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value = true;
			}
			if ((object)cellType.GetType() == maskCellType.GetType())
			{
				if (sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value != null)
				{
					string text = sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value.ToString().Trim();
					if (text.Length == 10 && text.IndexOf("-") >= 0)
					{
						string value = DateTime.ParseExact(text, "yyyy-MM-dd", null).ToString("ddMMyyyy");
						sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value = value;
					}
				}
				if (sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value != null)
				{
					string text2 = sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value.ToString().Trim();
					if (text2.Length == 10 && text2.IndexOf("-") >= 0)
					{
						string value2 = DateTime.ParseExact(text2, "yyyy-MM-dd", null).ToString("ddMMyyyy");
						sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value = value2;
					}
				}
			}
			sView_KHBS.Cells[row_KHBS - 1, col6 - 1].CellType = numberCellType;
			sView_KHBS.Cells[row_KHBS - 1, col7 - 1].CellType = numberCellType;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FormatCol4567KHBSBySinfoStatic(HTKKCMCCellInfo fromCell, SheetView sView_KHBS, int row_KHBS, int col4, int col5, int col6, int col7)
	{
		try
		{
			NumberCellType numberCellType = new NumberCellType();
			numberCellType.DecimalPlaces = 0;
			CurrencyCellType currencyCellType = new CurrencyCellType();
			TextCellType textCellType = new TextCellType();
			CheckBoxCellType checkBoxCellType = new CheckBoxCellType();
			MaskCellType maskCellType = new MaskCellType();
			object obj = FormatCelltypeFromSinfo(fromCell.Validate, fromCell.DefaultValue, fromCell.Controltype);
			ICellType cellType = sView_KHBS.Cells[0, 0].CellType;
			if (!(fromCell.Controltype == "16"))
			{
				cellType = ((fromCell.Controltype == "2") ? ((CheckBoxCellType)obj) : ((!(fromCell.Controltype == "14")) ? ((ICellType)(TextCellType)obj) : ((ICellType)(MaskCellType)obj)));
			}
			else
			{
				try
				{
					cellType = (NumberCellType)obj;
				}
				catch (Exception)
				{
					cellType = (CurrencyCellType)obj;
				}
			}
			object obj2 = sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value;
			object obj3 = sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value;
			if ((object)cellType.GetType() == numberCellType.GetType() || (object)cellType.GetType() == currencyCellType.GetType())
			{
				if (obj2 != null)
				{
					obj2 = obj2.ToString().Replace(",", ".");
				}
				if (obj3 != null)
				{
					obj3 = obj3.ToString().Replace(",", ".");
				}
				sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value = ((obj2 != null && !string.IsNullOrEmpty(obj2.ToString().Trim()) && IsNumeric(obj2.ToString())) ? Convert.ToDouble(obj2) : 0.0);
				sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value = ((obj3 != null && !string.IsNullOrEmpty(obj3.ToString().Trim()) && IsNumeric(obj3.ToString())) ? Convert.ToDouble(obj3) : 0.0);
				sView_KHBS.Cells[row_KHBS - 1, col4 - 1].CellType = cellType;
				sView_KHBS.Cells[row_KHBS - 1, col5 - 1].CellType = cellType;
				sView_KHBS.Cells[row_KHBS - 1, col6 - 1].CellType = cellType;
				sView_KHBS.Cells[row_KHBS - 1, col7 - 1].CellType = cellType;
				return;
			}
			sView_KHBS.Cells[row_KHBS - 1, col4 - 1].CellType = cellType;
			sView_KHBS.Cells[row_KHBS - 1, col5 - 1].CellType = cellType;
			if ((object)cellType.GetType() == checkBoxCellType.GetType() && (sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Text.Trim() == "1" || sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Text.Trim().ToLower() == "true"))
			{
				sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value = true;
			}
			if ((object)cellType.GetType() == checkBoxCellType.GetType() && (sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Text.Trim() == "1" || sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Text.Trim().ToLower() == "true"))
			{
				sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value = true;
			}
			if ((object)cellType.GetType() == maskCellType.GetType())
			{
				if (sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value != null)
				{
					string text = sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value.ToString().Trim();
					if (text.Length == 10 && text.IndexOf("-") >= 0)
					{
						string value = DateTime.ParseExact(text, "yyyy-MM-dd", null).ToString("ddMMyyyy");
						sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value = value;
					}
				}
				if (sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value != null)
				{
					string text2 = sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value.ToString().Trim();
					if (text2.Length == 10 && text2.IndexOf("-") >= 0)
					{
						string value2 = DateTime.ParseExact(text2, "yyyy-MM-dd", null).ToString("ddMMyyyy");
						sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value = value2;
					}
				}
			}
			sView_KHBS.Cells[row_KHBS - 1, col6 - 1].CellType = numberCellType;
			sView_KHBS.Cells[row_KHBS - 1, col7 - 1].CellType = numberCellType;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static object FormatCelltypeFromSinfo(HTKKCMCValidate validate, string defaultValue, string controlType)
	{
		TextCellType textCellType = new TextCellType();
		try
		{
			switch (controlType)
			{
			case "16":
			{
				int decimalPlaces;
				if (validate.MaxValue.Length <= 15)
				{
					NumberCellType numberCellType = new NumberCellType();
					decimalPlaces = 0;
					numberCellType.ShowSeparator = true;
					numberCellType.NegativeFormat = NegativeFormat.Parentheses;
					if (HTKKFormater.isNullString(defaultValue) || defaultValue == "0")
					{
						decimalPlaces = 0;
						numberCellType.DecimalSeparator = ",";
						numberCellType.Separator = ".";
					}
					else if (defaultValue.IndexOf(".") > -1)
					{
						decimalPlaces = defaultValue.Substring(defaultValue.LastIndexOf(".") + 1).Length;
						numberCellType.DecimalSeparator = ".";
						numberCellType.Separator = ",";
					}
					else if (defaultValue.IndexOf(",") > -1)
					{
						decimalPlaces = defaultValue.Substring(defaultValue.LastIndexOf(",") + 1).Length;
						numberCellType.DecimalSeparator = ",";
						numberCellType.Separator = ".";
					}
					numberCellType.MinimumValue = getMinNumber(validate.MinValue);
					numberCellType.MaximumValue = getMaxNumber(validate.MaxValue);
					numberCellType.DecimalPlaces = decimalPlaces;
					return numberCellType;
				}
				CurrencyCellType currencyCellType = new CurrencyCellType();
				decimalPlaces = 0;
				currencyCellType.ShowSeparator = true;
				currencyCellType.ShowCurrencySymbol = false;
				if (HTKKFormater.isNullString(defaultValue) || defaultValue == "0")
				{
					decimalPlaces = 0;
					currencyCellType.DecimalSeparator = ",";
					currencyCellType.Separator = ".";
				}
				else if (defaultValue.IndexOf(".") > -1)
				{
					decimalPlaces = defaultValue.Substring(defaultValue.LastIndexOf(".") + 1).Length;
					currencyCellType.DecimalSeparator = ".";
					currencyCellType.Separator = ",";
				}
				else if (defaultValue.IndexOf(",") > -1)
				{
					decimalPlaces = defaultValue.Substring(defaultValue.LastIndexOf(",") + 1).Length;
					currencyCellType.DecimalSeparator = ",";
					currencyCellType.Separator = ".";
				}
				currencyCellType.MinimumValue = getMinNumber(validate.MinValue);
				currencyCellType.MaximumValue = getMaxNumber(validate.MaxValue);
				currencyCellType.DecimalPlaces = decimalPlaces;
				return currencyCellType;
			}
			case "2":
				return new CheckBoxCellType();
			case "14":
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/00/0000";
				maskCellType.MaskChar = '.';
				return maskCellType;
			}
			case "96":
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00:00:00";
				maskCellType.MaskChar = '.';
				return maskCellType;
			}
			case "94":
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/00/0000 00:00:00";
				maskCellType.MaskChar = '.';
				return maskCellType;
			}
			default:
			{
				int num = (textCellType.MaxLength = getMaxLen(validate.MaxLen));
				if (num != 7)
				{
					textCellType.WordWrap = true;
					textCellType.Multiline = true;
				}
				return textCellType;
			}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return textCellType;
		}
	}

	public static double CalculateSumHeaderByGroup(SheetView sView, List<HTKKCMCTableInfo> tables, int colSum)
	{
		double num = 0.0;
		try
		{
			foreach (HTKKCMCTableInfo table in tables)
			{
				int num2 = -1;
				if (table.Header != null && table.Header.StaticCells.Count > 0)
				{
					num2 = table.Header.StaticCells[0].CellTemplate.RowIndex;
				}
				if (num2 >= 0)
				{
					double num3 = Convert.ToDouble(sView.Cells[num2 - 1, colSum - 1].Value);
					num += num3;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return num;
	}

	public static void SetLaiDongConBangF5(SheetView sView, HTKKCMCTableInfo table, object value, int colChilden)
	{
		try
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				sView.Cells[dataLocationTemplateRowStart + i - 1, colChilden - 1].Value = value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string FomartTenChiTieu(string chiTieu, string cot, string freetext)
	{
		if (!string.IsNullOrEmpty(chiTieu))
		{
			return "Chỉ tiêu [" + chiTieu + "] ";
		}
		if (!string.IsNullOrEmpty(chiTieu))
		{
			return "Cột (" + cot + ") ";
		}
		return freetext + " ";
	}

	private static int CountWeekendDays(DateTime startDate, DateTime endDate, DayOfWeek dayOfWeek)
	{
		int num = 0;
		while (startDate <= endDate)
		{
			if (startDate.DayOfWeek == dayOfWeek)
			{
				num++;
			}
			startDate = startDate.AddDays(1.0);
		}
		return num;
	}

	public static void ComboAutocomplete(SheetView sView, int idxRow, int idxCol, Dictionary<string, string> dicDM)
	{
		try
		{
			ComboBoxCellType comboBoxCellType = (ComboBoxCellType)sView.Cells[idxRow, idxCol].CellType;
			string currentText = sView.Cells[idxRow, idxCol].Text;
			IEnumerable<KeyValuePair<string, string>> source = dicDM.Where((KeyValuePair<string, string> pair) => pair.Value.ToLower().Contains(currentText.ToLower()));
			string[] itemData = source.Select((KeyValuePair<string, string> pair) => pair.Key).ToArray();
			string[] array = source.Select((KeyValuePair<string, string> pair) => pair.Value).ToArray();
			comboBoxCellType.Clear();
			comboBoxCellType.Items = array;
			comboBoxCellType.ItemData = itemData;
			comboBoxCellType.ShowSubEditor();
			sView.Cells[idxRow, idxCol].CellType = comboBoxCellType;
			if (array.Length == 0)
			{
				sView.Cells[idxRow, idxCol].Text = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void ResetCombobox(SheetView sView, int idxRow, int idxCol, Dictionary<string, string> dicDM)
	{
		try
		{
			string textCombo = sView.Cells[idxRow, idxCol].Text;
			IEnumerable<KeyValuePair<string, string>> source = dicDM.Where((KeyValuePair<string, string> pair) => pair.Value.ToLower().Equals(textCombo.ToLower(), StringComparison.OrdinalIgnoreCase));
			string[] array = source.Select((KeyValuePair<string, string> pair) => pair.Key).ToArray();
			if (array.Length == 0)
			{
				sView.Cells[idxRow, idxCol].Text = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string RemoveDiacritics(string text)
	{
		string text2 = text.Normalize(NormalizationForm.FormD);
		StringBuilder stringBuilder = new StringBuilder();
		string text3 = text2;
		foreach (char c in text3)
		{
			UnicodeCategory unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
			if (unicodeCategory != UnicodeCategory.NonSpacingMark)
			{
				stringBuilder.Append(c);
			}
		}
		return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
	}

	public static void setButton(SheetView sView, int row, int col, string text)
	{
		ButtonCellType buttonCellType = new ButtonCellType();
		buttonCellType.TwoState = false;
		buttonCellType.Text = text;
		buttonCellType.ShadowSize = 0;
		buttonCellType.TextColor = Color.Blue;
		sView.Cells[row - 1, col - 1].CellType = buttonCellType;
	}

	public static void FillDataMstCu(XmlNode xmlNode)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/mst");
			if (val != null && !string.IsNullOrEmpty(val.InnerText))
			{
				string innerText = val.InnerText.Trim();
				XmlNode val2 = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/mst_cu");
				if (val2 != null)
				{
					val2.InnerText = innerText;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void ReplaceMaHH(XmlNode xmlNode, string path)
	{
		//IL_001a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0020: Expected O, but got Unknown
		try
		{
			XmlNodeList val = xmlNode.SelectNodes(path);
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				if (val2.InnerText.Contains("a"))
				{
					val2.InnerText = val2.InnerText.Replace("a", "");
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static Dictionary<string, string> loadCqtCapCucTheoDiaBanXa(string maXa)
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		if (string.IsNullOrEmpty(maXa))
		{
			return dictionary;
		}
		CQTDiaBan cQTDiaBan = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaXa == maXa);
		dictionary.Add("", "");
		if (cQTDiaBan != null)
		{
			string maCuc = cQTDiaBan.MaCQT.Substring(0, 3) + "00";
			CQTDiaBan cQTDiaBan2 = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == maCuc);
			dictionary.Add(maCuc, cQTDiaBan2.TenCQT);
			Dictionary<string, string> dictionary2 = HTKKWinDictionaryDiaDanh._cqtDiaBan.Where((CQTDiaBan x) => x.MaCQT == "82500" || x.MaCQT == "82700").ToDictionary((CQTDiaBan x) => x.MaCQT, (CQTDiaBan x) => x.TenCQT);
			foreach (KeyValuePair<string, string> item in dictionary2)
			{
				dictionary.Add(item.Key, item.Value);
			}
		}
		return dictionary;
	}

	public static Dictionary<string, string> loadCqtCapChiCucTheoDiaBanXa(string maXa)
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		if (string.IsNullOrEmpty(maXa))
		{
			return dictionary;
		}
		CQTDiaBan cQTDiaBan = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaXa == maXa);
		dictionary.Add("", "");
		if (cQTDiaBan != null)
		{
			string maCuc = cQTDiaBan.MaCQT.Substring(0, 3) + "00";
			CQTDiaBan cQTDiaBan2 = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == maCuc);
			dictionary.Add(cQTDiaBan2.MaCQT, cQTDiaBan2.TenCQT);
			dictionary.Add(cQTDiaBan.MaCQT, cQTDiaBan.TenCQT);
			Dictionary<string, string> dictionary2 = HTKKWinDictionaryDiaDanh._cqtDiaBan.Where((CQTDiaBan x) => x.MaCQT == "82500" || x.MaCQT == "82700").ToDictionary((CQTDiaBan x) => x.MaCQT, (CQTDiaBan x) => x.TenCQT);
			foreach (KeyValuePair<string, string> item in dictionary2)
			{
				dictionary.Add(item.Key, item.Value);
			}
		}
		return dictionary;
	}

	public static Dictionary<string, string> loadCqtCapChiCucTheoDiaBanXa(string maCqt, string maXa)
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		if (string.IsNullOrEmpty(maXa))
		{
			return dictionary;
		}
		if (maCqt == "82500")
		{
			dictionary.Add("", "");
			dictionary.Add(maCqt, "Chi cục Thuế doanh nghiệp lớn");
			return dictionary;
		}
		if (maCqt == "82700")
		{
			dictionary.Add("", "");
			dictionary.Add(maCqt, "Chi cục Thuế Thương mại điện tử");
			return dictionary;
		}
		CQTDiaBan cQTDiaBan = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaXa == maXa);
		if (cQTDiaBan != null)
		{
			if (!string.IsNullOrEmpty(maCqt))
			{
				string maCuc = cQTDiaBan.MaCQT.Substring(0, 3) + "00";
				CQTDiaBan cQTDiaBan2 = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == maCuc);
				dictionary.Add(cQTDiaBan2.MaCQT, cQTDiaBan2.TenCQT);
			}
			dictionary.Add(cQTDiaBan.MaCQT, cQTDiaBan.TenCQT);
		}
		return dictionary;
	}

	public static void setSviewCqtCapCucTheoMaXa_Dinamic(SheetView sView, int iRow, HTKKCMCRowInfo rowInfo, string maXa)
	{
		HTKKCMCTableCellInfo hTKKCMCTableCellInfo = rowInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Controltype == "8");
		if (hTKKCMCTableCellInfo != null)
		{
			Dictionary<string, string> dic = loadCqtCapCucTheoDiaBanXa(maXa);
			InsertCombobox(sView, iRow, hTKKCMCTableCellInfo.CellTemplate.ColIndex, dic);
		}
	}

	public static void setSviewCqtCapChiCucTheoMaCqtCuc_Dinamic(SheetView sView, int iRow, HTKKCMCRowInfo rowInfo, string maCqtCuc)
	{
		HTKKCMCTableCellInfo hTKKCMCTableCellInfo = rowInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Controltype == "9");
		if (hTKKCMCTableCellInfo == null)
		{
			return;
		}
		HTKKCMCTableCellInfo hTKKCMCTableCellInfo2 = rowInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Controltype == "101");
		if (hTKKCMCTableCellInfo2 != null && !string.IsNullOrEmpty(hTKKCMCTableCellInfo2.indexColumnSelectedValue))
		{
			object value = sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex(hTKKCMCTableCellInfo2.indexColumnSelectedValue) - 1].Value;
			if (value != null)
			{
				Dictionary<string, string> dic = loadCqtCapChiCucTheoDiaBanXa(maCqtCuc, value.ToString());
				InsertCombobox(sView, iRow, hTKKCMCTableCellInfo.CellTemplate.ColIndex, dic);
			}
		}
	}

	public static void setSviewCqtCapChiCucTheoMaCqtCuc_Static(SheetView sView, HTKKCMCSheetInfo sInfo, string maCqtCuc)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Controltype == "9");
		if (hTKKCMCCellInfo == null)
		{
			return;
		}
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Controltype == "101");
		if (hTKKCMCCellInfo2 != null)
		{
			object value = sView.Cells[hTKKCMCCellInfo2.CellSelectedValue.RowIndex - 1, hTKKCMCCellInfo2.CellSelectedValue.ColIndex - 1].Value;
			if (value != null)
			{
				Dictionary<string, string> dic = loadCqtCapChiCucTheoDiaBanXa(maCqtCuc, value.ToString());
				InsertCombobox(sView, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex, dic);
			}
		}
	}

	public static void setSviewCqtCapCucTheoMaXa_Static(SheetView sView, HTKKCMCSheetInfo sheetInfo, string maXa)
	{
		IEnumerable<HTKKCMCCellInfo> enumerable = sheetInfo.StaticCells.Where((HTKKCMCCellInfo x) => x.Controltype == "8");
		if (enumerable == null)
		{
			return;
		}
		foreach (HTKKCMCCellInfo item in enumerable)
		{
			Dictionary<string, string> dic = loadCqtCapCucTheoDiaBanXa(maXa);
			InsertCombobox(sView, item.CellTemplate.RowIndex, item.CellTemplate.ColIndex, dic);
		}
	}

	public static void setSviewCqtCapChiCucTheoMaXa_Static(SheetView sView, HTKKCMCSheetInfo sheetInfo, string maXa)
	{
		IEnumerable<HTKKCMCCellInfo> enumerable = sheetInfo.StaticCells.Where((HTKKCMCCellInfo x) => x.Controltype == "9");
		if (enumerable == null)
		{
			return;
		}
		foreach (HTKKCMCCellInfo item in enumerable)
		{
			Dictionary<string, string> dic = loadCqtCapChiCucTheoDiaBanXa(maXa);
			InsertCombobox(sView, item.CellTemplate.RowIndex, item.CellTemplate.ColIndex, dic);
		}
	}

	public static string timMaChildenFromRowCol_Dinamic(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol, string controltype)
	{
		int tableIndex = sInfo.GetTableIndex(iRow, iCol);
		if (tableIndex >= 0)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[tableIndex];
			HTKKCMCTableCellInfo hTKKCMCTableCellInfo = hTKKCMCTableInfo.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Controltype == controltype);
			if (hTKKCMCTableCellInfo != null)
			{
				return sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex(hTKKCMCTableCellInfo.indexColumnSelectedValue) - 1].Text;
			}
		}
		return "";
	}

	public static string timMaChildenFromRowCol_Static(SheetView sView, List<HTKKCMCCellInfo> staticCells, string controltype)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = staticCells.Find((HTKKCMCCellInfo x) => x.Controltype == controltype);
		if (hTKKCMCCellInfo != null)
		{
			return sView.Cells[hTKKCMCCellInfo.CellSelectedValue.RowIndex - 1, hTKKCMCCellInfo.CellSelectedValue.ColIndex - 1].Text;
		}
		return "";
	}

	public static string GetMaTinhMoiByMaTinhCu(string maTinhCu)
	{
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		//IL_0041: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Dia_Ban_Tinh_Map.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				XmlAttribute val3 = val2.Attributes["MaCu"];
				XmlAttribute val4 = val2.Attributes["MaMoi"];
				if (val3 == null || val4 == null)
				{
					continue;
				}
				string[] array = ((XmlNode)val3).Value.Split(new char[1] { ',' });
				string[] array2 = array;
				foreach (string text2 in array2)
				{
					if (maTinhCu.Equals(text2.Trim()))
					{
						return ((XmlNode)val4).Value;
					}
				}
			}
			return "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static string getMaComboboxSView(SheetView sView, int iRow, int iCol)
	{
		try
		{
			Dictionary<string, string> dMFromCombobox = GetDMFromCombobox(sView, iRow, iCol);
			string tenCqt = sView.Cells[iRow - 1, iCol - 1].Text;
			if (string.IsNullOrEmpty(tenCqt))
			{
				return "";
			}
			return dMFromCombobox.First((KeyValuePair<string, string> x) => x.Value == tenCqt).Key;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static void resetCqtTheoTinhStatic(SheetView sView, List<HTKKCMCCellInfo> cells)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = cells.Find((HTKKCMCCellInfo x) => x.Controltype == "8");
		if (hTKKCMCCellInfo != null)
		{
			sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = "";
			sView.Cells[hTKKCMCCellInfo.CellSelectedValue.RowIndex - 1, hTKKCMCCellInfo.CellSelectedValue.ColIndex - 1].Value = "";
		}
		HTKKCMCCellInfo hTKKCMCCellInfo2 = cells.Find((HTKKCMCCellInfo x) => x.Controltype == "9");
		if (hTKKCMCCellInfo2 != null)
		{
			sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = "";
			sView.Cells[hTKKCMCCellInfo2.CellSelectedValue.RowIndex - 1, hTKKCMCCellInfo2.CellSelectedValue.ColIndex - 1].Value = "";
		}
	}

	public static void resetCqtTheoTinhDinamic(SheetView sView, HTKKCMCRowInfo rowInfo, int iRow)
	{
		HTKKCMCTableCellInfo hTKKCMCTableCellInfo = rowInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Controltype == "8");
		if (hTKKCMCTableCellInfo != null)
		{
			sView.Cells[iRow - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Value = "";
			sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex(hTKKCMCTableCellInfo.indexColumnSelectedValue) - 1].Value = "";
		}
		HTKKCMCTableCellInfo hTKKCMCTableCellInfo2 = rowInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Controltype == "9");
		if (hTKKCMCTableCellInfo2 != null)
		{
			sView.Cells[iRow - 1, hTKKCMCTableCellInfo2.CellTemplate.ColIndex - 1].Value = "";
			sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex(hTKKCMCTableCellInfo2.indexColumnSelectedValue) - 1].Value = "";
		}
	}

	public static HTKKCMCTableInfo getDataOfTable(HTKKCMCTableInfo fullData, int pageIndex)
	{
		int pageSizeNumberLarge = getPageSizeNumberLarge(fullData.RowsValues.Count);
		int num = pageIndex * pageSizeNumberLarge;
		int num2 = Math.Min(num + pageSizeNumberLarge, fullData.RowsValues.Count);
		HTKKCMCTableInfo hTKKCMCTableInfo = fullData.CloneWithoutData();
		for (int i = num; i < num2; i++)
		{
			hTKKCMCTableInfo.RowsValues.Add(fullData.RowsValues[i]);
		}
		return hTKKCMCTableInfo;
	}

	public static void fomartCellStyleOfTable1(FpSpread winFpSpread, int sheetIndex, HTKKCMCTableInfo table, string nameColEnd)
	{
		try
		{
			TextCellType textCellType = new TextCellType();
			winFpSpread.SuspendLayout();
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			if (table.RowsValues.Count <= 1)
			{
				return;
			}
			foreach (HTKKCMCTableCellInfo cell3 in table.RowsInfo.Cells)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				int num = dataLocationTemplateRowStart + table.RowsValues.Count - 1;
				Cell cell = sheetView.Cells[dataLocationTemplateRowStart - 1, cell3.CellTemplate.ColIndex - 1];
				if ((object)cell.CellType.GetType() == textCellType.GetType())
				{
					TextCellType textCellType2 = (TextCellType)cell.CellType;
					textCellType2.WordWrap = true;
					textCellType2.Multiline = true;
				}
				sheetView.Cells[dataLocationTemplateRowStart - 1, cell3.CellTemplate.ColIndex - 1, num - 1, cell3.CellTemplate.ColIndex - 1].CellType = cell.CellType;
				sheetView.Cells[dataLocationTemplateRowStart - 1, cell3.CellTemplate.ColIndex - 1, num - 1, cell3.CellTemplate.ColIndex - 1].Formatter = cell.Formatter;
				sheetView.Cells[dataLocationTemplateRowStart - 1, cell3.CellTemplate.ColIndex - 1, num - 1, cell3.CellTemplate.ColIndex - 1].Locked = cell.Locked;
				sheetView.Cells[dataLocationTemplateRowStart - 1, cell3.CellTemplate.ColIndex - 1, num - 1, cell3.CellTemplate.ColIndex - 1].BackColor = cell.BackColor;
				sheetView.Cells[dataLocationTemplateRowStart - 1, cell3.CellTemplate.ColIndex - 1, num - 1, cell3.CellTemplate.ColIndex - 1].ForeColor = cell.ForeColor;
				sheetView.Cells[dataLocationTemplateRowStart - 1, cell3.CellTemplate.ColIndex - 1, num - 1, cell3.CellTemplate.ColIndex - 1].HorizontalAlignment = cell.HorizontalAlignment;
				sheetView.Cells[dataLocationTemplateRowStart - 1, cell3.CellTemplate.ColIndex - 1, num - 1, cell3.CellTemplate.ColIndex - 1].Font = cell.Font;
				sheetView.Cells[dataLocationTemplateRowStart - 1, cell3.CellTemplate.ColIndex - 1, num - 1, cell3.CellTemplate.ColIndex - 1].CellPadding = cell.CellPadding;
				sheetView.Cells[dataLocationTemplateRowStart - 1, cell3.CellTemplate.ColIndex - 1, num - 1, cell3.CellTemplate.ColIndex - 1].Border = cell.Border;
			}
			int dataLocationTemplateRowStart2 = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				float preferredRowHeight = sheetView.GetPreferredRowHeight(dataLocationTemplateRowStart2 + i - 1);
				SetRowHeight(sheetView, dataLocationTemplateRowStart2 + i, preferredRowHeight);
			}
			for (int j = 0; j < table.RowsInfo.Cells[0].CellTemplate.ColIndex - 1; j++)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				int num = dataLocationTemplateRowStart + table.RowsValues.Count;
				Cell cell = sheetView.Cells[dataLocationTemplateRowStart - 1, j];
				NamedStyle namedStyle = new NamedStyle("RowStyle1");
				namedStyle.BackColor = cell.BackColor;
				namedStyle.Locked = cell.Locked;
				winFpSpread.NamedStyles.Add(namedStyle);
				sheetView.Cells[dataLocationTemplateRowStart - 1, j, num - 1, j].StyleName = "RowStyle1";
			}
			int num2 = HTKKFile.ColumnLettersToIndex(nameColEnd);
			int dataLocationTemplateRowStart3 = table.Location.DataLocationTemplateRowStart;
			int num3 = dataLocationTemplateRowStart3 + table.RowsValues.Count;
			Cell cell2 = sheetView.Cells[dataLocationTemplateRowStart3 - 1, num2 - 1];
			NamedStyle namedStyle2 = new NamedStyle("RowStyle2");
			namedStyle2.BackColor = cell2.BackColor;
			namedStyle2.Locked = cell2.Locked;
			namedStyle2.Border = cell2.Border;
			winFpSpread.NamedStyles.Add(namedStyle2);
			sheetView.Cells[dataLocationTemplateRowStart3 - 1, num2 - 1, num3 - 1, num2 - 1].StyleName = "RowStyle2";
			winFpSpread.ResumeLayout();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ReindexColStt(HTKKCMCTableInfo table)
	{
		try
		{
			int count = table.RowsValues.Count;
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + count;
			for (int i = 0; i < count; i++)
			{
				table.RowsValues[i].RowValue[0] = Convert.ToString(i + 1);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetNodeBackGroundPink(SheetView sView, int rowIndex, int colIndex, string msg, string rowSTT, string table)
	{
		try
		{
			sView.Cells[rowIndex, colIndex].BackColor = Color.Pink;
			sView.SetNote(rowIndex, colIndex, msg);
			sView.Cells[rowIndex, colIndex].NoteStyle = NoteStyle.PopupNote;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetNodeBackGroundYellow(SheetView sView, int rowIndex, int colIndex, string msg, string rowSTT, string table)
	{
		try
		{
			sView.Cells[rowIndex, colIndex].BackColor = Color.Yellow;
			sView.SetNote(rowIndex, colIndex, msg);
			sView.Cells[rowIndex, colIndex].NoteStyle = NoteStyle.PopupNote;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void LogErrorBigData(string sheetName, string tableName, string msg, string rowSTT)
	{
		try
		{
			string message = $"[{sheetName} {tableName}: STT {rowSTT}] {msg}";
			HTKKFile.LogErrorBigData(message);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ResetNoteOfTable(SheetView sView, HTKKCMCTableInfo table)
	{
		try
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			List<HTKKCMCTableCellInfo> cells = table.RowsInfo.Cells;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				foreach (HTKKCMCTableCellInfo item in cells)
				{
					Cell cell = sView.Cells[dataLocationTemplateRowStart + i - 1, item.CellTemplate.ColIndex - 1];
					if (cell != null && cell.NoteStyle == NoteStyle.PopupNote && !cell.Locked)
					{
						HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + i - 1, item.CellTemplate.ColIndex - 1);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string RemoveSpecialCharactersUnicode(string input)
	{
		try
		{
			string pattern = "[^\\p{L}\\p{N}\\s\\.,\\-_+=#$@!/*?{}()]";
			return Regex.Replace(input, pattern, "").Trim();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return input;
		}
	}

	public static int getPageSizeNumberLarge(int rowOfTable)
	{
		if (rowOfTable > 50)
		{
			return 50;
		}
		return 50;
	}

	public static void showPageLargeData1(FpSpread winFpSpread, int sheetIndex, HTKKCMCSheetInfo sInfo, int page, string nameColEnd, string menuId = "")
	{
		try
		{
			winFpSpread.SuspendLayout();
			SheetView sView = winFpSpread.Sheets[sheetIndex];
			if (dataInfoOld != null)
			{
				foreach (HTKKCMCBigTable item in dataInfoOld.listTable)
				{
					ResetNoteOfTable(sView, item.table);
				}
			}
			int offsetOfTable = getOffsetOfTable(sInfo);
			int num = 0;
			if (menuId == "P2")
			{
				num += 4;
			}
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
				num = ((i != 0) ? (num + (hTKKCMCTableInfo.RowsValues.Count + offsetOfTable)) : (num + hTKKCMCTableInfo.RowsValues.Count));
			}
			int pageSizeNumberLarge = getPageSizeNumberLarge(num);
			HTKKCMCBigDataInfo dataOfTable = getDataOfTable1(sInfo, page, pageSizeNumberLarge, menuId);
			hideShowSViewDynamic(sView, dataInfoOld, "true");
			hideShowFooterSViewStatic(sView, dataInfoOld, "true");
			bool flag = removeRowSView(sView, sInfo, dataOfTable, dataInfoOld);
			if (flag)
			{
				insertRowSView(sView, dataOfTable.listTable);
			}
			foreach (HTKKCMCBigTable item2 in dataOfTable.listTable)
			{
				BindDataToSviewFromTable2(sView, item2.table);
			}
			if (flag)
			{
				foreach (HTKKCMCBigTable item3 in dataOfTable.listTable)
				{
					fomartCellStyleOfTable1(winFpSpread, sheetIndex, item3.table, nameColEnd);
				}
			}
			hideShowSViewDynamic(sView, dataOfTable, "");
			hideShowFooterSViewStatic(sView, dataOfTable, "");
			bindDataStatic(sView, dataOfTable);
			dataInfoOld = dataOfTable;
			winFpSpread.ResumeLayout();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void BindDataToSviewFromTable2(SheetView sView, HTKKCMCTableInfo table)
	{
		try
		{
			if (table.Header.StaticCells.Count > 0)
			{
				for (int i = 0; i < table.Header.StaticCells.Count; i++)
				{
					fpSpBindDataCell(sView, table.Header.StaticCells[i].CellTemplate.RowIndex, table.Header.StaticCells[i].CellTemplate.ColIndex, table.Header.StaticCells[i].Controltype, table.Header.StaticCells[i].Value);
				}
			}
			BindDataLargeSetValueByCols(sView, table);
			if (table.Footer.StaticCells.Count > 0)
			{
				for (int i = 0; i < table.Footer.StaticCells.Count; i++)
				{
					fpSpBindDataCell(sView, table.Footer.StaticCells[i].CellTemplate.RowIndex, table.Footer.StaticCells[i].CellTemplate.ColIndex, table.Footer.StaticCells[i].Controltype, table.Footer.StaticCells[i].Value);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void BindDataLargeSetValueByCols1(SheetView sView, HTKKCMCTableInfo table, int startRow)
	{
		try
		{
			if (table == null || table.RowsValues.Count == 0)
			{
				return;
			}
			int count = table.RowsValues.Count;
			List<object[,]> list = new List<object[,]>();
			for (int i = 0; i < table.RowsInfo.Cells.Count; i++)
			{
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells[i];
				object[,] array = new object[count, 1];
				for (int j = 0; j < count; j++)
				{
					string text = table.RowsValues[j].RowValue[i].Trim();
					if (hTKKCMCTableCellInfo.Controltype == "16")
					{
						array[j, 0] = Convert.ToDouble(text.Replace(".", "").Replace(",", ""));
					}
					else if (hTKKCMCTableCellInfo.Controltype == "2")
					{
						array[j, 0] = ((text.ToLower() == "1" || text.ToLower() == "true" || text.ToLower() == "x") ? true : false);
					}
					else
					{
						array[j, 0] = text;
					}
				}
				if (((Control)sView.FpSpread).InvokeRequired)
				{
					sView.SetArray(startRow - 1, table.RowsInfo.Cells[i].CellTemplate.ColIndex - 1, array);
				}
				else
				{
					sView.SetArray(startRow - 1, table.RowsInfo.Cells[i].CellTemplate.ColIndex - 1, array);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static HTKKCMCBigDataInfo getDataOfTable1(HTKKCMCSheetInfo sInfo, int page, int pageSize, string menuId)
	{
		HTKKCMCBigDataInfo hTKKCMCBigDataInfo = new HTKKCMCBigDataInfo();
		List<HTKKCMCBigTable> list = new List<HTKKCMCBigTable>();
		int startRow = page * pageSize;
		int num = timBangBatDau(sInfo, startRow);
		if (num < 0)
		{
			return hTKKCMCBigDataInfo;
		}
		int num2 = timHangBatDau(sInfo, num, startRow);
		if (num2 < 0)
		{
			return hTKKCMCBigDataInfo;
		}
		int num3 = 0;
		for (int i = 0; i < sInfo.Tables.Count; i++)
		{
			if (i < num)
			{
				HTKKCMCBigTable hTKKCMCBigTable = new HTKKCMCBigTable();
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i].CloneWithoutData();
				hTKKCMCTableInfo.RowsValues.Clear();
				hTKKCMCTableInfo.RowsValues.Add(GetRowValueDefault(hTKKCMCTableInfo.RowsInfo));
				hTKKCMCBigTable.table = hTKKCMCTableInfo;
				hTKKCMCBigTable.tableIndex = i;
				hTKKCMCBigTable.isShow = false;
				list.Add(hTKKCMCBigTable);
			}
			else if (num3 < pageSize)
			{
				HTKKCMCBigTable hTKKCMCBigTable = new HTKKCMCBigTable();
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i].CloneWithoutData();
				for (int j = num2; j < sInfo.Tables[i].RowsValues.Count; j++)
				{
					hTKKCMCTableInfo.RowsValues.Add(sInfo.Tables[i].RowsValues[j]);
					num3++;
					if (num3 >= pageSize)
					{
						break;
					}
				}
				hTKKCMCBigTable.table = hTKKCMCTableInfo;
				hTKKCMCBigTable.tableIndex = i;
				hTKKCMCBigTable.isShow = true;
				list.Add(hTKKCMCBigTable);
				num2 = 0;
			}
			else
			{
				HTKKCMCBigTable hTKKCMCBigTable = new HTKKCMCBigTable();
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i].CloneWithoutData();
				hTKKCMCTableInfo.RowsValues.Clear();
				hTKKCMCTableInfo.RowsValues.Add(GetRowValueDefault(hTKKCMCTableInfo.RowsInfo));
				hTKKCMCBigTable.table = hTKKCMCTableInfo;
				hTKKCMCBigTable.tableIndex = i;
				hTKKCMCBigTable.isShow = false;
				list.Add(hTKKCMCBigTable);
			}
		}
		foreach (HTKKCMCBigTable item in list)
		{
			if (item.table.RowsValues.Count == 0)
			{
				item.isShow = false;
				item.table.RowsValues.Add(GetRowValueDefault(item.table.RowsInfo));
			}
		}
		hTKKCMCBigDataInfo.listTable = list;
		if (num3 >= pageSize)
		{
			hTKKCMCBigDataInfo.showFooter = false;
		}
		else
		{
			hTKKCMCBigDataInfo.showFooter = true;
		}
		int dataLocationTemplateRowStart = sInfo.Tables[0].Location.DataLocationTemplateRowStart;
		int offsetOfTable = getOffsetOfTable(sInfo);
		for (int i = 0; i < list.Count; i++)
		{
			if (i == 0)
			{
				list[i].table.Location.DataLocationTemplateRowStart = dataLocationTemplateRowStart;
				continue;
			}
			if (list[i - 1].table.RowsValues.Count > 1 && offsetOfTable >= 0)
			{
				list[i].table.Location.DataLocationTemplateRowStart = dataLocationTemplateRowStart + list[i - 1].table.RowsValues.Count + offsetOfTable;
				continue;
			}
			if (list[i].table.RowsValues.Count > 1 && offsetOfTable >= 0)
			{
				list[i].table.Location.DataLocationTemplateRowStart = dataLocationTemplateRowStart + list[i - 1].table.RowsValues.Count + offsetOfTable;
			}
			if (list[i - 1].table.RowsValues.Count == 1 && offsetOfTable >= 0)
			{
				list[i].table.Location.DataLocationTemplateRowStart = dataLocationTemplateRowStart + list[i - 1].table.RowsValues.Count + offsetOfTable;
			}
		}
		if (sInfo.StaticCells.Count > 0)
		{
			num3 = 0;
			foreach (HTKKCMCBigTable item2 in list)
			{
				num3 += item2.table.RowsValues.Count;
			}
			int num4 = offsetOfTable;
			if (list.Count > 1)
			{
				num4 = offsetOfTable * (list.Count - 1) + 1;
			}
			if (menuId == "U8")
			{
				num4 += 2;
			}
			List<HTKKCMCCellInfo> list2 = sInfo.CloneStaticCells();
			int num5 = 0;
			foreach (HTKKCMCCellInfo item3 in list2)
			{
				if (item3.CellTemplate.RowIndex > list[0].table.Location.DataLocationTemplateRowStart)
				{
					break;
				}
				num5++;
			}
			if (num5 < list2.Count)
			{
				int rowIndex = list2[num5].CellTemplate.RowIndex;
				int num6 = rowIndex - (dataLocationTemplateRowStart + num3 + num4);
				if (num6 > 0)
				{
					foreach (HTKKCMCCellInfo item4 in list2)
					{
						if (item4.CellTemplate.RowIndex > list[0].table.Location.DataLocationTemplateRowStart)
						{
							item4.CellTemplate.RowIndex = item4.CellTemplate.RowIndex - num6;
						}
					}
				}
			}
			hTKKCMCBigDataInfo.staticCells = list2;
		}
		return hTKKCMCBigDataInfo;
	}

	public static int timHangBatDau(HTKKCMCSheetInfo sInfo, int startTable, int startRow)
	{
		int num = 0;
		for (int i = 0; i < sInfo.Tables.Count; i++)
		{
			if (i < startTable)
			{
				num += sInfo.Tables[i].RowsValues.Count;
				continue;
			}
			return startRow - num;
		}
		return -1;
	}

	public static int timBangBatDau(HTKKCMCSheetInfo sInfo, int startRow)
	{
		int num = 0;
		for (int i = 0; i < sInfo.Tables.Count; i++)
		{
			num += sInfo.Tables[i].RowsValues.Count;
			if (num >= startRow)
			{
				return i;
			}
		}
		return -1;
	}

	public static int timBangBatDauSView(HTKKCMCSheetInfo sInfo, int startRowSView)
	{
		for (int i = 0; i < sInfo.Tables.Count; i++)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count;
			if (startRowSView >= dataLocationTemplateRowStart && startRowSView <= num)
			{
				return i;
			}
		}
		return -1;
	}

	public static bool removeRowSView(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKCMCBigDataInfo dataInfoNow, HTKKCMCBigDataInfo dataInfoOld)
	{
		bool flag = false;
		try
		{
			if (dataInfoOld == null || dataInfoOld.listTable.Count == 0)
			{
				return true;
			}
			if (!flag && dataInfoOld.showFooter != dataInfoNow.showFooter)
			{
				flag = true;
			}
			if (!flag && dataInfoOld.listTable.Count != dataInfoNow.listTable.Count)
			{
				flag = true;
			}
			if (!flag)
			{
				foreach (HTKKCMCBigTable item in dataInfoNow.listTable)
				{
					string tablePath = item.table.TablePath;
					HTKKCMCBigTable hTKKCMCBigTable = dataInfoOld.listTable.Find((HTKKCMCBigTable x) => x.table.TablePath == tablePath);
					if (hTKKCMCBigTable == null || item.table.Location.DataLocationTemplateRowStart != hTKKCMCBigTable.table.Location.DataLocationTemplateRowStart || item.table.RowsValues.Count != hTKKCMCBigTable.table.RowsValues.Count)
					{
						flag = true;
						break;
					}
				}
			}
			if (!flag)
			{
				return flag;
			}
			for (int num = dataInfoOld.listTable.Count - 1; num >= 0; num--)
			{
				HTKKCMCTableInfo table = dataInfoOld.listTable[num].table;
				if (table.RowsValues.Count > 1)
				{
					sView.Rows.Remove(table.Location.DataLocationTemplateRowStart, table.RowsValues.Count - 1);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return flag;
	}

	public static int getOffsetOfTable(HTKKCMCSheetInfo sInfo)
	{
		int num = 0;
		try
		{
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				if (i > 0)
				{
					num += sInfo.Tables[i].Location.DataLocationTemplateRowStart - (sInfo.Tables[i - 1].Location.DataLocationTemplateRowStart + sInfo.Tables[i - 1].RowsValues.Count);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return num;
	}

	public static void insertRowSView(SheetView sView, List<HTKKCMCBigTable> listTable)
	{
		try
		{
			for (int i = 0; i < listTable.Count; i++)
			{
				HTKKCMCTableInfo table = listTable[i].table;
				if (table.RowsValues.Count > 1)
				{
					sView.Rows.Add(table.Location.DataLocationTemplateRowStart, table.RowsValues.Count - 1);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void hideShowSViewDynamic(SheetView sView, HTKKCMCBigDataInfo dataInfo, string isShow)
	{
		try
		{
			if (dataInfo == null || dataInfo.listTable.Count == 0)
			{
				return;
			}
			foreach (HTKKCMCBigTable item in dataInfo.listTable)
			{
				HTKKCMCTableInfo table = item.table;
				if (!string.IsNullOrEmpty(isShow))
				{
					sView.Rows[table.Location.DataLocationTemplateRowStart - 3].Visible = isShow == "true";
					sView.Rows[table.Location.DataLocationTemplateRowStart - 1].Visible = isShow == "true";
					sView.Rows[table.Location.DataLocationTemplateRowStart - 2].Visible = isShow == "true";
					sView.Rows[table.Location.DataLocationTemplateRowStart + (table.RowsValues.Count - 1)].Visible = isShow == "true";
				}
				else
				{
					sView.Rows[table.Location.DataLocationTemplateRowStart - 3].Visible = (item.isShow ? true : false);
					sView.Rows[table.Location.DataLocationTemplateRowStart - 1].Visible = (item.isShow ? true : false);
					sView.Rows[table.Location.DataLocationTemplateRowStart - 2].Visible = false;
					sView.Rows[table.Location.DataLocationTemplateRowStart + (table.RowsValues.Count - 1)].Visible = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void hideShowFooterSViewStatic(SheetView sView, HTKKCMCBigDataInfo dataInfo, string isShow)
	{
		try
		{
			if (dataInfo == null || dataInfo.staticCells.Count == 0)
			{
				return;
			}
			foreach (HTKKCMCCellInfo staticCell in dataInfo.staticCells)
			{
				List<HTKKCMCBigTable> listTable = dataInfo.listTable;
				if (listTable.Count <= 0 || staticCell.CellTemplate.RowIndex >= listTable[0].table.Location.DataLocationTemplateRowStart)
				{
					if (!string.IsNullOrEmpty(isShow))
					{
						sView.Rows[staticCell.CellTemplate.RowIndex - 1].Visible = isShow == "true";
					}
					else
					{
						sView.Rows[staticCell.CellTemplate.RowIndex - 1].Visible = (dataInfo.showFooter ? true : false);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void bindDataStatic(SheetView sView, HTKKCMCBigDataInfo dataInfo)
	{
		try
		{
			foreach (HTKKCMCCellInfo staticCell in dataInfo.staticCells)
			{
				sView.Cells[staticCell.CellTemplate.RowIndex - 1, staticCell.CellTemplate.ColIndex - 1].Value = staticCell.Value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static int rowCountSum(HTKKCMCSheetInfo sInfo)
	{
		int num = -1;
		try
		{
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				num += table.RowsValues.Count;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1;
		}
		return num;
	}

	public static Dictionary<string, string[]> getTableFromSInfo(HTKKCMCSheetInfo sInfo)
	{
		Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
		try
		{
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				string[] array = table.TablePath.Split(new char[1] { '/' });
				if (array.Length > 1)
				{
					dictionary.Add(array[1], array);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return dictionary;
	}

	public static string GetTextByDateRange(DateTime fromDate, DateTime toDate)
	{
		List<DateRangeText> list = new List<DateRangeText>();
		list.Add(new DateRangeText(new DateTime(2024, 7, 1), new DateTime(2024, 12, 31), "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 142/2024/QH15"));
		list.Add(new DateRangeText(new DateTime(2025, 1, 1), new DateTime(2025, 6, 30), "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 174/2024/QH15"));
		list.Add(new DateRangeText(new DateTime(2025, 7, 1), new DateTime(2026, 12, 31), "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 204/2025/QH15"));
		List<string> list2 = new List<string>();
		foreach (DateRangeText item in list)
		{
			if (IsOverlap(fromDate, toDate, item.Start, item.End))
			{
				list2.Add(item.Text);
			}
		}
		int count = list2.Count;
		switch (count)
		{
		case 0:
			return "";
		case 1:
			return list2[0];
		case 2:
			return list2[0] + " VÀ " + list2[1];
		default:
		{
			string text = "";
			for (int i = 0; i < count - 1; i++)
			{
				text = text + list2[i] + " VÀ ";
			}
			return text + "VÀ " + list2[count - 1];
		}
		}
	}

	private static bool IsOverlap(DateTime start1, DateTime end1, DateTime start2, DateTime end2)
	{
		return start1 <= end2 && start2 <= end1;
	}

	public static void lockUnlockMultiCellsStatic(SheetView sView, HTKKCMCSheetInfo sInfo, int startRow, int endRow, bool isLock, bool isDeleteData)
	{
		foreach (HTKKCMCCellInfo staticCell in sInfo.StaticCells)
		{
			if (staticCell.CellTemplate.RowIndex >= startRow && staticCell.CellTemplate.RowIndex <= endRow)
			{
				sView.Cells[staticCell.CellTemplate.RowIndex - 1, staticCell.CellTemplate.ColIndex - 1].Locked = isLock;
				if (isDeleteData)
				{
					sView.Cells[staticCell.CellTemplate.RowIndex - 1, staticCell.CellTemplate.ColIndex - 1].Value = staticCell.DefaultValue;
				}
				sView.Cells[staticCell.CellTemplate.RowIndex - 1, staticCell.CellTemplate.ColIndex - 1].BackColor = (isLock ? ColorTranslator.FromHtml(MA_COLOR_LOCKED) : ColorTranslator.FromHtml(MA_COLOR_UNLOCK));
			}
		}
	}

	public static void resetOtherCheckboxStatic(SheetView sView, HTKKCMCSheetInfo sInfo, int startRow, int endRow, int rowClick, int colClick)
	{
		bool flag = false;
		object value = sView.Cells[rowClick - 1, colClick - 1].Value;
		if (value != null && !string.IsNullOrEmpty(value.ToString()) && (value.ToString().ToLower() == "true" || value.ToString().ToLower() == "false"))
		{
			flag = Convert.ToBoolean(value);
		}
		bool flag2 = false;
		foreach (HTKKCMCCellInfo staticCell in sInfo.StaticCells)
		{
			if (staticCell.CellTemplate.RowIndex >= startRow && staticCell.CellTemplate.RowIndex <= endRow && staticCell.CellTemplate.RowIndex == rowClick && staticCell.CellTemplate.ColIndex == colClick)
			{
				flag2 = true;
				break;
			}
		}
		if (!flag || !flag2)
		{
			return;
		}
		foreach (HTKKCMCCellInfo staticCell2 in sInfo.StaticCells)
		{
			if (staticCell2.CellTemplate.RowIndex >= startRow && staticCell2.CellTemplate.RowIndex <= endRow && (staticCell2.CellTemplate.RowIndex != rowClick || staticCell2.CellTemplate.ColIndex != colClick || !(staticCell2.Controltype == "2")) && staticCell2.Controltype == "2")
			{
				sView.Cells[staticCell2.CellTemplate.RowIndex - 1, staticCell2.CellTemplate.ColIndex - 1].Value = false;
			}
		}
	}

	public static void showRowsErrorLarge(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			sView.FpSpread.SuspendLayout();
			SheetTablePage sheetTablePage = BuildSheetTablePage(sInfo, _page);
			foreach (SheetTablePageSegment segment in sheetTablePage.Segments)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[segment.TableIndex];
				segment.TargetDataRow = MapTemplateExpandedRow(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, sInfo, sheetTablePage);
			}
			RemoveSheetTablePageOldWarnings(sView, sheetTablePage);
			foreach (SheetTablePageSegment segment2 in sheetTablePage.Segments)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[segment2.TableIndex];
				List<HTKKCMCTableCellInfo> cells = hTKKCMCTableInfo.RowsInfo.Cells;
				for (int i = 0; i < segment2.DataCount; i++)
				{
					int index = segment2.FirstDataIndex + i;
					HTKKCMCRowValue hTKKCMCRowValue = hTKKCMCTableInfo.RowsValues[index];
					if (hTKKCMCRowValue.RowValue.Count < cells.Count + 2)
					{
						continue;
					}
					int num = cells.Count;
					while (num < 50 && hTKKCMCRowValue.RowValue.Count > num)
					{
						num++;
						int num2 = Convert.ToInt32(hTKKCMCRowValue.RowValue[num]);
						num++;
						string text = hTKKCMCRowValue.RowValue[num];
						num++;
						string value = hTKKCMCRowValue.RowValue[num];
						int num3 = -1;
						if (!string.IsNullOrEmpty(value))
						{
							num3 = Convert.ToInt32(value);
						}
						if (!string.IsNullOrEmpty(text))
						{
							if (num3 == 2)
							{
								SetNodeBackGroundPink(sView, segment2.TargetDataRow + i - 1, num2 - 1, text, hTKKCMCRowValue.RowValue[0], (segment2.TableIndex == 0) ? "Mục I" : "Mục II");
							}
							else
							{
								SetNodeBackGroundYellow(sView, segment2.TargetDataRow + i - 1, num2 - 1, text, hTKKCMCRowValue.RowValue[0], (segment2.TableIndex == 0) ? "Mục I" : "Mục II");
							}
						}
						num++;
					}
				}
			}
			sView.FpSpread.ResumeLayout();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void showRowsErrorLarge(SheetView sView)
	{
		try
		{
			sView.FpSpread.SuspendLayout();
			int num = 0;
			foreach (HTKKCMCBigTable item in dataInfoOld.listTable)
			{
				HTKKCMCTableInfo table = item.table;
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				List<HTKKCMCTableCellInfo> cells = table.RowsInfo.Cells;
				for (int i = 0; i < table.RowsValues.Count; i++)
				{
					HTKKCMCRowValue hTKKCMCRowValue = table.RowsValues[i];
					if (hTKKCMCRowValue.RowValue.Count < cells.Count + 2)
					{
						continue;
					}
					int num2 = cells.Count;
					while (num2 < 50 && hTKKCMCRowValue.RowValue.Count > num2)
					{
						num2++;
						int num3 = Convert.ToInt32(hTKKCMCRowValue.RowValue[num2]);
						num2++;
						string text = hTKKCMCRowValue.RowValue[num2];
						num2++;
						string value = hTKKCMCRowValue.RowValue[num2];
						int num4 = -1;
						if (!string.IsNullOrEmpty(value))
						{
							num4 = Convert.ToInt32(value);
						}
						if (!string.IsNullOrEmpty(text))
						{
							if (num4 == 2)
							{
								SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, num3 - 1, text, hTKKCMCRowValue.RowValue[0], (num == 0) ? "Mục I" : "Mục II");
							}
							else
							{
								SetNodeBackGroundYellow(sView, dataLocationTemplateRowStart + i - 1, num3 - 1, text, hTKKCMCRowValue.RowValue[0], (num == 0) ? "Mục I" : "Mục II");
							}
						}
						num2++;
					}
				}
				num++;
			}
			sView.FpSpread.ResumeLayout();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void IncrementTemplateRowStartInTable(int frRow, int rowCount, HTKKCMCBigDataInfo dataInfo)
	{
		try
		{
			foreach (HTKKCMCBigTable item in dataInfo.listTable)
			{
				HTKKCMCTableInfo table = item.table;
				if (table.Location.DataLocationTemplateRowStart > frRow)
				{
					table.Increment(rowCount);
				}
				table.IncrementFooter(frRow, rowCount);
			}
			foreach (HTKKCMCCellInfo staticCell in dataInfo.staticCells)
			{
				if (staticCell.CellTemplate.RowIndex > frRow)
				{
					staticCell.Increment(rowCount);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void RemoveDataRowsNoSViewBigData(int frRow, int rowCount, HTKKCMCBigDataInfo dataInfo)
	{
		try
		{
			foreach (HTKKCMCBigTable item in dataInfo.listTable)
			{
				HTKKCMCTableInfo table = item.table;
				if (table.Location.DataLocationTemplateRowStart > frRow + 1)
				{
					table.Increment(rowCount * -1);
				}
				if (table.Footer != null && table.Footer.StaticCells.Count > 0)
				{
					table.IncrementFooter(frRow, rowCount * -1);
				}
			}
			foreach (HTKKCMCCellInfo staticCell in dataInfo.staticCells)
			{
				if (staticCell.CellTemplate.RowIndex > frRow)
				{
					staticCell.Increment(rowCount * -1);
				}
			}
		}
		catch (Exception)
		{
		}
	}

	public static bool UpdateFromSViewToSInfoLarge(SheetView sView, int row, int col, HTKKCMCRowInfo rowInfo, HTKKCMCSheetInfo sInfo, int tableIndex)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[tableIndex];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int rowStartSView = _page * _pageSize + row;
			int num = sInfo.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == rowStartSView);
			if (num >= 0)
			{
				return false;
			}
			int num2 = 0;
			if (tableIndex > 0 && !dataInfoOld.listTable[tableIndex - 1].isShow)
			{
				num2++;
			}
			int index = rowStartSView - (hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + num2);
			int num3 = 0;
			if (col >= 0)
			{
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo = hTKKCMCTableInfo.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == col);
				num3 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == col);
				if (hTKKCMCTableCellInfo == null || num3 < 0)
				{
					return true;
				}
				object obj = sView.Cells[row - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Value;
				if (hTKKCMCTableCellInfo.Controltype == "16")
				{
					if (obj != null && !string.IsNullOrEmpty(obj.ToString()) && obj.ToString().EndsWith(".0"))
					{
						obj = obj.ToString().Substring(0, obj.ToString().Length - 2);
					}
					hTKKCMCTableInfo.RowsValues[index].RowValue[num3] = ((obj == null) ? "0" : obj.ToString());
				}
				else if (hTKKCMCTableCellInfo.Controltype == "2")
				{
					hTKKCMCTableInfo.RowsValues[index].RowValue[num3] = ((obj != null && Convert.ToBoolean(obj)) ? "1" : "0");
				}
				else
				{
					hTKKCMCTableInfo.RowsValues[index].RowValue[num3] = ((obj != null) ? obj.ToString() : "");
				}
			}
			else
			{
				foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo.RowsInfo.Cells)
				{
					object obj = sView.Cells[row - 1, cell.CellTemplate.ColIndex - 1].Value;
					if (cell.Controltype == "16")
					{
						if (obj != null && !string.IsNullOrEmpty(obj.ToString()) && obj.ToString().EndsWith(".0"))
						{
							obj = obj.ToString().Substring(0, obj.ToString().Length - 2);
						}
						hTKKCMCTableInfo.RowsValues[index].RowValue[num3] = ((obj == null) ? "0" : obj.ToString());
					}
					else if (cell.Controltype == "2")
					{
						hTKKCMCTableInfo.RowsValues[index].RowValue[num3] = ((obj != null && Convert.ToBoolean(obj)) ? "1" : "0");
					}
					else
					{
						hTKKCMCTableInfo.RowsValues[index].RowValue[num3] = ((obj != null) ? obj.ToString() : "");
					}
					num3++;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return true;
	}

	public static void UpdateFromSViewToSInfoStaticLarge(SheetView sView, int row, int col, HTKKCMCSheetInfo sInfo)
	{
		object value = sView.Cells[row - 1, col - 1].Value;
		int rowStartSView = _page * _pageSize + row;
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == rowStartSView && x.CellTemplate.ColIndex == col);
		if (hTKKCMCCellInfo != null)
		{
			hTKKCMCCellInfo.Value = ((value == null) ? "" : value.ToString());
		}
	}

	public static int tinhRowCount_Large(HTKKCMCSheetInfo sInfo, string menuId)
	{
		int num = rowCountSum(sInfo);
		if (menuId == "P2")
		{
			num = (dataInfoOld.showFooter ? (num + 1) : (num + 4));
		}
		return num;
	}

	public static int tinhPageSize_Large(HTKKCMCSheetInfo sInfo, string menuId)
	{
		int rowOfTable = tinhRowCount_Large(sInfo, menuId);
		return getPageSizeNumberLarge(rowOfTable);
	}

	public static int tinhTotalPage_Large(HTKKCMCSheetInfo sInfo, string menuId)
	{
		int num = tinhRowCount_Large(sInfo, menuId);
		int pageSizeNumberLarge = getPageSizeNumberLarge(num);
		int num2 = (int)Math.Ceiling((double)num / (double)pageSizeNumberLarge);
		return (num2 == 0) ? 1 : num2;
	}

	public static void resetOtherCheckboxDinamic(SheetView sView, HTKKCMCSheetInfo sInfo, int rowClick, int colClick, List<int> colsCheckbox)
	{
		bool flag = false;
		object value = sView.Cells[rowClick - 1, colClick - 1].Value;
		if (value != null && !string.IsNullOrEmpty(value.ToString()) && (value.ToString().ToLower() == "true" || value.ToString().ToLower() == "false"))
		{
			flag = Convert.ToBoolean(value);
		}
		if (!flag)
		{
			return;
		}
		foreach (int item in colsCheckbox)
		{
			if (item != colClick)
			{
				sView.Cells[rowClick - 1, item - 1].Value = false;
			}
		}
	}

	public static void CopyTableToTable4(SheetView fromSView, SheetView toSView, HTKKCMCSheetInfo toSInfo, HTKKCMCTableInfo fromTable, int toTableIndex, Dictionary<int, int> dicInfo, bool deleteDataOld, ref Dictionary<string, int> checkExists, bool usingCellType = false)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = toSInfo.Tables[toTableIndex];
			int dataLocationTemplateRowStart = fromTable.Location.DataLocationTemplateRowStart;
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			if (deleteDataOld)
			{
				if (hTKKCMCTableInfo.RowsValues.Count > 1)
				{
					RemoveDataRows(toSView, toTableIndex, num - 1, hTKKCMCTableInfo.RowsValues.Count - 1, toSInfo);
				}
				foreach (HTKKCMCTableCellInfo cell2 in hTKKCMCTableInfo.RowsInfo.Cells)
				{
					toSView.Cells[num - 1, cell2.CellTemplate.ColIndex - 1].Value = cell2.DefaultValue;
				}
			}
			NumberCellType numberCellType = new NumberCellType();
			CurrencyCellType currencyCellType = new CurrencyCellType();
			TextCellType textCellType = new TextCellType();
			GeneralCellType generalCellType = new GeneralCellType();
			for (int i = 0; i < fromTable.RowsValues.Count; i++)
			{
				string text = "";
				foreach (int key in dicInfo.Keys)
				{
					int num2 = dicInfo[key];
					Cell cell = fromSView.Cells[dataLocationTemplateRowStart + i - 1, key - 1];
					ICellType cellType = fromSView.Cells[dataLocationTemplateRowStart + i - 1, key - 1].CellType;
					if (cell != null && (object)cellType.GetType() != numberCellType.GetType() && (object)cellType.GetType() != currencyCellType.GetType())
					{
						if (string.IsNullOrEmpty(cell.Text.Trim()))
						{
							text = "";
							break;
						}
						text += cell.Text.Trim();
					}
				}
				if (string.IsNullOrEmpty(text))
				{
					continue;
				}
				text = hTKKCMCTableInfo.TableName + hTKKCMCTableInfo.TableID + text;
				if (checkExists.ContainsKey(text))
				{
					num = checkExists[text];
				}
				else
				{
					object value = toSView.Cells[num - 1, dicInfo.Values.ToArray()[0] - 1].Value;
					if (value != null && !string.IsNullOrEmpty(value.ToString().Trim()))
					{
						num = checkExists.Last().Value;
						checkExists.Add(text, num + 1);
						InsertTableRow(toSView, num, -1, num, -1, -1, toSInfo, toTableIndex);
						num++;
					}
					else
					{
						checkExists.Add(text, num);
					}
				}
				foreach (int key2 in dicInfo.Keys)
				{
					int num2 = dicInfo[key2];
					object value2 = fromSView.Cells[dataLocationTemplateRowStart + i - 1, key2 - 1].Value;
					ICellType cellType = fromSView.Cells[dataLocationTemplateRowStart + i - 1, key2 - 1].CellType;
					if (value2 == null)
					{
						continue;
					}
					if ((object)cellType.GetType() == numberCellType.GetType() || (object)cellType.GetType() == currencyCellType.GetType())
					{
						if (usingCellType)
						{
							toSView.Cells[num - 1, num2 - 1].CellType = cellType;
						}
						toSView.Cells[num - 1, num2 - 1].Value = Convert.ToDouble(toSView.Cells[num - 1, num2 - 1].Value) + Convert.ToDouble(value2);
					}
					else
					{
						toSView.Cells[num - 1, num2 - 1].Value = value2;
					}
				}
			}
			fpSpReindexingTableTemplate(toSView, hTKKCMCTableInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool HasInvalidXmlChars(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		foreach (char ch in text)
		{
			if (!IsValidXmlChar(ch))
			{
				return true;
			}
		}
		return false;
	}

	private static bool IsValidXmlChar(char ch)
	{
		return ch == '\t' || ch == '\n' || ch == '\r' || (ch >= ' ' && ch <= '\ud7ff') || (ch >= '\ue000' && ch <= '\ufffd');
	}

	public static bool checkCCCD_CMND(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		if (!HTKKFormater.IsNumber(text))
		{
			return false;
		}
		if (text.Length == 9 || text.Length == 12)
		{
			return true;
		}
		return false;
	}

	public static bool checkCCCD(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		if (!HTKKFormater.IsNumber(text))
		{
			return false;
		}
		if (text.Length == 12)
		{
			return true;
		}
		return false;
	}

	public static bool IsValidTime(string timeStr)
	{
		if (string.IsNullOrEmpty(timeStr) || timeStr.Trim().Length == 0)
		{
			return false;
		}
		string format = "HH:mm:ss";
		DateTime result;
		return DateTime.TryParseExact(timeStr, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
	}

	public static bool IsValidDateTime(string dateTimeStr)
	{
		if (string.IsNullOrEmpty(dateTimeStr) || dateTimeStr.Trim().Length == 0)
		{
			return false;
		}
		string format = "dd/MM/yyyy HH:mm:ss";
		DateTime result;
		return DateTime.TryParseExact(dateTimeStr, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
	}

	public static int checkValidateTableEmpty(SheetView sView, HTKKCMCTableInfo table, List<int> colsCheck, string mess, List<int> colsRemoveValidate = null)
	{
		int result = 0;
		try
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				foreach (HTKKCMCTableCellInfo cell in table.RowsInfo.Cells)
				{
					if (!colsCheck.Contains(cell.CellTemplate.ColIndex))
					{
						continue;
					}
					if (colsRemoveValidate == null || !colsRemoveValidate.Contains(cell.CellTemplate.ColIndex))
					{
						HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1);
					}
					if (cell.Controltype == "16")
					{
						double num = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1].Value);
						if (num == Convert.ToDouble(cell.DefaultValue))
						{
							HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1, mess);
							result = 2;
						}
					}
					else
					{
						string text = sView.Cells[dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1].Text;
						if (text == cell.DefaultValue)
						{
							HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1, mess);
							result = 2;
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static List<MapCongThuc> MapCongThucChoSua(string IDTokhai)
	{
		//IL_001a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0020: Expected O, but got Unknown
		//IL_009c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a3: Expected O, but got Unknown
		List<MapCongThuc> list = new List<MapCongThuc>();
		try
		{
			XmlDocument val = new XmlDocument();
			val.Load("MapCongThucChoSua.xml");
			XmlNodeList source = ((XmlNode)val.DocumentElement).SelectNodes("/Root/Map");
			XmlNode val2 = ((IEnumerable)source).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["ID"]).Value == IDTokhai));
			if (val2 != null && val2.ChildNodes.Count > 0)
			{
				foreach (XmlNode childNode in val2.ChildNodes)
				{
					XmlNode val3 = childNode;
					MapCongThuc mapCongThuc = new MapCongThuc();
					mapCongThuc.Dynamic = Convert.ToInt32(((XmlNode)val3.Attributes["Dynamic"]).Value);
					mapCongThuc.Formula = ((XmlNode)val3.Attributes["Formula"]).Value;
					mapCongThuc.MCTPath = ((XmlNode)val3.Attributes["MCTPath"]).Value;
					mapCongThuc.SheetIndex = Convert.ToInt32(((XmlNode)val3.Attributes["SheetIndex"]).Value);
					mapCongThuc.TablePath = ((mapCongThuc.Dynamic == 0) ? "" : ((XmlNode)val3.Attributes["TablePath"]).Value);
					mapCongThuc.References = ((val3.OuterXml.IndexOf("References") >= 0) ? ((XmlNode)val3.Attributes["References"]).Value : "");
					list.Add(mapCongThuc);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	public static void CapNhatCongThucLoaiChoSuaCell(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol, List<MapCongThuc> listCongThuc)
	{
		string oldValue = "_ROWREPLACE";
		int tableIndex = DetectIndexOfTable(sInfo, iRow);
		try
		{
			if (tableIndex < 0)
			{
				List<HTKKCMCCellInfo> list = sInfo.StaticCells.FindAll((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == iRow);
				foreach (HTKKCMCCellInfo cell in list)
				{
					Predicate<MapCongThuc> match = (MapCongThuc x) => x.MCTPath == cell.Path && x.Dynamic == 0;
					MapCongThuc mapCongThuc = listCongThuc.Find(match);
					if (mapCongThuc != null && iCol < cell.CellTemplate.ColIndex)
					{
						string formula = mapCongThuc.Formula.Replace(oldValue, string.Concat(iRow));
						sView.Cells[cell.CellTemplate.RowIndex - 1, cell.CellTemplate.ColIndex - 1].Formula = formula;
					}
				}
				List<MapCongThuc> list2 = listCongThuc.FindAll((MapCongThuc x) => !string.IsNullOrEmpty(x.References) && x.References.Contains(iRow.ToString()));
				{
					foreach (MapCongThuc congThuc in list2)
					{
						List<HTKKCMCCellInfo> staticCells = sInfo.StaticCells;
						Predicate<HTKKCMCCellInfo> match2 = (HTKKCMCCellInfo x) => x.Path == congThuc.MCTPath;
						HTKKCMCCellInfo hTKKCMCCellInfo = staticCells.Find(match2);
						if (hTKKCMCCellInfo != null)
						{
							sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Formula = congThuc.Formula;
						}
					}
					return;
				}
			}
			List<MapCongThuc> list3 = listCongThuc.FindAll((MapCongThuc x) => x.TablePath == sInfo.Tables[tableIndex].TablePath);
			foreach (HTKKCMCTableCellInfo cell2 in sInfo.Tables[tableIndex].RowsInfo.Cells)
			{
				Predicate<MapCongThuc> match3 = (MapCongThuc x) => x.MCTPath == cell2.Path && x.Dynamic == 1;
				MapCongThuc mapCongThuc = list3.Find(match3);
				if (mapCongThuc != null && iCol < cell2.CellTemplate.ColIndex)
				{
					string formula = mapCongThuc.Formula.Replace(oldValue, string.Concat(iRow));
					sView.Cells[iRow - 1, cell2.CellTemplate.ColIndex - 1].Formula = formula;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static List<int> getAllColumnsInTable(HTKKCMCTableInfo table)
	{
		List<int> list = new List<int>();
		foreach (HTKKCMCTableCellInfo cell in table.RowsInfo.Cells)
		{
			list.Add(cell.CellTemplate.ColIndex);
		}
		return list;
	}

	public static int checkTrungGiaTri1CotTrongBang(SheetView sView, HTKKCMCTableInfo table, int colCheck, string mess, bool isRemoveValidate = false, int colNote = -1)
	{
		int result = 0;
		try
		{
			if (colNote < 0)
			{
				colNote = colCheck;
			}
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			List<object> list = new List<object>();
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				if (isRemoveValidate)
				{
					HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + i - 1, colNote - 1);
				}
				list.Add(sView.Cells[dataLocationTemplateRowStart + i - 1, colCheck - 1].Value);
			}
			var source = (from v in list
				where v != null
				group v by v into g
				where g.Count() > 1
				select new
				{
					Value = g.Key,
					Count = g.Count()
				}).ToList();
			if (source.Any())
			{
				HashSet<object> hashSet = new HashSet<object>(source.Select(d => d.Value));
				for (int i = 0; i < table.RowsValues.Count; i++)
				{
					int row = dataLocationTemplateRowStart + i - 1;
					object value = sView.Cells[row, colCheck - 1].Value;
					if (value != null && hashSet.Contains(value))
					{
						sView.Cells[row, colCheck - 1].BackColor = Color.LightPink;
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, colNote - 1, mess);
						result = 2;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static void TinhTongTheoTieuMuc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int colSoTien, int colMaTM, int colSoTienTo, int colMaTMTo, int colMucTo, int colTMTo, int colMaMucTo, int colStt, int colHanNop, int colCoQuanThu, int colMaCoQuanThu, int idxTableFrom, int idxTableTo)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		if (hTKKCMCSheetInfo == null)
		{
			return;
		}
		SheetView sheetView = winFpSpread.Sheets[0];
		HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[idxTableFrom];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[idxTableTo];
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		DateTime hanNopToKhai = GetHanNopToKhai(PeriodInfo);
		List<Item> list = new List<Item>();
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			Item item = new Item();
			item.Ma = sheetView.Cells[dataLocationTemplateRowStart + i - 1, colMaTM - 1].Text;
			item.SoTien = ConvertHelpers.ConvertToDouble(sheetView.Cells[dataLocationTemplateRowStart + i - 1, colSoTien - 1].Value);
			list.Add(item);
		}
		List<Item> list2 = new List<Item>();
		var enumerable = from x in list
			group x by x.Ma into g
			select new
			{
				Ma = g.Key,
				TongSoTien = g.Sum((Item x) => x.SoTien)
			};
		foreach (var item3 in enumerable)
		{
			Item item2 = new Item();
			item2.Ma = item3.Ma;
			item2.SoTien = item3.TongSoTien;
			list2.Add(item2);
		}
		if (hTKKCMCTableInfo2.RowsValues.Count > 1)
		{
			RemoveDataRows(sheetView, 1, dataLocationTemplateRowStart2 - 1, hTKKCMCTableInfo2.RowsValues.Count - 1, hTKKCMCSheetInfo);
		}
		foreach (HTKKCMCTableCellInfo cell in hTKKCMCSheetInfo.Tables[1].RowsInfo.Cells)
		{
			sheetView.Cells[dataLocationTemplateRowStart2 - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
		}
		for (int num = 0; num < list2.Count - 1; num++)
		{
			InsertTableRow(sheetView, dataLocationTemplateRowStart2, -1, dataLocationTemplateRowStart2, -1, -1, hTKKCMCSheetInfo, 1);
		}
		DateTime ngayDauKyKeKhai = GetNgayDauKyKeKhai(TTinChung, PeriodInfo);
		for (int num = 0; num < list2.Count; num++)
		{
			sheetView.Cells[dataLocationTemplateRowStart2 + num - 1, colStt - 1].Text = (num + 1).ToString();
			string[] mucByMaTm = GetMucByMaTm(list2[num].Ma, ngayDauKyKeKhai);
			if (mucByMaTm != null && mucByMaTm.Length == 2)
			{
				sheetView.Cells[dataLocationTemplateRowStart2 + num - 1, colMaMucTo - 1].Text = mucByMaTm[0];
				sheetView.Cells[dataLocationTemplateRowStart2 + num - 1, colMucTo - 1].Text = mucByMaTm[1];
			}
			sheetView.Cells[dataLocationTemplateRowStart2 + num - 1, colMaTMTo - 1].Text = list2[num].Ma;
			sheetView.Cells[dataLocationTemplateRowStart2 + num - 1, colSoTienTo - 1].Value = list2[num].SoTien;
			sheetView.Cells[dataLocationTemplateRowStart2 + num - 1, colTMTo - 1].Text = GetTenTmByMaTm(list2[num].Ma, ngayDauKyKeKhai);
			sheetView.Cells[dataLocationTemplateRowStart2 + num - 1, colHanNop - 1].Text = hanNopToKhai.Day + "/" + hanNopToKhai.Month + "/" + hanNopToKhai.Year;
		}
	}

	public static void RemoveValidateTable(SheetView sView, HTKKCMCTableInfo table, List<int> colsRemoveValidate = null)
	{
		int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < table.RowsValues.Count; i++)
		{
			foreach (HTKKCMCTableCellInfo cell in table.RowsInfo.Cells)
			{
				if (colsRemoveValidate != null && !colsRemoveValidate.Contains(cell.CellTemplate.ColIndex))
				{
					HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1);
				}
			}
		}
	}

	public static int GetAutoWidth(string[] items)
	{
		int num = 0;
		Graphics val = Graphics.FromHwnd(IntPtr.Zero);
		try
		{
			foreach (string text in items)
			{
				Size size = TextRenderer.MeasureText((IDeviceContext)(object)val, text, SystemFonts.DefaultFont);
				if (size.Width > num)
				{
					num = size.Width;
				}
			}
		}
		finally
		{
			((IDisposable)val)?.Dispose();
		}
		if (num > 1000)
		{
			num = 1000;
		}
		return num + 50;
	}

	public static void UpdateFileMapMctKhbs(string Id, string findAttribute, string findAttributeValue, string updateAttribute, string updateAttributeValue)
	{
		string text = "MapMCT.xml";
		XDocument val = XDocument.Load(text);
		XElement val2 = ((XContainer)val).Descendants(XName.op_Implicit("Map")).FirstOrDefault((Func<XElement, bool>)((XElement x) => (string)x.Attribute(XName.op_Implicit("ID")) == Id));
		if (val2 != null)
		{
			XElement val3 = ((XContainer)val2).Descendants(XName.op_Implicit("Item")).FirstOrDefault((Func<XElement, bool>)((XElement x) => (string)x.Attribute(XName.op_Implicit(findAttribute)) == findAttributeValue));
			if (val3 != null)
			{
				val3.SetAttributeValue(XName.op_Implicit(updateAttribute), (object)updateAttributeValue);
				val.Save(text);
			}
		}
	}

	public static void lockUnlockColsInTableByFile(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol, JArray mapAnHiens)
	{
		try
		{
			if (iRow >= 0)
			{
				int num = DetectIndexOfTable(sInfo, iRow);
				if (num < 0)
				{
					return;
				}
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[num];
				{
					foreach (JObject item in (IEnumerable<JToken>)mapAnHiens)
					{
						int num2 = item["sheetIndex"].Value<int>();
						if (sInfo.SheetIndex != num2)
						{
							continue;
						}
						JArray jArray = (JArray)item["tables"];
						foreach (JObject item2 in (IEnumerable<JToken>)jArray)
						{
							string text = item2["tablePath"].Value<string>();
							if (hTKKCMCTableInfo.TablePath != text)
							{
								continue;
							}
							string column = item2["colCoThayDoi"].Value<string>();
							if (iCol != HTKKFile.ColumnLettersToIndex(column))
							{
								continue;
							}
							string column2 = item2["colMa"].Value<string>();
							int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
							List<int> list = new List<int>();
							foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo.RowsInfo.Cells)
							{
								list.Add(cell.CellTemplate.ColIndex);
							}
							LockUnLockCols(sView, iRow, list, locked: false, deleteData: false);
							for (int i = 1; i < 10; i++)
							{
								JObject jObject3 = (JObject)item2["cap" + i];
								if (jObject3 == null)
								{
									break;
								}
								string text2 = jObject3["valueMa"].Value<string>();
								string text3 = jObject3["colLocks"].Value<string>();
								string text4 = sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex(column2) - 1].Text;
								if (!string.IsNullOrEmpty(text3) && text4 == text2)
								{
									string[] array = text3.Split(new char[1] { ',' });
									foreach (string column3 in array)
									{
										LockUnLockCol(sView, iRow, HTKKFile.ColumnLettersToIndex(column3), locked: true, deleteData: true);
									}
								}
							}
						}
					}
					return;
				}
			}
			foreach (JObject item3 in (IEnumerable<JToken>)mapAnHiens)
			{
				int num2 = item3["sheetIndex"].Value<int>();
				if (sInfo.SheetIndex != num2)
				{
					break;
				}
				JArray jArray = (JArray)item3["tables"];
				foreach (JObject item4 in (IEnumerable<JToken>)jArray)
				{
					string tablePath = item4["tablePath"].Value<string>();
					HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath);
					string column2 = item4["colMa"].Value<string>();
					int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
					List<int> list = new List<int>();
					foreach (HTKKCMCTableCellInfo cell2 in hTKKCMCTableInfo.RowsInfo.Cells)
					{
						list.Add(cell2.CellTemplate.ColIndex);
					}
					for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsValues.Count; num3++)
					{
						LockUnLockCols(sView, dataLocationTemplateRowStart + num3, list, locked: false, deleteData: false);
						for (int i = 1; i < 10; i++)
						{
							JObject jObject3 = (JObject)item4["cap" + i];
							if (jObject3 == null)
							{
								break;
							}
							string text2 = jObject3["valueMa"].Value<string>();
							string text3 = jObject3["colLocks"].Value<string>();
							string text4 = sView.Cells[dataLocationTemplateRowStart + num3 - 1, HTKKFile.ColumnLettersToIndex(column2) - 1].Text;
							if (!string.IsNullOrEmpty(text3) && text4 == text2)
							{
								string[] array = text3.Split(new char[1] { ',' });
								foreach (string column3 in array)
								{
									LockUnLockCol(sView, dataLocationTemplateRowStart + num3, HTKKFile.ColumnLettersToIndex(column3), locked: true, deleteData: true);
								}
							}
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string sumInTableByFile(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol, JArray mapSums)
	{
		string text = "";
		try
		{
			List<CapSumInTable> list = new List<CapSumInTable>();
			if (iRow >= 0)
			{
				int num = DetectIndexOfTable(sInfo, iRow);
				if (num < 0)
				{
					return "";
				}
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[num];
				foreach (JObject item in (IEnumerable<JToken>)mapSums)
				{
					int num2 = item["sheetIndex"].Value<int>();
					if (sInfo.SheetIndex != num2)
					{
						continue;
					}
					JArray jArray = (JArray)item["tables"];
					foreach (JObject item2 in (IEnumerable<JToken>)jArray)
					{
						string text2 = item2["tablePath"].Value<string>();
						if (hTKKCMCTableInfo.TablePath != text2)
						{
							continue;
						}
						string column = item2["colCoThayDoi"].Value<string>();
						if (iCol != HTKKFile.ColumnLettersToIndex(column))
						{
							continue;
						}
						string column2 = item2["colMa"].Value<string>();
						int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
						HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart - 1, HTKKFile.ColumnLettersToIndex(column2) - 1);
						string text3 = sView.Cells[dataLocationTemplateRowStart - 1, HTKKFile.ColumnLettersToIndex(column2) - 1].Text;
						JObject jObject3 = (JObject)item2["cap1"];
						if (jObject3 == null)
						{
							return "";
						}
						if (jObject3["colsSum"] == null)
						{
							return "";
						}
						string text4 = jObject3["valueMa"].Value<string>();
						if (text3 != text4 || (string.IsNullOrEmpty(text3) && hTKKCMCTableInfo.RowsValues.Count > 1))
						{
							text = "Bắt buộc cấp đầu tiên là cấp 1";
							HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart - 1, HTKKFile.ColumnLettersToIndex(column2) - 1, text);
							return text;
						}
						CapSumInTable capSumInTable = new CapSumInTable();
						capSumInTable.Cap = "cap1";
						capSumInTable.RowSum = dataLocationTemplateRowStart;
						capSumInTable.RowsChilden = new List<int>();
						capSumInTable.Value = text4;
						capSumInTable.ColsSum = jObject3["colsSum"].Value<string>();
						list.Add(capSumInTable);
						for (int i = 1; i < hTKKCMCTableInfo.RowsValues.Count; i++)
						{
							string valueMaSview = sView.Cells[dataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex(column2) - 1].Text;
							if (string.IsNullOrEmpty(valueMaSview))
							{
								continue;
							}
							JProperty jProperty = item2.Properties().FirstOrDefault((JProperty p) => p.Name.StartsWith("cap") && (string)p.Value["valueMa"] == valueMaSview);
							if (jProperty == null)
							{
								continue;
							}
							JObject jObject4 = (JObject)jProperty.Value;
							string capHienTaiSo = GetNumberInString(jProperty.Name);
							if (string.IsNullOrEmpty(capHienTaiSo) || Convert.ToInt32(capHienTaiSo) < 1)
							{
								continue;
							}
							List<CapSumInTable> list2 = list.FindAll((CapSumInTable x) => x.Cap == "cap" + (Convert.ToInt32(capHienTaiSo) - 1));
							string numberInString = GetNumberInString(list[list.Count - 1].Cap);
							if (list2.Count == 0)
							{
								CapSumInTable capSumInTable2 = new CapSumInTable();
								capSumInTable2.Cap = jProperty.Name;
								capSumInTable2.RowSum = dataLocationTemplateRowStart + i;
								capSumInTable2.RowsChilden = new List<int>();
								capSumInTable2.Value = valueMaSview;
								capSumInTable2.ColsSum = jObject4["colsSum"].Value<string>();
								list.Add(capSumInTable2);
								continue;
							}
							CapSumInTable capSumInTable3 = list2[list2.Count - 1];
							string numberInString2 = GetNumberInString(capSumInTable3.Cap);
							if (Convert.ToInt32(capHienTaiSo) - 1 != Convert.ToInt32(numberInString2))
							{
								CapSumInTable capSumInTable2 = new CapSumInTable();
								capSumInTable2.Cap = jProperty.Name;
								capSumInTable2.RowSum = dataLocationTemplateRowStart + i;
								capSumInTable2.RowsChilden = new List<int>();
								capSumInTable2.Value = valueMaSview;
								capSumInTable2.ColsSum = jObject4["colsSum"].Value<string>();
								list.Add(capSumInTable2);
							}
							else
							{
								capSumInTable3.RowsChilden.Add(dataLocationTemplateRowStart + i);
								CapSumInTable capSumInTable2 = new CapSumInTable();
								capSumInTable2.Cap = jProperty.Name;
								capSumInTable2.RowSum = dataLocationTemplateRowStart + i;
								capSumInTable2.RowsChilden = new List<int>();
								capSumInTable2.Value = valueMaSview;
								capSumInTable2.ColsSum = jObject4["colsSum"].Value<string>();
								list.Add(capSumInTable2);
							}
						}
						foreach (CapSumInTable item3 in list)
						{
							if (string.IsNullOrEmpty(item3.ColsSum))
							{
								continue;
							}
							string[] array = item3.ColsSum.Split(new char[1] { ',' });
							foreach (string text5 in array)
							{
								string text6 = "";
								foreach (int item4 in item3.RowsChilden)
								{
									object obj = text6;
									text6 = string.Concat(obj, text5, item4, "+");
								}
								if (text6.Length > 0)
								{
									text6 = text6.Substring(0, text6.Length - 1);
								}
								sView.Cells[item3.RowSum - 1, HTKKFile.ColumnLettersToIndex(text5) - 1].Formula = text6;
							}
						}
					}
				}
			}
			return "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return ex.Message;
		}
	}

	public static void TongHopTuBangSangBangTheoLoai(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int sheetIndexFrom, int sheetIndexTo, string tablePathFrom, string tablePathTo, List<int> colsFrom, List<int> colsTo, List<string> types, int colLoai)
	{
		SheetView sheetView = winFpSpread.Sheets[sheetIndexFrom];
		SheetView sheetView2 = winFpSpread.Sheets[sheetIndexTo];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndexFrom];
		HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[sheetIndexTo];
		HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePathFrom);
		int num = hTKKCMCSheetInfo2.Tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == tablePathTo);
		HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo2.Tables[num];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		List<HTKKCMCRowValue> list = new List<HTKKCMCRowValue>();
		for (int num2 = 0; num2 < hTKKCMCTableInfo.RowsValues.Count; num2++)
		{
			string text = sheetView.Cells[dataLocationTemplateRowStart + num2 - 1, colLoai - 1].Text;
			if (!types.Contains(text))
			{
				continue;
			}
			HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
			foreach (int item in colsFrom)
			{
				object value = sheetView.Cells[dataLocationTemplateRowStart + num2 - 1, item - 1].Value;
				hTKKCMCRowValue.RowValue.Add((value == null) ? "" : value.ToString());
			}
			list.Add(hTKKCMCRowValue);
		}
		RemoveDataRows(sheetView2, num, dataLocationTemplateRowStart2, hTKKCMCTableInfo2.RowsValues.Count - 1, hTKKCMCSheetInfo2);
		setDefaultAllTable(sheetView2, hTKKCMCTableInfo2);
		int num3 = 0;
		foreach (HTKKCMCRowValue item2 in list)
		{
			if (num3 > 0)
			{
				InsertTableRow(sheetView2, dataLocationTemplateRowStart2 + num3 - 1, -1, dataLocationTemplateRowStart2 + num3 - 1, -1, -1, hTKKCMCSheetInfo2, num);
			}
			num3++;
		}
		for (int num2 = 0; num2 < list.Count; num2++)
		{
			int num4 = 0;
			foreach (int item3 in colsTo)
			{
				sheetView2.Cells[dataLocationTemplateRowStart2 + num2 - 1, item3 - 1].Value = list[num2].RowValue[num4];
				num4++;
			}
		}
	}

	public static void setDefaultAllTable(SheetView sView, HTKKCMCTableInfo table)
	{
		int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < table.RowsValues.Count; i++)
		{
			foreach (HTKKCMCTableCellInfo cell in table.RowsInfo.Cells)
			{
				sView.Cells[dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
			}
		}
	}

	public static void ExportXmlKHBSSuaDonViTien(string filePath, FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int sheetIndex, string pathTenDonViTien, string pathMaDonViTien)
	{
		if (File.Exists(filePath))
		{
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTenDonViTien);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMaDonViTien);
			string text = winFpSpread.Sheets[sheetIndex].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
			string text2 = winFpSpread.Sheets[sheetIndex].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text;
			XDocument val = XDocument.Load(filePath);
			XNamespace val2 = XNamespace.op_Implicit("http://kekhaithue.gdt.gov.vn/TKhaiThue");
			XElement val3 = ((XContainer)((XContainer)((XContainer)((XContainer)val.Root).Element(val2 + "HSoKhaiThue")).Element(val2 + "PLuc")).Element(val2 + "PL01_KHBS")).Element(val2 + "ma_DonViTien");
			XElement val4 = ((XContainer)((XContainer)((XContainer)((XContainer)val.Root).Element(val2 + "HSoKhaiThue")).Element(val2 + "PLuc")).Element(val2 + "PL01_KHBS")).Element(val2 + "ten_DonViTien");
			XElement val5 = ((XContainer)((XContainer)((XContainer)((XContainer)val.Root).Element(val2 + "HSoKhaiThue")).Element(val2 + "PLuc")).Element(val2 + "PL01_1_KHBS")).Element(val2 + "ma_DonViTien");
			XElement val6 = ((XContainer)((XContainer)((XContainer)((XContainer)val.Root).Element(val2 + "HSoKhaiThue")).Element(val2 + "PLuc")).Element(val2 + "PL01_1_KHBS")).Element(val2 + "ten_DonViTien");
			if (val3 != null)
			{
				val3.Value = text2;
			}
			if (val4 != null)
			{
				val4.Value = text;
			}
			if (val5 != null)
			{
				val5.Value = text2;
			}
			if (val6 != null)
			{
				val6.Value = text;
			}
			val.Save(filePath);
		}
	}

	public static string Chen1KyTuGiuaCacTu(string input)
	{
		return string.Join(".", input.Select((char c) => c.ToString()).ToArray());
	}

	public static string GetValueInDm(string path, string attr, string valueAttr, string retureAttr)
	{
		XDocument val = XDocument.Load(path);
		XElement val2 = ((XContainer)val).Descendants(XName.op_Implicit("Item")).FirstOrDefault((Func<XElement, bool>)((XElement x) => x.Attribute(XName.op_Implicit(attr)) != null && x.Attribute(XName.op_Implicit(attr)).Value.Contains(valueAttr)));
		if (val2 != null)
		{
			return val2.Attribute(XName.op_Implicit(retureAttr)).Value;
		}
		return "";
	}

	public static string GetSttVietnameseLetter(int number)
	{
		string text = "abcdefghijklmnopqrstuvwxyz";
		if (number <= 0 || number > text.Length)
		{
			return null;
		}
		return text[number - 1].ToString();
	}

	public static void CheckMaDiaDiemKd(SheetView sView, HTKKCMCSheetInfo sInfo, string pathMaDiaDiemKD, ref int iwarningColor)
	{
		string msg = "Mã điểm điểm kinh doanh bắt buộc đủ 5 số";
		List<HTKKCMCCellInfo> list = sInfo.StaticCells.FindAll((HTKKCMCCellInfo x) => x.Path.ToLower().IndexOf(pathMaDiaDiemKD) >= 0);
		foreach (HTKKCMCCellInfo item in list)
		{
			string text = sView.Cells[item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1].Text;
			if (!string.IsNullOrEmpty(text))
			{
				if (text.Length != 5)
				{
					HTKKFile.SetNodeBackGroundPink(sView, item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1, msg);
					iwarningColor = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1);
				}
			}
		}
		foreach (HTKKCMCTableInfo table in sInfo.Tables)
		{
			List<HTKKCMCTableCellInfo> list2 = table.RowsInfo.Cells.FindAll((HTKKCMCTableCellInfo x) => x.Path.ToLower().IndexOf(pathMaDiaDiemKD) >= 0);
			foreach (HTKKCMCTableCellInfo item2 in list2)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int num = 0; num < table.RowsValues.Count; num++)
				{
					string text = sView.Cells[dataLocationTemplateRowStart + num - 1, item2.CellTemplate.ColIndex - 1].Text;
					if (!string.IsNullOrEmpty(text))
					{
						if (text.Length != 5)
						{
							HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + num - 1, item2.CellTemplate.ColIndex - 1, msg);
							iwarningColor = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + num - 1, item2.CellTemplate.ColIndex - 1);
						}
					}
				}
			}
		}
		foreach (HTKKCMCTableInfo table2 in sInfo.Tables)
		{
			List<HTKKCMCCellInfo> list3 = table2.Header.StaticCells.FindAll((HTKKCMCCellInfo x) => x.Path.ToLower().IndexOf(pathMaDiaDiemKD) >= 0);
			foreach (HTKKCMCCellInfo item3 in list3)
			{
				string text = sView.Cells[item3.CellTemplate.RowIndex - 1, item3.CellTemplate.ColIndex - 1].Text;
				if (!string.IsNullOrEmpty(text))
				{
					if (text.Length != 5)
					{
						HTKKFile.SetNodeBackGroundPink(sView, item3.CellTemplate.RowIndex - 1, item3.CellTemplate.ColIndex - 1, msg);
						iwarningColor = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sView, item3.CellTemplate.RowIndex - 1, item3.CellTemplate.ColIndex - 1);
					}
				}
			}
		}
	}

	public static void DeleteNodeXmlByPath(XmlNode xmlNode, string pathDelete)
	{
		try
		{
			if (xmlNode == null)
			{
				return;
			}
			string[] array = pathDelete.Split(new char[1] { '/' });
			XmlNode val = xmlNode;
			int num = 0;
			if (array[0] == xmlNode.Name)
			{
				num = 1;
			}
			for (int i = num; i < array.Length; i++)
			{
				val = val.SelectSingleNode(array[i]);
				if (val == null)
				{
					HTKKFile.Log("Không tìm thấy node: " + array[i]);
					return;
				}
			}
			if (val.ParentNode != null)
			{
				val.ParentNode.RemoveChild(val);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool ValidateTableEmpty(SheetView sView, HTKKCMCTableInfo table, List<int> colsCheck)
	{
		bool result = false;
		try
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				foreach (HTKKCMCTableCellInfo cell in table.RowsInfo.Cells)
				{
					if (!colsCheck.Contains(cell.CellTemplate.ColIndex))
					{
						continue;
					}
					if (cell.Controltype == "16")
					{
						double num = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1].Value);
						if (num != Convert.ToDouble(cell.DefaultValue))
						{
							result = true;
							break;
						}
					}
					else if (cell.Controltype == "2")
					{
						string text = sView.Cells[dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1].Text;
						if (text.ToLower() != cell.DefaultValue.ToLower())
						{
							result = true;
							break;
						}
					}
					else
					{
						string text = sView.Cells[dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1].Text;
						if (text != cell.DefaultValue)
						{
							result = true;
							break;
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static string CheckToKhaiHetHieuLuc(string menuId, HTKKPeriodInfo periodInfo)
	{
		//IL_0041: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_007a: Expected O, but got Unknown
		if (string.IsNullOrEmpty(menuId) || !"78~A1~A2~R4~C9~D6~".Split(new char[1] { '~' }).Contains(menuId))
		{
			return "";
		}
		try
		{
			XmlDocument val = new XmlDocument();
			val.Load("Menu.xml");
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("Root/Menu");
			foreach (XmlNode item in val2)
			{
				XmlNode val3 = item;
				XmlAttribute val4 = val3.Attributes["ID"];
				if (val4 == null || ((XmlNode)val4).Value != menuId)
				{
					continue;
				}
				XmlNode val5 = val3.SelectSingleNode("Validity");
				if (val5 == null || val5.Attributes["EndDate"] == null)
				{
					return "";
				}
				string value = ((XmlNode)val5.Attributes["EndDate"]).Value;
				if (!string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(periodInfo.DenNgay))
				{
					DateTime dateTime = DateTime.ParseExact(value, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					DateTime dateTime2 = DateTime.ParseExact(periodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					if (dateTime2 > dateTime)
					{
						return value;
					}
				}
				return "";
			}
			return "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static string CheckToKhaiBatDauCoHieuLuc(string menuId, HTKKPeriodInfo periodInfo)
	{
		//IL_0041: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_007a: Expected O, but got Unknown
		if (string.IsNullOrEmpty(menuId) || !"U9~X1~X2~".Split(new char[1] { '~' }).Contains(menuId))
		{
			return "";
		}
		try
		{
			XmlDocument val = new XmlDocument();
			val.Load("Menu.xml");
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("Root/Menu");
			foreach (XmlNode item in val2)
			{
				XmlNode val3 = item;
				XmlAttribute val4 = val3.Attributes["ID"];
				if (val4 == null || ((XmlNode)val4).Value != menuId)
				{
					continue;
				}
				XmlNode val5 = val3.SelectSingleNode("Validity");
				if (val5 == null || val5.Attributes["StartDate"] == null)
				{
					return "";
				}
				string value = ((XmlNode)val5.Attributes["StartDate"]).Value;
				if (!string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(periodInfo.TuNgay))
				{
					DateTime dateTime = DateTime.ParseExact(value, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					DateTime dateTime2 = DateTime.ParseExact(periodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					if (dateTime2 < dateTime)
					{
						return value;
					}
				}
				return "";
			}
			return "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static DateTime ngayBatDauTheoTaiChinh(HTKKHeaderInfo headerInfo, string nam)
	{
		string s = headerInfo.NgayBatDauNamTaiChinh + "/" + nam;
		return DateTime.ParseExact(s, "dd/MM/yyyy", CultureInfo.CurrentCulture);
	}

	public static DateTime ngayKetThuc6ThangDauTheoTaiChinh(DateTime ngayBatDauTheoTaiChinh)
	{
		return ngayBatDauTheoTaiChinh.AddMonths(6).AddDays(-1.0);
	}

	public static DateTime ngayBatDau6ThangCuoiNam_TheoTaiChinh(DateTime ngayCuoiNamTaiChinh)
	{
		return ngayCuoiNamTaiChinh.AddMonths(-6).AddDays(1.0);
	}

	public static string removeKyTuDacBietFileName(string fileName)
	{
		fileName = RemoveSpecialCharactersUnicode(fileName);
		return fileName.Replace("/", "---").Replace("\\", "###");
	}

	public static void ShowSheetTablePage(FpSpread winFpSpread, int sheetIndex, HTKKCMCSheetInfo sInfo, int page)
	{
		if (winFpSpread == null)
		{
			throw new ArgumentNullException("winFpSpread");
		}
		if (sInfo == null)
		{
			throw new ArgumentNullException("sInfo");
		}
		if (sheetIndex < 0 || sheetIndex >= winFpSpread.Sheets.Count)
		{
			throw new ArgumentOutOfRangeException("sheetIndex");
		}
		SheetView sheetView = winFpSpread.Sheets[sheetIndex];
		winFpSpread.SuspendLayout();
		try
		{
			RemoveSheetTablePageOldWarnings(sheetView);
			ResetSheetTablePage(sheetView, winFpSpread, sheetIndex);
			int sheetTablePageCount = GetSheetTablePageCount(sInfo);
			if (sheetTablePageCount == 0)
			{
				_page = 0;
				_pageSize = Math.Max(1, 50);
				_totalPage = 0;
				return;
			}
			int num = (_page = Math.Max(0, Math.Min(page, sheetTablePageCount - 1)));
			_pageSize = Math.Max(1, 50);
			_totalPage = sheetTablePageCount;
			SheetTablePage sheetTablePage = BuildSheetTablePage(sInfo, num);
			SheetTablePageRenderState sheetTablePageRenderState = new SheetTablePageRenderState();
			SetSheetTablePageRenderState(winFpSpread, sheetIndex, sheetTablePageRenderState);
			foreach (SheetTablePageSegment segment in sheetTablePage.Segments)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[segment.TableIndex];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				int num2 = (segment.TargetDataRow = MapTemplateExpandedRow(dataLocationTemplateRowStart, sInfo, sheetTablePage));
				int num3 = Math.Max(0, segment.DataCount - 1);
				HTKKFile.Log(string.Format(CultureInfo.InvariantCulture, "ShowSheetTablePage page={0}, table={1}, sourceStart={2}, sourceCount={3}, templateRow={4}, targetRow={5}, insertCount={6}, sheetRowCount={7}", num, segment.TableIndex, segment.FirstDataIndex, segment.DataCount, dataLocationTemplateRowStart, num2, num3, sheetView.RowCount));
				if (num3 > 0)
				{
					int num4 = num2;
					sheetView.Rows.Add(num4, num3);
					for (int i = 0; i < num3; i++)
					{
						sheetView.CopyRange(num2 - 1, 0, num4 + i, 0, 1, sheetView.ColumnCount, dataOnly: false);
					}
					sheetTablePageRenderState.InsertedRows.Add(new KeyValuePair<int, int>(num4, num3));
				}
			}
			RemoveSheetTablePageOldWarnings(sheetView);
			BindVisibleStaticCells(sheetView, sInfo, sheetTablePage, sheetTablePageRenderState);
			SetTablePageVisibility(sheetView, sInfo, sheetTablePage, sheetTablePageRenderState);
			foreach (SheetTablePageSegment segment2 in sheetTablePage.Segments)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[segment2.TableIndex];
				BindTableStaticCells(sheetView, hTKKCMCTableInfo, sInfo, sheetTablePage);
				BindSheetTablePage(sheetView, hTKKCMCTableInfo, segment2, segment2.TargetDataRow);
			}
			HideSheetTableTemplateZeroRows(sheetView, sInfo, sheetTablePage, sheetTablePageRenderState);
			RemoveSheetTablePageOldWarnings(sheetView, sheetTablePage);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw;
		}
		finally
		{
			winFpSpread.ResumeLayout(performLayout: true);
			((Control)winFpSpread).Invalidate(true);
			((Control)winFpSpread).Update();
		}
	}

	private static void RemoveSheetTablePageOldWarnings(SheetView sView)
	{
		int num = Math.Min(sView.RowCount, sView.NonEmptyRowCount);
		int num2 = Math.Min(sView.ColumnCount, sView.NonEmptyColumnCount);
		for (int i = 0; i < num; i++)
		{
			for (int j = 0; j < num2; j++)
			{
				Cell cell = sView.Cells[i, j];
				if (!string.IsNullOrEmpty(cell.Note))
				{
					cell.ResetNote();
					cell.BackColor = Color.White;
				}
			}
		}
	}

	private static void RemoveSheetTablePageOldWarnings(SheetView sView, SheetTablePage selectedPage)
	{
		foreach (SheetTablePageSegment segment in selectedPage.Segments)
		{
			for (int i = 0; i < segment.DataCount; i++)
			{
				int num = segment.TargetDataRow + i - 1;
				if (num < 0 || num >= sView.RowCount)
				{
					continue;
				}
				for (int j = 0; j < sView.ColumnCount; j++)
				{
					Cell cell = sView.Cells[num, j];
					if (!string.IsNullOrEmpty(cell.Note))
					{
						cell.ResetNote();
						cell.BackColor = Color.White;
					}
				}
			}
		}
	}

	public static int GetSheetTablePageCount(HTKKCMCSheetInfo sInfo)
	{
		if (sInfo == null)
		{
			return 0;
		}
		long sheetTableDataRowCount = GetSheetTableDataRowCount(sInfo);
		if (sheetTableDataRowCount == 0)
		{
			return 0;
		}
		long num = Math.Max(1, 50);
		long num2 = (sheetTableDataRowCount + num - 1) / num;
		return (int)((num2 > int.MaxValue) ? int.MaxValue : num2);
	}

	private static long GetSheetTableDataRowCount(HTKKCMCSheetInfo sInfo)
	{
		long num = 0L;
		if (sInfo == null || sInfo.Tables == null)
		{
			return num;
		}
		foreach (HTKKCMCTableInfo table in sInfo.Tables)
		{
			if (table != null && table.RowsValues != null)
			{
				num += table.RowsValues.Count;
			}
		}
		return num;
	}

	private static SheetTablePage BuildSheetTablePage(HTKKCMCSheetInfo sInfo, int page)
	{
		SheetTablePage sheetTablePage = new SheetTablePage();
		long num = Math.Max(1, 50);
		long num2 = page * num;
		long num3 = num2 + num;
		long num4 = 0L;
		for (int i = 0; i < sInfo.Tables.Count; i++)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
			int num5 = ((hTKKCMCTableInfo != null && hTKKCMCTableInfo.RowsValues != null) ? hTKKCMCTableInfo.RowsValues.Count : 0);
			long num6 = num4 + num5;
			long num7 = Math.Max(num2, num4);
			long num8 = Math.Min(num3, num6);
			if (num7 < num8)
			{
				int num9 = (int)(num8 - num7);
				sheetTablePage.Segments.Add(new SheetTablePageSegment
				{
					TableIndex = i,
					FirstDataIndex = (int)(num7 - num4),
					DataCount = num9
				});
				sheetTablePage.DataCount += num9;
			}
			num4 = num6;
			if (num4 >= num3)
			{
				break;
			}
		}
		return sheetTablePage;
	}

	private static void BindSheetTablePage(SheetView sView, HTKKCMCTableInfo table, SheetTablePageSegment segment, int targetRow)
	{
		for (int i = 0; i < segment.DataCount; i++)
		{
			HTKKCMCRowValue sheetTablePageSourceRow = GetSheetTablePageSourceRow(table, segment.FirstDataIndex + i);
			if (sheetTablePageSourceRow == null || sheetTablePageSourceRow.RowValue == null)
			{
				continue;
			}
			int num = Math.Min(table.RowsInfo.Cells.Count, sheetTablePageSourceRow.RowValue.Count);
			for (int j = 0; j < num; j++)
			{
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells[j];
				int num2 = ((hTKKCMCTableCellInfo.CellTemplate == null) ? (-1) : hTKKCMCTableCellInfo.CellTemplate.ColIndex);
				if (num2 > 0)
				{
					fpSpBindDataCell(sView, targetRow + i, num2, hTKKCMCTableCellInfo.Controltype, sheetTablePageSourceRow.RowValue[j]);
				}
			}
			if (sheetTablePageSourceRow.RowValue.Count > 0)
			{
				int sheetTableSttReportColumn = GetSheetTableSttReportColumn(table);
				if (sheetTableSttReportColumn > 0)
				{
					Cell cell = sView.Cells[targetRow + i - 1, sheetTableSttReportColumn - 1];
					cell.ResetFormula();
					cell.Text = ((sheetTablePageSourceRow.RowValue[0] == null) ? "" : sheetTablePageSourceRow.RowValue[0]);
				}
			}
		}
	}

	private static int GetSheetTableSttReportColumn(HTKKCMCTableInfo table)
	{
		if (table.iColumnSTT > 0)
		{
			return table.iColumnSTT;
		}
		if (table.RowsInfo != null && table.RowsInfo.Cells != null && table.RowsInfo.Cells.Count > 0)
		{
			HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells[0];
			if (hTKKCMCTableCellInfo.CellTemplate != null && hTKKCMCTableCellInfo.CellTemplate.ColIndex > 0)
			{
				return hTKKCMCTableCellInfo.CellTemplate.ColIndex;
			}
		}
		return -1;
	}

	private static void HideSheetTableTemplateZeroRows(SheetView sView, HTKKCMCSheetInfo sInfo, SheetTablePage selectedPage, SheetTablePageRenderState state)
	{
		if (sView == null || sInfo == null || sInfo.Tables == null || selectedPage == null)
		{
			return;
		}
		foreach (SheetTablePageSegment segment in selectedPage.Segments)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[segment.TableIndex];
			if (hTKKCMCTableInfo == null)
			{
				continue;
			}
			int sheetTableSttReportColumn = GetSheetTableSttReportColumn(hTKKCMCTableInfo);
			if (sheetTableSttReportColumn <= 0 || sheetTableSttReportColumn > sView.ColumnCount)
			{
				continue;
			}
			int num = Math.Max(0, MapTemplateExpandedRow(GetTableFirstTemplateRow(hTKKCMCTableInfo), sInfo, selectedPage) - 1);
			int num2 = Math.Min(sView.RowCount - 1, segment.TargetDataRow + segment.DataCount - 1);
			for (int i = num; i <= num2; i++)
			{
				string text = sView.Cells[i, sheetTableSttReportColumn - 1].Text;
				if (string.Equals((text == null) ? "" : text.Trim(), "0", StringComparison.Ordinal))
				{
					SetPageRowVisible(sView, i, visible: false, state);
				}
			}
		}
	}

	private static HTKKCMCRowValue GetSheetTablePageSourceRow(HTKKCMCTableInfo table, int sourceIndex)
	{
		if (table == null || table.RowsValues == null || sourceIndex < 0 || sourceIndex >= table.RowsValues.Count)
		{
			return null;
		}
		return table.RowsValues[sourceIndex];
	}

	private static void BindTableStaticCells(SheetView sView, HTKKCMCTableInfo table, HTKKCMCSheetInfo sInfo, SheetTablePage selectedPage)
	{
		BindTableStaticCellList(sView, (table.Header == null) ? null : table.Header.StaticCells, sInfo, selectedPage);
		BindTableStaticCellList(sView, (table.Footer == null) ? null : table.Footer.StaticCells, sInfo, selectedPage);
	}

	private static void BindTableStaticCellList(SheetView sView, List<HTKKCMCCellInfo> cells, HTKKCMCSheetInfo sInfo, SheetTablePage selectedPage)
	{
		if (cells == null)
		{
			return;
		}
		foreach (HTKKCMCCellInfo cell in cells)
		{
			if (cell != null && cell.CellTemplate != null && cell.CellTemplate.RowIndex > 0 && cell.CellTemplate.ColIndex > 0)
			{
				fpSpBindDataCell(sView, MapTemplateExpandedRow(cell.CellTemplate.RowIndex, sInfo, selectedPage), cell.CellTemplate.ColIndex, cell.Controltype, cell.Value);
			}
		}
	}

	private static void SetTablePageVisibility(SheetView sView, HTKKCMCSheetInfo sInfo, SheetTablePage selectedPage, SheetTablePageRenderState state)
	{
		int tableIndex;
		for (tableIndex = 0; tableIndex < sInfo.Tables.Count; tableIndex++)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[tableIndex];
			SheetTablePageSegment sheetTablePageSegment = selectedPage.Segments.Find((SheetTablePageSegment x) => x.TableIndex == tableIndex);
			int num = MapTemplateExpandedRow(GetTableFirstTemplateRow(hTKKCMCTableInfo), sInfo, selectedPage);
			int num2 = MapTemplateExpandedRow(GetTableLastTemplateRow(hTKKCMCTableInfo), sInfo, selectedPage);
			for (int num3 = num; num3 <= num2; num3++)
			{
				SetPageRowVisible(sView, num3 - 1, visible: false, state);
			}
			if (sheetTablePageSegment != null)
			{
				bool flag = sheetTablePageSegment.FirstDataIndex + sheetTablePageSegment.DataCount >= hTKKCMCTableInfo.RowsValues.Count;
				bool flag2 = sheetTablePageSegment == selectedPage.Segments[selectedPage.Segments.Count - 1];
				bool flag3 = selectedPage.DataCount >= 50;
				int num4 = ((flag && (!flag2 || !flag3)) ? num2 : (sheetTablePageSegment.TargetDataRow + sheetTablePageSegment.DataCount - 1));
				for (int num3 = num; num3 <= num4; num3++)
				{
					SetPageRowVisible(sView, num3 - 1, visible: true, state);
				}
				for (int num5 = 0; num5 < sheetTablePageSegment.DataCount; num5++)
				{
					SetPageRowVisible(sView, sheetTablePageSegment.TargetDataRow + num5 - 1, visible: true, state);
				}
			}
		}
	}

	private static int GetTableFirstTemplateRow(HTKKCMCTableInfo table)
	{
		if (table == null || table.Location == null)
		{
			return 1;
		}
		int num = ((table.Location.TableTemplateRowStart > 0) ? table.Location.TableTemplateRowStart : table.Location.DataLocationTemplateRowStart);
		if (num <= 0)
		{
			num = 1;
		}
		if (table.Header != null && table.Header.StaticCells != null)
		{
			foreach (HTKKCMCCellInfo staticCell in table.Header.StaticCells)
			{
				if (staticCell != null && staticCell.CellTemplate != null && staticCell.CellTemplate.RowIndex > 0)
				{
					num = Math.Min(num, staticCell.CellTemplate.RowIndex);
				}
			}
		}
		return num;
	}

	private static int GetTableLastTemplateRow(HTKKCMCTableInfo table)
	{
		if (table == null || table.Location == null)
		{
			return 1;
		}
		int tableFirstTemplateRow = GetTableFirstTemplateRow(table);
		int num = Math.Max(table.Location.DataLocationTemplateRowStart, tableFirstTemplateRow + Math.Max(0, table.TableLength));
		if (table.Header != null && table.Header.StaticCells != null)
		{
			foreach (HTKKCMCCellInfo staticCell in table.Header.StaticCells)
			{
				if (staticCell != null && staticCell.CellTemplate != null && staticCell.CellTemplate.RowIndex > 0)
				{
					num = Math.Max(num, staticCell.CellTemplate.RowIndex);
				}
			}
		}
		if (table.Footer != null && table.Footer.StaticCells != null)
		{
			foreach (HTKKCMCCellInfo staticCell2 in table.Footer.StaticCells)
			{
				if (staticCell2 != null && staticCell2.CellTemplate != null && staticCell2.CellTemplate.RowIndex > 0)
				{
					num = Math.Max(num, staticCell2.CellTemplate.RowIndex);
				}
			}
		}
		return num;
	}

	private static void BindVisibleStaticCells(SheetView sView, HTKKCMCSheetInfo sInfo, SheetTablePage selectedPage, SheetTablePageRenderState state)
	{
		SheetTablePageSegment sheetTablePageSegment = selectedPage.Segments[selectedPage.Segments.Count - 1];
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[sheetTablePageSegment.TableIndex];
		int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
		if (selectedPage.DataCount < 50)
		{
			num = int.MaxValue;
		}
		if (sInfo.StaticCells == null)
		{
			return;
		}
		foreach (HTKKCMCCellInfo staticCell in sInfo.StaticCells)
		{
			if (staticCell != null && staticCell.CellTemplate != null && staticCell.CellTemplate.RowIndex > 0 && staticCell.CellTemplate.ColIndex > 0)
			{
				int rowIndex = staticCell.CellTemplate.RowIndex;
				int num2 = MapTemplateExpandedRow(rowIndex, sInfo, selectedPage);
				bool flag = rowIndex <= num;
				SetPageRowVisible(sView, num2 - 1, flag, state);
				if (flag)
				{
					fpSpBindDataCell(sView, num2, staticCell.CellTemplate.ColIndex, staticCell.Controltype, staticCell.Value);
				}
			}
		}
	}

	private static int MapTemplateExpandedRow(int expandedRow, HTKKCMCSheetInfo sInfo, SheetTablePage selectedPage)
	{
		int num = expandedRow;
		if (sInfo.Tables == null)
		{
			return num;
		}
		int tableIndex;
		for (tableIndex = 0; tableIndex < sInfo.Tables.Count; tableIndex++)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[tableIndex];
			if (hTKKCMCTableInfo != null && hTKKCMCTableInfo.Location != null && hTKKCMCTableInfo.RowsValues != null)
			{
				if (hTKKCMCTableInfo.Location.DataLocationTemplateRowStart >= expandedRow)
				{
					break;
				}
				int val = selectedPage.Segments.Find((SheetTablePageSegment x) => x.TableIndex == tableIndex)?.DataCount ?? 0;
				int num2 = Math.Max(1, val);
				num -= Math.Max(0, hTKKCMCTableInfo.RowsValues.Count - num2);
			}
		}
		return num;
	}

	private static void SetPageRowVisible(SheetView sView, int rowIndex, bool visible, SheetTablePageRenderState state)
	{
		if (rowIndex >= 0 && rowIndex < sView.RowCount)
		{
			if (!state.OriginalRowVisibility.ContainsKey(rowIndex))
			{
				state.OriginalRowVisibility.Add(rowIndex, sView.Rows[rowIndex].Visible);
			}
			sView.Rows[rowIndex].Visible = visible;
		}
	}

	private static void ResetSheetTablePage(SheetView sView, FpSpread spread, int sheetIndex)
	{
		if (!SheetTablePageRenderStates.TryGetValue(spread, out var value) || !value.TryGetValue(sheetIndex, out var value2))
		{
			return;
		}
		foreach (KeyValuePair<int, bool> item in value2.OriginalRowVisibility)
		{
			if (item.Key >= 0 && item.Key < sView.RowCount)
			{
				sView.Rows[item.Key].Visible = item.Value;
			}
		}
		for (int num = value2.InsertedRows.Count - 1; num >= 0; num--)
		{
			sView.Rows.Remove(value2.InsertedRows[num].Key, value2.InsertedRows[num].Value);
		}
		value.Remove(sheetIndex);
	}

	private static void SetSheetTablePageRenderState(FpSpread spread, int sheetIndex, SheetTablePageRenderState state)
	{
		if (!SheetTablePageRenderStates.TryGetValue(spread, out var value))
		{
			value = new Dictionary<int, SheetTablePageRenderState>();
			SheetTablePageRenderStates.Add(spread, value);
		}
		value[sheetIndex] = state;
	}

	public static string GetCatalogueDMNNKDItem(string textSearch)
	{
		//IL_0031: Unknown result type (might be due to invalid IL or missing references)
		//IL_0037: Expected O, but got Unknown
		//IL_0061: Unknown result type (might be due to invalid IL or missing references)
		//IL_0067: Expected O, but got Unknown
		try
		{
			if (string.IsNullOrEmpty(textSearch))
			{
				return null;
			}
			if (!File.Exists("..\\InterfaceIni\\Catalogue_DM_NNKD.xml"))
			{
				return null;
			}
			XmlDocument val = new XmlDocument();
			val.Load("..\\InterfaceIni\\Catalogue_DM_NNKD.xml");
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				XmlAttribute val3 = val2.Attributes["Value"];
				if (val3 != null && ((XmlNode)val3).Value.IndexOf(textSearch, StringComparison.Ordinal) >= 0)
				{
					return ((XmlNode)val3).Value;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return null;
	}

	public static int GetColFromPath_Dinamic(HTKKCMCTableInfo table, string path)
	{
		HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == path);
		return hTKKCMCTableCellInfo.CellTemplate.ColIndex;
	}

	public static int GetIndexColFromPath_Dinamic(HTKKCMCTableInfo table, string path)
	{
		return table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == path);
	}
}
