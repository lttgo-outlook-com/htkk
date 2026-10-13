using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using HTKKDanhmuc;
using HTKKSerializer;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ttdb02TT80Nckt251
{
	private Dictionary<string, string[]> dicDMThueDB = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMThueDBTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDichVuThueDB = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDichVuThueDBTmp = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private int[] cell11c = GeneralFunctions.ArrayCellId("M_30");

	private int[] cell11cMa = GeneralFunctions.ArrayCellId("BA_30");

	private int[] cell11b = GeneralFunctions.ArrayCellId("Z_30");

	private int[] cell11bMa = GeneralFunctions.ArrayCellId("BB_30");

	private int[] cell11a = GeneralFunctions.ArrayCellId("AK_30");

	private int[] cell11aMa = GeneralFunctions.ArrayCellId("BC_30");

	private int col2 = HTKKFile.ColumnLettersToIndex("G");

	private int idxCol_F => HTKKFile.ColumnLettersToIndex("F") - 1;

	private int idxCol_G => HTKKFile.ColumnLettersToIndex("G") - 1;

	private int idxCol_N => HTKKFile.ColumnLettersToIndex("N") - 1;

	private int idxCol_O => HTKKFile.ColumnLettersToIndex("O") - 1;

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			UpdateDanhMucXang(sView, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			ChangeCobobox_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, iRow, iCol);
			int num = HTKKFile.ColumnLettersToIndex("G");
			int num2 = HTKKFile.ColumnLettersToIndex("AB");
			int num3 = HTKKFile.ColumnLettersToIndex("AH");
			int num4 = HTKKFile.ColumnLettersToIndex("AN");
			int num5 = HTKKFile.ColumnLettersToIndex("AP");
			int num6 = HTKKFile.ColumnLettersToIndex("AU");
			int num7 = HTKKFile.ColumnLettersToIndex("BE");
			int num8 = HTKKFile.ColumnLettersToIndex("BF");
			int num9 = HTKKFile.ColumnLettersToIndex("BG");
			int num10 = HTKKFile.ColumnLettersToIndex("AK");
			winFpSpread.Sheets[0].Cells[(int)iRow - 1, num10 - 1].Value = winFpSpread.Sheets[0].Cells[(int)iRow - 1, num7 - 1].Value;
			winFpSpread.Sheets[0].Cells[(int)iRow - 1, num5 - 1].Value = winFpSpread.Sheets[0].Cells[(int)iRow - 1, num8 - 1].Value;
			winFpSpread.Sheets[0].Cells[(int)iRow - 1, num6 - 1].Value = winFpSpread.Sheets[0].Cells[(int)iRow - 1, num9 - 1].Value;
			int num11 = 0;
			if (iCol == col2)
			{
				string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34"));
				string text = sheetView.Cells[(int)iRow - 1, col2 - 1].Text.Trim();
				if (text.StartsWith("-----") || text.StartsWith("---"))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, col2 - 1, message);
					num11 = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, col2 - 1);
				}
				if (num11 == 2)
				{
					winFpSpread.TabStrip[0].BackColor = Color.Pink;
				}
				else
				{
					winFpSpread.TabStrip[0].BackColor = Color.White;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			bool result = true;
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("M_45");
			string pathCt10 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct10";
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28"));
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt10));
			string text = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text) && text.Length != 5 && !ValidateBase.CheckMaSoThue(text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message);
				num = 2;
			}
			else if (!string.IsNullOrEmpty(text) && text.Length == 5 && !ValidateBase.CheckMaSoThue5So(text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array2 = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
			for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsValues.Count; num3++)
			{
				array2[3] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + num3, idxCol_G].Value);
				array2[5] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + num3, idxCol_N].Value);
				double num4 = Convert.ToDouble(sheetView.Cells[num2 + num3, idxCol_O].Value);
				HTKKFile.RemoveValidate(sheetView, num2 + num3, idxCol_G);
				if (string.IsNullOrEmpty(array2[3]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + num3, idxCol_G, message6);
					num = 2;
				}
				string hhdv = array2[3] + array2[5];
				if (hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[3] + p.RowValue[5] == hhdv) > 1)
				{
					XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + num3, idxCol_G, xMLMessage.GetMessage("0261"));
					num = 2;
				}
				string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34"));
				string text2 = sheetView.Cells[num2 + num3, idxCol_G].Text.Trim();
				if (text2.StartsWith("-----"))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + num3, idxCol_G, message7);
					num = 2;
				}
			}
			switch (num)
			{
			case 2:
				winFpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public Dictionary<string, string> GetDMThueTieuThuDB(string pathFile)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		//IL_0079: Unknown result type (might be due to invalid IL or missing references)
		//IL_0080: Expected O, but got Unknown
		try
		{
			if (dicDMThueDBTmp.Count > 0)
			{
				return dicDMThueDBTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(pathFile))
			{
				return null;
			}
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 2 && array[1].StartsWith("1") && array[^1] == "1")
				{
					string key = array[1];
					string text = array[2].Trim();
					string text2 = array[4].Trim();
					if (!string.IsNullOrEmpty(text2))
					{
						text = text + " (" + text2 + " %)";
					}
					if (!dictionary2.ContainsKey(key))
					{
						dictionary2.Add(key, text);
					}
					if (!dictionary.ContainsKey(text))
					{
						dictionary.Add(text, array);
					}
				}
			}
			dicDMThueDB = dictionary;
			dicDMThueDBTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMThueTieuThuDBNew(List<PathDm> pathDms)
	{
		//IL_04bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_04c2: Expected O, but got Unknown
		//IL_04f0: Unknown result type (might be due to invalid IL or missing references)
		//IL_04f7: Expected O, but got Unknown
		try
		{
			if (dicDMThueDBTmp.Count > 0)
			{
				return dicDMThueDBTmp;
			}
			Dictionary<string, string[]> temp = new Dictionary<string, string[]>();
			Dictionary<string, string> dic = new Dictionary<string, string>();
			Serializer serializer = new Serializer();
			if (pathDms == null || pathDms.Count == 0)
			{
				return null;
			}
			foreach (PathDm pathDm in pathDms)
			{
				if (pathDm == null)
				{
					continue;
				}
				string path;
				if (pathDm.dmNew)
				{
					if (HTKKFormater.isNullString(pathDm.pathDataDm) || HTKKFormater.isNullString(pathDm.pathSaveInt))
					{
						continue;
					}
					path = pathDm.pathSaveDataFile;
					if (!File.Exists(path))
					{
						path = pathDm.pathSaveInt;
					}
					string input = File.ReadAllText(path);
					dmconfig dmconfig = serializer.Deserialize<dmconfig>(input);
					input = File.ReadAllText(pathDm.pathDataDm);
					danhmuc danhmuc = serializer.Deserialize<danhmuc>(input);
					if (danhmuc == null || danhmuc.items.Count <= 0 || dmconfig == null || dmconfig.dmcfs.Count <= 0)
					{
						continue;
					}
					List<item> items = danhmuc.items;
					dmcf dmcf = null;
					foreach (dmcf dmcf2 in dmconfig.dmcfs)
					{
						if (!HTKKFormater.isNullString(dmcf2.type) && dmcf2.type.Equals(danhmuc.type))
						{
							dmcf = dmcf2;
							break;
						}
					}
					if (dmcf == null || dmcf.cells == null || dmcf.cells.Count == 0)
					{
						return null;
					}
					for (int i = 0; i < items.Count; i++)
					{
						item item = items[i];
						item item2 = item;
						if (HTKKFormater.isNullString(dmcf.active))
						{
							if (!HTKKFormater.isNullString(item.check) && Convert.ToBoolean(item.check) && item.ma.StartsWith("1"))
							{
								string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
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
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1 && item.ma.StartsWith("1") && item.check.Equals("true"))
						{
							string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
							string key = array[1];
							string text = array[2];
							if (!string.IsNullOrEmpty(item.thuesuat))
							{
								text = text + " (" + item.thuesuat + " %)";
							}
							GeneralFunctions.GetParentDanhMuc(ref dic, ref temp, item.macha, danhmuc.items);
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
					continue;
				}
				path = pathDm.pathSaveDataFileOld;
				if (!File.Exists(path))
				{
					path = pathDm.pathSaveIntOld;
				}
				XmlDocument val = new XmlDocument();
				val.Load(path);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				foreach (XmlNode item3 in elementsByTagName)
				{
					XmlNode val2 = item3;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					if (array.Length > 2 && array[1].StartsWith("1") && array[^1] == "1")
					{
						string key = array[1];
						string text = array[2].Trim();
						string text2 = array[4].Trim();
						if (!string.IsNullOrEmpty(text2))
						{
							text = text + " (" + text2 + " %)";
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
			dicDMThueDB = temp;
			dicDMThueDBTmp = dic;
			return dic;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDichVuThueTieuThuDB(string pathFile)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		//IL_0079: Unknown result type (might be due to invalid IL or missing references)
		//IL_0080: Expected O, but got Unknown
		try
		{
			if (dicDichVuThueDBTmp.Count > 0)
			{
				return dicDichVuThueDBTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(pathFile))
			{
				return null;
			}
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 2 && array[1].StartsWith("2") && array[^1] == "1")
				{
					string key = array[1];
					string text = array[2].Trim();
					if (!dictionary2.ContainsKey(key))
					{
						dictionary2.Add(key, text);
					}
					if (!dictionary.ContainsKey(text))
					{
						dictionary.Add(text, array);
					}
				}
			}
			dicDichVuThueDB = dictionary;
			dicDichVuThueDBTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDichVuThueTieuThuDBNew(List<PathDm> pathDms)
	{
		//IL_0477: Unknown result type (might be due to invalid IL or missing references)
		//IL_047e: Expected O, but got Unknown
		//IL_04ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_04b3: Expected O, but got Unknown
		try
		{
			if (dicDichVuThueDBTmp.Count > 0)
			{
				return dicDichVuThueDBTmp;
			}
			Dictionary<string, string[]> temp = new Dictionary<string, string[]>();
			Dictionary<string, string> dic = new Dictionary<string, string>();
			Serializer serializer = new Serializer();
			if (pathDms == null || pathDms.Count == 0)
			{
				return null;
			}
			foreach (PathDm pathDm in pathDms)
			{
				if (pathDm == null)
				{
					continue;
				}
				string path;
				if (pathDm.dmNew)
				{
					if (HTKKFormater.isNullString(pathDm.pathDataDm) || HTKKFormater.isNullString(pathDm.pathSaveInt))
					{
						continue;
					}
					path = pathDm.pathSaveDataFile;
					if (!File.Exists(path))
					{
						path = pathDm.pathSaveInt;
					}
					string input = File.ReadAllText(path);
					dmconfig dmconfig = serializer.Deserialize<dmconfig>(input);
					input = File.ReadAllText(pathDm.pathDataDm);
					danhmuc danhmuc = serializer.Deserialize<danhmuc>(input);
					if (danhmuc == null || danhmuc.items.Count <= 0 || dmconfig == null || dmconfig.dmcfs.Count <= 0)
					{
						continue;
					}
					List<item> items = danhmuc.items;
					dmcf dmcf = null;
					foreach (dmcf dmcf2 in dmconfig.dmcfs)
					{
						if (!HTKKFormater.isNullString(dmcf2.type) && dmcf2.type.Equals(danhmuc.type))
						{
							dmcf = dmcf2;
							break;
						}
					}
					if (dmcf == null || dmcf.cells == null || dmcf.cells.Count == 0)
					{
						return null;
					}
					for (int i = 0; i < items.Count; i++)
					{
						item item = items[i];
						item item2 = item;
						if (HTKKFormater.isNullString(dmcf.active))
						{
							if (!HTKKFormater.isNullString(item.check) && Convert.ToBoolean(item.check) && item.ma.StartsWith("2"))
							{
								string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
								string key = array[1];
								string text = array[2];
								GeneralFunctions.GetParentDanhMuc(ref dic, ref temp, item.macha, danhmuc.items);
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
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1 && item.ma.StartsWith("2") && item.check.Equals("true"))
						{
							string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
							string key = array[1];
							string text = array[2];
							GeneralFunctions.GetParentDanhMuc(ref dic, ref temp, item.macha, danhmuc.items);
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
					continue;
				}
				path = pathDm.pathSaveDataFileOld;
				if (!File.Exists(path))
				{
					path = pathDm.pathSaveIntOld;
				}
				XmlDocument val = new XmlDocument();
				val.Load(path);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				foreach (XmlNode item3 in elementsByTagName)
				{
					XmlNode val2 = item3;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					if (array.Length > 2 && array[1].StartsWith("2") && array[^1] == "1")
					{
						string key = array[1];
						string text = array[2].Trim();
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
			dicDichVuThueDB = temp;
			dicDichVuThueDBTmp = dic;
			return dic;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMTTDBDonVi(string pathFile)
	{
		//IL_002c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0061: Expected O, but got Unknown
		try
		{
			if (dicDMDonVi.Count > 0)
			{
				return dicDMDonVi;
			}
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			int num = 0;
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["value"]).Value;
				num++;
				dictionary.Add(num.ToString(), value);
			}
			dicDMDonVi = dictionary;
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("BD");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sheetView.Cells[num2 + i, idxCol_G].Text;
				string[] array2 = dicDMThueDB[text];
				string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array2[1], "TTDB");
				sheetView.Cells[num2 + i, num - 1].Value = maTieuMucByMaHH;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateDanhMucXang(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		List<HTKKCMCTableInfo> tables = ToKhaiInfo.SheetInfos[0].Tables;
		Dictionary<string, string> danhMucXang = GetDanhMucXang(sView, ToKhaiInfo);
		if (danhMucXang.Count <= 0)
		{
			return;
		}
		foreach (HTKKCMCTableInfo item in tables)
		{
			int dataLocationTemplateRowStart = item.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < item.RowsValues.Count; i++)
			{
				GeneralFunctions.InsertCombobox(sView, dataLocationTemplateRowStart + i, col2, danhMucXang, deleteText: false);
			}
		}
	}

	public Dictionary<string, string> GetDanhMucXang(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		List<HTKKCMCTableInfo> tables = ToKhaiInfo.SheetInfos[0].Tables;
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (HTKKCMCTableInfo item in tables)
		{
			int dataLocationTemplateRowStart = item.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < item.RowsValues.Count; i++)
			{
				Dictionary<string, string> dMFromCombobox = GeneralFunctions.GetDMFromCombobox(sView, dataLocationTemplateRowStart + i, col2);
				if (dMFromCombobox.Count <= 0)
				{
					continue;
				}
				foreach (KeyValuePair<string, string> item2 in dMFromCombobox)
				{
					if ((item2.Key == "10501" || item2.Key == "10502" || item2.Key == "10503" || item2.Key == "10504" || item2.Key == "10505" || item2.Key == "10506") && !dictionary.ContainsKey(item2.Key))
					{
						dictionary.Add(item2.Key, item2.Value);
					}
				}
			}
		}
		return dictionary;
	}

	public void ChangeCobobox_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("G");
			if (iCol != num)
			{
				return;
			}
			int num2 = HTKKFile.ColumnLettersToIndex("AZ");
			int num3 = HTKKFile.ColumnLettersToIndex("N");
			int num4 = HTKKFile.ColumnLettersToIndex("AN");
			int num5 = HTKKFile.ColumnLettersToIndex("BD");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			if (iRow >= dataLocationTemplateRowStart)
			{
				string text = sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim();
				if (HTKKFormater.isNullString(text))
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].ResetValue();
					sheetView.Cells[(int)iRow - 1, num3 - 1].ResetValue();
					sheetView.Cells[(int)iRow - 1, num3 - 1].Locked = true;
					sheetView.Cells[(int)iRow - 1, num4 - 1].Value = 0;
				}
				else if (!dicDMThueDB.ContainsKey(text))
				{
					sheetView.Cells[(int)iRow - 1, num - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num3 - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, num4 - 1].Value = 0;
					sheetView.Cells[(int)iRow - 1, num5 - 1].Value = "";
				}
				else
				{
					string[] array = dicDMThueDB[text];
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[1];
					sheetView.Cells[(int)iRow - 1, num3 - 1].Locked = false;
					sheetView.Cells[(int)iRow - 1, num3 - 1].Value = array[3];
					sheetView.Cells[(int)iRow - 1, num4 - 1].Value = ConvertHelpers.ConvertToDouble(array[4]);
					string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array[1], "TTDB");
					sheetView.Cells[(int)iRow - 1, num5 - 1].Value = maTieuMucByMaHH;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
