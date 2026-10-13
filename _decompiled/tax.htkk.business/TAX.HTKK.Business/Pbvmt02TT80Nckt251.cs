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

public class Pbvmt02TT80Nckt251
{
	private Dictionary<string, string[]> dicDMKhoangSan = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMKhoangSanTmp = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMKSBVMT = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMKSBVMTTmp = new Dictionary<string, string>();

	private List<item> _dmKSData = new List<item>();

	private List<item> _dmBVMTData = new List<item>();

	private bool _isDmMoi = false;

	private string namKK = "";

	private string kieuKyKK = "";

	private int[] cell09c = GeneralFunctions.ArrayCellId("G_46");

	private int[] cell09cMa = GeneralFunctions.ArrayCellId("AV_46");

	private int[] cell09b = GeneralFunctions.ArrayCellId("G_48");

	private int[] cell09bMa = GeneralFunctions.ArrayCellId("AV_48");

	private int[] cell09a = GeneralFunctions.ArrayCellId("G_50");

	private int[] cell09aMa = GeneralFunctions.ArrayCellId("AV_50");

	private int colMucPhi200 = HTKKFile.ColumnLettersToIndex("AH");

	private int col2 = HTKKFile.ColumnLettersToIndex("C");

	private int col4 = HTKKFile.ColumnLettersToIndex("L");

	private int col2Ma = HTKKFile.ColumnLettersToIndex("AV");

	private int colTieuMuc = HTKKFile.ColumnLettersToIndex("AW");

	private int[] cellTenTHQT_TK = GeneralFunctions.ArrayCellId("I_15");

	private int[] cellMaTHQT_TK = GeneralFunctions.ArrayCellId("H_15");

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
		if (!string.IsNullOrEmpty(TTinChung.maCQTQuanLy) && headerInfo.maCQTQuanLy != TTinChung.maCQTQuanLy)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[1];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			if (hTKKCMCTableInfo.RowsValues.Count > 1)
			{
				GeneralFunctions.RemoveDataRows(sView, 1, dataLocationTemplateRowStart - 1, hTKKCMCTableInfo.RowsValues.Count - 1, ToKhaiInfo.SheetInfos[0]);
			}
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo.RowsInfo.Cells)
				{
					sView.Cells[dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, cell.CellTemplate.ColIndex, locked: true, deleteData: false);
				}
			}
		}
		GeneralFunctions.ThayDoiTinhTheoCQT(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sView, PeriodInfo, cell09c, cell09cMa, cell09b, cell09bMa, cell09a, cell09aMa, null, null, null);
		if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
		{
			sView.Cells[cellTenTHQT_TK[1], cellTenTHQT_TK[0]].Value = PeriodInfo.TenQuyetToan;
			sView.Cells[cellMaTHQT_TK[1], cellMaTHQT_TK[0]].Value = PeriodInfo.MaQuyetToan;
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			if (iCol != col2)
			{
				return;
			}
			bool flag = false;
			if (PeriodInfo.dicVar != null && PeriodInfo.dicVar.Count > 0 && PeriodInfo.dicVar.ContainsKey("IS_DM_MOI"))
			{
				string text = PeriodInfo.dicVar["IS_DM_MOI"].ToString();
				if (text == "1")
				{
					flag = true;
				}
			}
			string[] array = ((!flag) ? dicDMKhoangSan[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()] : dicDMKSBVMT[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()]);
			if (array.Length > 2)
			{
				sheetView.Cells[(int)iRow - 1, col2Ma - 1].Value = array[1];
				string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array[1], "PBVMT");
				sheetView.Cells[(int)iRow - 1, colTieuMuc - 1].Value = maTieuMucByMaHH;
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("V");
			int num3 = HTKKFile.ColumnLettersToIndex("Y");
			int num4 = HTKKFile.ColumnLettersToIndex("AV");
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num5 = 0;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32"));
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			if (GeneralFunctions.CheckTinhTheoCQT(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sheetView, PeriodInfo, cell09c, cell09cMa, cell09b, cell09bMa, cell09a, cell09aMa, null))
			{
				num5 = 2;
			}
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < table.RowsValues.Count; i++)
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num - 1);
					string item = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
					double value = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Value);
					double num6 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Value);
					List<int> list = new List<int>();
					list.Add(num);
					list.Add(num2);
					list.Add(num3);
					List<string> list2 = new List<string>();
					list2.Add(item);
					list2.Add(value.ToString());
					list2.Add(num6.ToString());
					List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, table, list, list2);
					if (rowIndexInSViewByCompares.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, message2);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1, message2);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, message2);
						num5 = 2;
					}
					string maLoaiKhoangSan = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Text.Trim();
					if (string.IsNullOrEmpty(maLoaiKhoangSan))
					{
						continue;
					}
					string[] array = dicDMKSBVMT.Values.FirstOrDefault((string[] x) => x[1] == maLoaiKhoangSan);
					if (array != null && array.Count() > 9 && Convert.ToDouble(array[8]) > 0.0 && Convert.ToDouble(array[9]) > 0.0 && (Convert.ToDouble(value) > Convert.ToDouble(array[9]) || Convert.ToDouble(value) < Convert.ToDouble(array[8])))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1, message3);
						if (num5 != 2)
						{
							num5 = 1;
						}
					}
				}
			}
			switch (num5)
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

	public Dictionary<string, string> GetDMKhoangSan(string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		try
		{
			if (dicDMKhoangSanTmp.Count > 0)
			{
				return dicDMKhoangSanTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[1];
				string text = array[2];
				if (array[^1] == "1")
				{
					dictionary2.Add(key, text);
					dictionary.Add(text, array);
				}
			}
			dicDMKhoangSan = dictionary;
			dicDMKhoangSanTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMKhoangSanNew(List<PathDm> pathDms)
	{
		//IL_02ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f4: Expected O, but got Unknown
		//IL_0322: Unknown result type (might be due to invalid IL or missing references)
		//IL_0329: Expected O, but got Unknown
		try
		{
			_isDmMoi = true;
			if (dicDMKhoangSanTmp.Count > 0)
			{
				return dicDMKhoangSanTmp;
			}
			_dmKSData = new List<item>();
			List<item> list = new List<item>();
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
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
							if (!HTKKFormater.isNullString(item.check) && Convert.ToBoolean(item.check))
							{
								list.Add(item);
							}
						}
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1)
						{
							list.Add(item);
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
				foreach (XmlNode item4 in elementsByTagName)
				{
					XmlNode val2 = item4;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					item item3 = new item();
					item3.id = array[0];
					item3.ma = array[1];
					item3.ten = array[2];
					item3.dvt = array[3];
					item3.thuesuat = array[4];
					list.Add(item3);
				}
			}
			if (list != null && list.Count > 0)
			{
				foreach (item item5 in list)
				{
					string ten = item5.ten;
					List<item> list2 = list.FindAll((item obj) => obj.ten.Equals(ten));
					if (list2 != null && list2.Count > 1)
					{
						ten = ten + "( " + item5.ma + " )";
					}
					string text = item5.id + "_" + item5.ma + "_" + item5.thuesuat + "_" + item5.giatoithieu + "_" + item5.cqtgia;
					string[] array = new string[7] { text, item5.ma, ten, item5.dvt, item5.thuesuat, "0", "1" };
					item5.id = text;
					if (!dictionary.ContainsKey(text))
					{
						string value2 = array[2];
						dictionary2.Add(text, value2);
						dictionary.Add(text, array);
						_dmKSData.Add(item5);
					}
				}
			}
			dicDMKhoangSan = dictionary;
			dicDMKhoangSanTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public List<item> GetItemDm_GetDMKhoangSan()
	{
		return _dmKSData;
	}

	public Dictionary<string, string> GetDMDonVi(string pathFile)
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

	public Dictionary<string, string> GetDMKSBVMT(string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		try
		{
			if (dicDMKhoangSanTmp.Count > 0)
			{
				return dicDMKhoangSanTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[1];
				string text = array[2];
				dictionary2.Add(key, text);
				dictionary.Add(text, array);
			}
			dicDMKhoangSan = dictionary;
			dicDMKhoangSanTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMKSBVMTNew(List<PathDm> pathDms)
	{
		//IL_02ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f4: Expected O, but got Unknown
		//IL_0322: Unknown result type (might be due to invalid IL or missing references)
		//IL_0329: Expected O, but got Unknown
		try
		{
			_isDmMoi = true;
			if (dicDMKSBVMTTmp.Count > 0)
			{
				return dicDMKSBVMTTmp;
			}
			_dmBVMTData = new List<item>();
			List<item> list = new List<item>();
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
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
							if (!HTKKFormater.isNullString(item.check) && Convert.ToBoolean(item.check))
							{
								list.Add(item);
							}
						}
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1)
						{
							list.Add(item);
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
				foreach (XmlNode item4 in elementsByTagName)
				{
					XmlNode val2 = item4;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					if (array[^1] == "1")
					{
						item item3 = new item();
						item3.id = array[0];
						item3.ma = array[1];
						item3.ten = array[2];
						item3.dvt = array[3];
						item3.thuesuat = array[4];
						if (!namKK.Equals("2019") || !kieuKyKK.Equals("Y"))
						{
							list.Add(item3);
						}
					}
				}
			}
			if (list != null && list.Count > 0)
			{
				foreach (item item5 in list)
				{
					string ten = item5.ten;
					List<item> list2 = list.FindAll((item obj) => obj.ten.Equals(ten));
					if (list2 != null && list2.Count > 1)
					{
						ten = ten + "( " + item5.ma + " )";
					}
					string text = item5.id + "_" + item5.ma + "_" + item5.thuesuat + "_" + item5.giatoithieu + "_" + item5.cqtgia;
					string[] array = new string[10] { text, item5.ma, ten, item5.dvt, item5.thuesuat, "0", "1", item5.matieumuc, item5.giatoithieu, item5.giatoida };
					item5.id = text;
					if (!dictionary.ContainsKey(text))
					{
						string text2 = array[2];
						dictionary2.Add(text, text2);
						if (!dictionary.ContainsKey(text2))
						{
							dictionary.Add(text2, array);
						}
						_dmBVMTData.Add(item5);
					}
				}
			}
			dicDMKSBVMT = dictionary;
			dicDMKSBVMTTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public List<item> GetItemDm_GetDMKSBVMT()
	{
		return _dmBVMTData;
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AW");
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			bool flag = false;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num3 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					if (PeriodInfo.dicVar != null && PeriodInfo.dicVar.Count > 0 && PeriodInfo.dicVar.ContainsKey("IS_DM_MOI"))
					{
						string text = PeriodInfo.dicVar["IS_DM_MOI"].ToString();
						if (text == "1")
						{
							flag = true;
						}
					}
					string[] array = ((!flag) ? dicDMKhoangSan[sheetView.Cells[num3 + j, num2 - 1].Text.Trim()] : dicDMKSBVMT[sheetView.Cells[num3 + j, num2 - 1].Text.Trim()]);
					if (array.Length > 2)
					{
						string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array[1], "PBVMT");
						sheetView.Cells[num3 + j, num - 1].Value = maTieuMucByMaHH;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
