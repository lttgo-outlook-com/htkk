using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using HTKKDanhmuc;
using HTKKSerializer;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tain01Nckt202
{
	private Dictionary<string, string[]> dicDMKhoangSan = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMKhoangSanTmp = new Dictionary<string, string>();

	private List<item> _dmKSData = new List<item>();

	private List<string> pathDonGiaTinh = new List<string>();

	private bool _isDmMoi = false;

	private static string fmlCt19 = "R40+AD40";

	private static string fmlCt20 = "R41+AD41";

	private static string fmlCt21 = "R42+AD42";

	private static string fmlCt22 = "R43+AD43";

	private static string fmlCt23 = "R44+AD44";

	private static string fmlCt24 = "R46+AD46";

	private static string fmlCt25 = "R47+AD47";

	private static string fmlCt26 = "R48+AD48";

	private static string fmlCt27 = "R49+AD49";

	private static string fmlCt28 = "R50+AD50";

	private static string fmlCt29 = "R52+AD52";

	private static string fmlCt31 = "R54+AD54";

	private static string fmlCt33 = "R57+AD57";

	private static string fmlCt34 = "R58+AD58";

	private static string fmlCt35 = "R59+AD59";

	private static string fmlCt36 = "R60+AD60";

	private static string fmlCt37 = "R61+AD61";

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (hTKKCMCSheetInfo.Tables.Count <= 0)
			{
				return;
			}
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				if ((!(PeriodInfo.KieuKy == "M") || i != 2) && (!(PeriodInfo.KieuKy == "D") || i >= 2))
				{
					continue;
				}
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num = ConvertHelpers.ConvertToInt(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart);
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				List<HTKKCMCTableCellInfo> cells = hTKKCMCTableInfo.RowsInfo.Cells;
				if (rowsValues.Count <= 0)
				{
					continue;
				}
				for (int j = 0; j < rowsValues.Count; j++)
				{
					int num2 = num + j;
					if (cells.Count > 0)
					{
						for (int k = 0; k < cells.Count; k++)
						{
							sView.Cells[num2 - 1, cells[k].CellTemplate.ColIndex - 1].Locked = true;
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void SetFormula(SheetView sView)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void ResetFormula(SheetView sViewTK)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = HTKKFile.ColumnLettersToIndex("P");
			int num4 = HTKKFile.ColumnLettersToIndex("V");
			int num5 = HTKKFile.ColumnLettersToIndex("AA");
			int num6 = HTKKFile.ColumnLettersToIndex("AD");
			if (iCol == num2)
			{
				winFpSpread.TabStrip[0].BackColor = Color.White;
				HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, num2 - 1);
				foreach (KeyValuePair<string, string[]> item in dicDMKhoangSan)
				{
					string[] value = item.Value;
					if (value.Length > 3 && value[2].Trim() == sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim())
					{
						if (ConvertHelpers.ConvertToInt(value[4]) == 0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, num2 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
							winFpSpread.TabStrip[0].BackColor = Color.Pink;
						}
						sheetView.Cells[(int)iRow - 1, num - 1].Text = item.Key;
						sheetView.Cells[(int)iRow - 1, num3 - 1].Text = value[3];
						if (ConvertHelpers.ConvertToDouble1(sheetView.Cells[(int)iRow - 1, num6 - 1].Text) > 0.0)
						{
							sheetView.Cells[(int)iRow - 1, num5 - 1].Text = "0";
						}
						else
						{
							sheetView.Cells[(int)iRow - 1, num5 - 1].Text = value[4];
						}
						if (value[5] == "1" || value[1] == "1101")
						{
							sheetView.Cells[(int)iRow - 1, num5 - 1].Locked = false;
						}
						else
						{
							sheetView.Cells[(int)iRow - 1, num5 - 1].Locked = true;
						}
						sheetView.Cells[(int)iRow - 1, num3 - 1].Locked = false;
						double donGia = getDonGia(item.Key);
						sheetView.Cells[(int)iRow - 1, num4 - 1].Value = donGia;
						break;
					}
				}
			}
			if (iCol == num4)
			{
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, (int)iCol, "NumberCellType");
				if (ConvertHelpers.ConvertToDouble(valueCustom) > 0.0)
				{
					sheetView.Cells[(int)iRow - 1, num6 - 1].Locked = true;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num6 - 1].Locked = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = HTKKFile.ColumnLettersToIndex("P");
			int num4 = HTKKFile.ColumnLettersToIndex("V");
			int num5 = HTKKFile.ColumnLettersToIndex("AA");
			int num6 = HTKKFile.ColumnLettersToIndex("AD");
			if (iCol == num4)
			{
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, (int)iCol, "NumberCellType");
				if (ConvertHelpers.ConvertToDouble(valueCustom) > 0.0)
				{
					sheetView.Cells[(int)iRow - 1, num6 - 1].Text = "0";
					sheetView.Cells[(int)iRow - 1, num6 - 1].Locked = false;
					foreach (KeyValuePair<string, string[]> item in dicDMKhoangSan)
					{
						string[] value = item.Value;
						if (value.Length > 3 && value[2].Trim() == sheetView.Cells[(int)iRow - 1, num2 - 1].Text.Trim())
						{
							sheetView.Cells[(int)iRow - 1, num5 - 1].Text = value[4];
							break;
						}
					}
				}
				validateTaiNguyen(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, iRow, iCol);
			}
			if (iCol == num6 && ConvertHelpers.ConvertToDouble(sheetView.Cells[(int)iRow - 1, num6 - 1].Value) > 0.0)
			{
				sheetView.Cells[(int)iRow - 1, num5 - 1].Text = "0";
				sheetView.Cells[(int)iRow - 1, num4 - 1].Text = "0";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = HTKKFile.ColumnLettersToIndex("P");
			if (iCol == num3 && HTKKFormater.isNullString(sheetView.Cells[(int)iRow - 1, num - 1].Text))
			{
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Locked = true;
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
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		//IL_063a: Unknown result type (might be due to invalid IL or missing references)
		//IL_06d5: Unknown result type (might be due to invalid IL or missing references)
		int value = 0;
		levelWarning.TryGetValue(0, out value);
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("E");
			int iCol = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = HTKKFile.ColumnLettersToIndex("P");
			int num4 = HTKKFile.ColumnLettersToIndex("S");
			int num5 = HTKKFile.ColumnLettersToIndex("V");
			int iCol2 = HTKKFile.ColumnLettersToIndex("AA");
			int iCol3 = HTKKFile.ColumnLettersToIndex("AD");
			int iCol4 = HTKKFile.ColumnLettersToIndex("AG");
			int num6 = HTKKFile.ColumnLettersToIndex("AJ");
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			bool flag = false;
			Dictionary<int, int[]> dictionary2 = new Dictionary<int, int[]>();
			List<HTKKCMCTableInfo> tables = ToKhaiInfo.SheetInfos[0].Tables;
			for (int i = 0; i < tables.Count; i++)
			{
				int dataLocationTemplateRowStart = tables[i].Location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = tables[i].RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string valColMa = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol, "TextCellType");
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num4, "NumberCellType");
					string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num5, "NumberCellType");
					string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol2, "NumberCellType");
					string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol3, "NumberCellType");
					string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num6, "NumberCellType");
					string valueCustom6 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol4, "NumberCellType");
					string valueCustom7 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num2, "0");
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
					if (!HTKKFormater.isNullString(valColMa) && _dmKSData != null && _dmKSData.Count > 0)
					{
						item item = _dmKSData.Find((item x) => x.id.Contains(valColMa));
						if (item != null && (HTKKFormater.isNullString(item.thuesuat) || ConvertHelpers.ConvertToInt(item.thuesuat) == 0))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
							value = 2;
						}
					}
					if (string.IsNullOrEmpty(valueCustom7) && (ConvertHelpers.ConvertToDouble(valueCustom) > 0.0 || ConvertHelpers.ConvertToDouble(valueCustom2) > 0.0 || ConvertHelpers.ConvertToDouble(valueCustom3) > 0.0 || ConvertHelpers.ConvertToDouble(valueCustom4) > 0.0 || ConvertHelpers.ConvertToDouble(valueCustom5) > 0.0))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
						value = 2;
					}
					if (ConvertHelpers.ConvertToDouble(valueCustom5) > ConvertHelpers.ConvertToDouble(valueCustom6))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1);
					}
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1);
					if (!HTKKFormater.isNullString(valColMa) && _dmKSData != null && _dmKSData.Count > 0 && valueCustom4.ToNumber() == 0m && valueCustom2.ToNumber() != 0m)
					{
						bool flag2 = true;
						item item = _dmKSData.Find((item x) => x.id.Contains(valColMa));
						if (item != null)
						{
							if (item.cqtgia.ToNumber() > 0m)
							{
								flag2 = false;
								if (valueCustom2.ToNumber() < item.cqtgia.ToNumber())
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")));
									value = 2;
								}
							}
							if (flag2 && valueCustom2.ToNumber() < item.giatoithieu.ToNumber())
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
								value = 2;
							}
						}
					}
					if (!HTKKFormater.isNullString(valueCustom7) && !dictionary.ContainsKey(i))
					{
						dictionary.Add(i, j);
					}
					if (!string.IsNullOrEmpty(valueCustom7) && ConvertHelpers.ConvertToDouble(valueCustom) == 0.0)
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
						if (value == 0)
						{
							value = 1;
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1);
					}
				}
			}
			if (PeriodInfo.KieuKy == "M" && dictionary.ContainsKey(2))
			{
				winFpSpread.ResumeLayout();
				value = 2;
				Process[] processesByName = Process.GetProcessesByName("ProgressIndicatorTest");
				foreach (Process process in processesByName)
				{
					process.Kill();
				}
				frmMessage frmMessage = new frmMessage("0292", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
			}
			if (PeriodInfo.KieuKy == "D" && (dictionary.ContainsKey(1) || dictionary.ContainsKey(0)))
			{
				winFpSpread.ResumeLayout();
				value = 2;
				Process[] processesByName = Process.GetProcessesByName("ProgressIndicatorTest");
				foreach (Process process in processesByName)
				{
					process.Kill();
				}
				frmMessage frmMessage = new frmMessage("0293", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.White;
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
				string value2 = array[2];
				if (array[^1] == "1")
				{
					dictionary2.Add(key, value2);
					dictionary.Add(key, array);
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

	public Dictionary<string, string> GetDMKhoangSanNew(List<PathDm> pathDms, HTKKPeriodInfo PeriodInfo)
	{
		//IL_037c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0383: Expected O, but got Unknown
		//IL_03b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_03b8: Expected O, but got Unknown
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
					if (pathDm.pathDonGiaTinh != null && pathDm.pathDonGiaTinh.Count > 0)
					{
						foreach (string p in pathDm.pathDonGiaTinh)
						{
							List<string> list2 = pathDonGiaTinh;
							Predicate<string> match = (string o) => o.Equals(p);
							if (HTKKFormater.isNullString(list2.Find(match)))
							{
								pathDonGiaTinh.Add(p);
							}
						}
					}
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
					for (int num = 0; num < items.Count; num++)
					{
						item item = items[num];
						item item2 = item;
						if (HTKKFormater.isNullString(dmcf.active))
						{
							list.Add(item);
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
				foreach (XmlNode item7 in elementsByTagName)
				{
					XmlNode val2 = item7;
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
						list.Add(item3);
					}
				}
			}
			if (list != null && list.Count > 0)
			{
				List<item> dmTaiNguyenForTinh = GeneralFunctions.getDmTaiNguyenForTinh(pathDonGiaTinh, PeriodInfo);
				List<item> list3 = Serializer.DeepCopy(list);
				if (dmTaiNguyenForTinh != null && dmTaiNguyenForTinh.Count > 0)
				{
					foreach (item iT in list)
					{
						Predicate<item> match2 = (item obj) => obj.id.Equals(iT.id);
						List<item> list4 = dmTaiNguyenForTinh.FindAll(match2);
						if (list4 != null && list4.Count > 0)
						{
							item item4 = list3.Find((item obj) => obj.id == iT.id);
							int num2 = list3.FindIndex((item obj) => obj.id == iT.id);
							if (item4 != null)
							{
								list3.Remove(item4);
							}
							for (int num3 = 0; num3 < list4.Count; num3++)
							{
								item item5 = Serializer.DeepCopy(iT);
								item5.cqtgia = list4[num3].cqtgia;
								list3.Insert(num2 + num3, item5);
							}
						}
					}
					list = Serializer.DeepCopy(list3);
				}
				list3 = new List<item>();
				foreach (item item8 in list)
				{
					item item6 = Serializer.DeepCopy(item8);
					string ten = item6.ten;
					List<item> list5 = list.FindAll((item obj) => obj.ten.Equals(ten));
					if (list5 != null && list5.Count > 1)
					{
						item6 = GeneralFunctions.getTaiNguyenForCombobox(item6, list5);
						if (item6 == null && item8.id == list5[0].id)
						{
							list3.Add(Serializer.DeepCopy(item8));
						}
					}
					if (item6 != null)
					{
						if (!HTKKFormater.isNullString(item6.level) && HTKKFormater.IsNumber(item6.level))
						{
							item6.ten = HTKKFile.addSpace(ConvertHelpers.ConvertToInt(item6.level)) + item6.ten;
						}
						list3.Add(item6);
					}
				}
				for (int num4 = 0; num4 < list3.Count; num4++)
				{
					item item6 = Serializer.DeepCopy(list3[num4]);
					string text = item6.id + "_" + item6.ma + "_" + item6.thuesuat + "_" + item6.giatoithieu + "_" + item6.cqtgia;
					string[] array = new string[7] { text, item6.ma, item6.ten, item6.dvt, item6.thuesuat, "0", "1" };
					item6.id = text;
					if (!dictionary.ContainsKey(text))
					{
						string value2 = array[2];
						dictionary2.Add(text, value2);
						dictionary.Add(text, array);
						_dmKSData.Add(item6);
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

	public static void RemoveDataRows(SheetView sView, int frRow, int rowCount, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (rowCount == 0)
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void fpSpReindexingTableTemplate(SheetView sView, HTKKCMCTableInfo tableInfo)
	{
		try
		{
			for (int i = 0; i < tableInfo.RowsValues.Count; i++)
			{
				if (tableInfo.CellTableSTT != null)
				{
					sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].Text = tableInfo.CellTableSTT.Value + "." + (i + 1);
				}
				else
				{
					sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].Text = (i + 1).ToString();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ExportXML_0(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		//IL_0130: Unknown result type (might be due to invalid IL or missing references)
		//IL_0137: Expected O, but got Unknown
		//IL_0214: Unknown result type (might be due to invalid IL or missing references)
		//IL_021b: Expected O, but got Unknown
		try
		{
			XmlDocument val = new XmlDocument();
			string text = string.Format("{0}_xml_{1}.xml", ToKhaiInfo.InterfaceTemplateXml + ToKhaiInfo.SheetInfos[0].DataFile, ToKhaiInfo.VersionXML.Replace(".", ""));
			val.Load(text);
			XmlNode val2 = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/TaiNguyenKhaiThac");
			XmlNode val3 = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/TaiNguyenThuMuaGom");
			XmlNode val4 = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/TaiNguyenTichThuGiaoBan");
			if (val2.InnerXml == null || val2.InnerXml == "")
			{
				XmlNode val5 = genDefaultXmlNull("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/TaiNguyenKhaiThac/ThueTaiNguyenKhaiThac", val);
				val2.AppendChild(val5);
			}
			if (val3.InnerXml == null || val3.InnerXml == "")
			{
				XmlNode val5 = genDefaultXmlNull("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/TaiNguyenThuMuaGom/ThueTaiNguyenThuMuaGom", val);
				val3.InnerXml = val5.InnerXml;
				if (val3.HasChildNodes)
				{
					val5 = val3.ChildNodes[0];
					foreach (XmlAttribute item in (XmlNamedNodeMap)val5.Attributes)
					{
						XmlAttribute val6 = item;
						if (((XmlNode)val6).Name != "id")
						{
							val5.Attributes.Remove(val6);
						}
					}
				}
			}
			if (val4.InnerXml == null || val4.InnerXml == "")
			{
				XmlNode val5 = genDefaultXmlNull("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/TaiNguyenTichThuGiaoBan/ThueTaiNguyenTichThuGiaoBan", val);
				val4.InnerXml = val5.InnerXml;
				if (val4.HasChildNodes)
				{
					val5 = val4.ChildNodes[0];
					foreach (XmlAttribute item2 in (XmlNamedNodeMap)val5.Attributes)
					{
						XmlAttribute val6 = item2;
						if (((XmlNode)val6).Name != "id")
						{
							val5.Attributes.Remove(val6);
						}
					}
				}
			}
			if (PeriodInfo.dicVar != null && PeriodInfo.dicVar.Count > 0 && PeriodInfo.dicVar.ContainsKey("IS_DM_MOI"))
			{
				convertId2Ma(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/TaiNguyenKhaiThac/ThueTaiNguyenKhaiThac/ct1");
				convertId2Ma(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/TaiNguyenThuMuaGom/ThueTaiNguyenThuMuaGom/ct1");
				convertId2Ma(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/TaiNguyenTichThuGiaoBan/ThueTaiNguyenTichThuGiaoBan/ct1");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private XmlNode genDefaultXmlNull(string pathNode, XmlDocument reValues)
	{
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0045: Expected O, but got Unknown
		try
		{
			XmlNode val = ((XmlNode)reValues).SelectSingleNode(pathNode);
			((XmlNode)val.Attributes["id"]).Value = "ID_1";
			foreach (XmlNode childNode in val.ChildNodes)
			{
				XmlNode val2 = childNode;
				switch (val2.Name)
				{
				case "ct4":
					val2.InnerXml = "0.000";
					break;
				case "ct5":
					val2.InnerXml = "0.00";
					break;
				case "ct6":
					val2.InnerXml = "0.000";
					break;
				case "ct7":
					val2.InnerXml = "0";
					break;
				case "ct8":
					val2.InnerXml = "0";
					break;
				case "ct9":
					val2.InnerXml = "0.00";
					break;
				case "ct10":
					val2.InnerXml = "0";
					break;
				}
			}
			return val.ParentNode;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = HTKKFile.ColumnLettersToIndex("P");
			int iCol = HTKKFile.ColumnLettersToIndex("V");
			int num4 = HTKKFile.ColumnLettersToIndex("AA");
			int num5 = HTKKFile.ColumnLettersToIndex("AD");
			int num6 = sheetView.ActiveRow.Index + 1;
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, num6 + 1, iCol, "NumberCellType");
			if (ConvertHelpers.ConvertToDouble(valueCustom) > 0.0)
			{
				sheetView.Cells[num6, num5 - 1].Locked = true;
			}
			else
			{
				sheetView.Cells[num6, num5 - 1].Locked = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = HTKKFile.ColumnLettersToIndex("P");
			int iCol = HTKKFile.ColumnLettersToIndex("V");
			int num4 = HTKKFile.ColumnLettersToIndex("AA");
			int num5 = HTKKFile.ColumnLettersToIndex("AD");
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					foreach (KeyValuePair<string, string[]> item in dicDMKhoangSan)
					{
						string valueCustom = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol, "NumberCellType");
						if (ConvertHelpers.ConvertToDouble(valueCustom) > 0.0)
						{
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Locked = true;
						}
						else
						{
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Locked = false;
						}
						string[] value = item.Value;
						if (value.Length > 3 && value[2].Trim() == sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text.Trim())
						{
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text = item.Key;
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Text = value[3];
							if (ConvertHelpers.ConvertToDouble1(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Text) > 0.0)
							{
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text = "0";
							}
							else
							{
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text = value[4];
							}
							if (value[5] == "1" || value[1] == "1101")
							{
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Locked = false;
							}
							else
							{
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Locked = true;
							}
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Locked = false;
							break;
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public double getDonGia(string ma)
	{
		double result = 0.0;
		try
		{
			if (!HTKKFormater.isNullString(ma) && _dmKSData != null && _dmKSData.Count > 0)
			{
				item item = _dmKSData.Find((item x) => x.id.Contains(ma));
				if (item != null && item.cqtgia.ToNumber() > 0m)
				{
					result = ConvertHelpers.ConvertToDouble1(item.cqtgia.ToNumber());
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		return result;
	}

	public void convertId2Ma(XmlNode xmlNode, string path)
	{
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_007a: Expected O, but got Unknown
		try
		{
			if (!_isDmMoi || xmlNode == null)
			{
				return;
			}
			XmlNodeList val = null;
			val = xmlNode.SelectNodes(path);
			if (val == null || val.Count <= 0 || _dmKSData == null || _dmKSData.Count <= 0)
			{
				return;
			}
			foreach (XmlNode item2 in val)
			{
				XmlNode val2 = item2;
				string idTN = val2.InnerText;
				if (!HTKKFormater.isNullString(idTN))
				{
					item item = _dmKSData.Find((item obj) => obj.id.Equals(idTN));
					if (item != null)
					{
						val2.InnerText = item.ma;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void convertMa2Id(XmlNode xmlNode, string path, HTKKPeriodInfo PeriodInfo)
	{
		//IL_0077: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Expected O, but got Unknown
		try
		{
			if (!_isDmMoi || xmlNode == null)
			{
				return;
			}
			XmlNodeList val = null;
			val = xmlNode.SelectNodes(path);
			if (val == null || val.Count <= 0 || _dmKSData == null || _dmKSData.Count <= 0)
			{
				return;
			}
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				string maTN = val2.InnerText;
				string[] array = new string[2];
				if (!HTKKFormater.isNullString(maTN))
				{
					List<item> list = _dmKSData.FindAll((item obj) => obj.ma.Equals(maTN));
					if (list != null && list.Count > 0 && !HTKKFormater.isNullString(list[0].id))
					{
						array[0] = list[0].id;
					}
				}
				if (!HTKKFormater.isNullString(array[1]))
				{
					val2.InnerText = array[1];
				}
				else
				{
					val2.InnerText = array[0];
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void StartImportXML_0(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			if (PeriodInfo.dicVar != null && PeriodInfo.dicVar.Count > 0 && PeriodInfo.dicVar.ContainsKey("IS_DM_MOI"))
			{
				convertMa2Id(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/TaiNguyenKhaiThac/ThueTaiNguyenKhaiThac/ct1", PeriodInfo);
				convertMa2Id(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/TaiNguyenThuMuaGom/ThueTaiNguyenThuMuaGom/ct1", PeriodInfo);
				convertMa2Id(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/TaiNguyenTichThuGiaoBan/ThueTaiNguyenTichThuGiaoBan/ct1", PeriodInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void validateTaiNguyen(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			int iCol2 = HTKKFile.ColumnLettersToIndex("F");
			int num = HTKKFile.ColumnLettersToIndex("V");
			int iCol3 = HTKKFile.ColumnLettersToIndex("AA");
			int iCol4 = HTKKFile.ColumnLettersToIndex("AD");
			if (iCol != num)
			{
				return;
			}
			string valColMa = GeneralFunctions.GetValueCustom(sView, (int)iRow, iCol2, "TextCellType");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, (int)iRow, num, "NumberCellType");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sView, (int)iRow, iCol3, "NumberCellType");
			string valueCustom3 = GeneralFunctions.GetValueCustom(sView, (int)iRow, iCol4, "NumberCellType");
			HTKKFile.RemoveValidate(sView, (int)iRow - 1, num - 1);
			if (HTKKFormater.isNullString(valColMa) || _dmKSData == null || _dmKSData.Count <= 0 || !(valueCustom3.ToNumber() == 0m) || !(valueCustom.ToNumber() != 0m))
			{
				return;
			}
			bool flag = true;
			item item = _dmKSData.Find((item x) => x.id.Contains(valColMa));
			if (item == null)
			{
				return;
			}
			if (item.cqtgia.ToNumber() > 0m)
			{
				flag = false;
				if (valueCustom.ToNumber() < item.cqtgia.ToNumber())
				{
					HTKKFile.SetNodeBackGroundPink(sView, (int)iRow - 1, num - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_25")));
				}
			}
			if (flag && valueCustom.ToNumber() < item.giatoithieu.ToNumber())
			{
				HTKKFile.SetNodeBackGroundPink(sView, (int)iRow - 1, num - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_24")));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void validateThueSuat(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			int iCol2 = HTKKFile.ColumnLettersToIndex("F");
			int num = HTKKFile.ColumnLettersToIndex("V");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int iCol3 = HTKKFile.ColumnLettersToIndex("AA");
			int iCol4 = HTKKFile.ColumnLettersToIndex("AD");
			if (iCol != num2)
			{
				return;
			}
			string valColMa = GeneralFunctions.GetValueCustom(sView, (int)iRow, iCol2, "TextCellType");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, (int)iRow, num, "NumberCellType");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sView, (int)iRow, iCol3, "NumberCellType");
			string valueCustom3 = GeneralFunctions.GetValueCustom(sView, (int)iRow, iCol4, "NumberCellType");
			HTKKFile.RemoveValidate(sView, (int)iRow - 1, num - 1);
			if (HTKKFormater.isNullString(valColMa) || _dmKSData == null || _dmKSData.Count <= 0 || !(valueCustom3.ToNumber() == 0m) || !(valueCustom.ToNumber() != 0m))
			{
				return;
			}
			bool flag = true;
			item item = _dmKSData.Find((item x) => x.id.Contains(valColMa));
			if (item == null)
			{
				return;
			}
			if (item.cqtgia.ToNumber() > 0m)
			{
				flag = false;
				if (valueCustom.ToNumber() < item.cqtgia.ToNumber())
				{
					HTKKFile.SetNodeBackGroundPink(sView, (int)iRow - 1, num - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_25")));
				}
			}
			if (flag && valueCustom.ToNumber() < item.giatoithieu.ToNumber())
			{
				HTKKFile.SetNodeBackGroundPink(sView, (int)iRow - 1, num - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_24")));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
