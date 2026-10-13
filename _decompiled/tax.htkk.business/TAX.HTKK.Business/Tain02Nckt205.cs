using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using HTKKDanhmuc;
using HTKKSerializer;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tain02Nckt205
{
	private Dictionary<string, string[]> dicDMKhoangSan = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMKhoangSanTmp = new Dictionary<string, string>();

	private List<item> _dmKSData = new List<item>();

	private bool _isDmMoi = false;

	private List<string> pathDonGiaTinh = new List<string>();

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			if (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) >= 2014)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("AH");
			bool flag = true;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					sView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Locked = true;
				}
			}
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
			int num7 = HTKKFile.ColumnLettersToIndex("AH");
			if (iCol != num2)
			{
				return;
			}
			winFpSpread.TabStrip[0].BackColor = Color.White;
			HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, num2 - 1);
			foreach (KeyValuePair<string, string[]> item in dicDMKhoangSan)
			{
				string[] value = item.Value;
				if (value.Length <= 3)
				{
					continue;
				}
				if (value[2].Trim() == sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim())
				{
					if (ConvertHelpers.ConvertToInt(value[4]) == 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, num2 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
						winFpSpread.TabStrip[0].BackColor = Color.Pink;
					}
					string text = item.Key;
					if (item.Key.Split(new char[1] { '_' }).Length > 3)
					{
						text = item.Key.Split(new char[1] { '_' })[2];
					}
					sheetView.Cells[(int)iRow - 1, num - 1].Text = text;
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = value[3];
					sheetView.Cells[(int)iRow - 1, num5 - 1].Text = value[4];
					double donGia = getDonGia(item.Key);
					sheetView.Cells[(int)iRow - 1, num4 - 1].Value = donGia;
					sheetView.Cells[(int)iRow - 1, num3 - 1].Locked = false;
					break;
				}
				if (sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim().Equals("Tài nguyên khác"))
				{
					sheetView.Cells[(int)iRow - 1, num5 - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num5 - 1].Locked = true;
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
			int num7 = HTKKFile.ColumnLettersToIndex("S");
			int num8 = HTKKFile.ColumnLettersToIndex("AH");
			string text = "IF(AND(V###>0,AD###=0),ROUND(S###*V###*AA###/100,0),ROUND(S###*AD###,0))";
			if (iCol == num7 || iCol == num4 || iCol == num5 || iCol == num6)
			{
				sheetView.Cells[(int)iRow - 1, num8 - 1].Formula = text.Replace("###", iRow.ToString());
			}
			if (iCol == num4)
			{
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, (int)iCol, "NumberCellType");
				if (ConvertHelpers.ConvertToDouble(valueCustom) > 0.0)
				{
					sheetView.Cells[(int)iRow - 1, num6 - 1].Text = "0";
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
		int value = 0;
		levelWarning.TryGetValue(0, out value);
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int iCol = HTKKFile.ColumnLettersToIndex("F");
			int num = HTKKFile.ColumnLettersToIndex("G");
			int num2 = HTKKFile.ColumnLettersToIndex("P");
			int num3 = HTKKFile.ColumnLettersToIndex("S");
			int num4 = HTKKFile.ColumnLettersToIndex("V");
			int iCol2 = HTKKFile.ColumnLettersToIndex("AA");
			int iCol3 = HTKKFile.ColumnLettersToIndex("AD");
			int iCol4 = HTKKFile.ColumnLettersToIndex("AH");
			int num5 = HTKKFile.ColumnLettersToIndex("AK");
			int iCol5 = HTKKFile.ColumnLettersToIndex("AP");
			Dictionary<int, int[]> dictionary = new Dictionary<int, int[]>();
			List<HTKKCMCTableInfo> tables = ToKhaiInfo.SheetInfos[0].Tables;
			for (int i = 0; i < tables.Count; i++)
			{
				int dataLocationTemplateRowStart = tables[i].Location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = tables[i].RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string valColMa = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol, "TextCellType");
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num3, "NumberCellType");
					string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num4, "NumberCellType");
					string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol2, "NumberCellType");
					string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol3, "NumberCellType");
					string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num5, "NumberCellType");
					string valueCustom6 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol5, "NumberCellType");
					string valueCustom7 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol4, "NumberCellType");
					string text = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text;
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num - 1);
					item item;
					if (!HTKKFormater.isNullString(valColMa) && _dmKSData != null && _dmKSData.Count > 0)
					{
						item = _dmKSData.Find((item x) => x.id.Contains(valColMa));
						if (item != null && (HTKKFormater.isNullString(item.thuesuat) || ConvertHelpers.ConvertToInt(item.thuesuat) == 0))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
							value = 2;
						}
					}
					if (string.IsNullOrEmpty(text) && (ConvertHelpers.ConvertToDouble(valueCustom) > 0.0 || ConvertHelpers.ConvertToDouble(valueCustom2) > 0.0 || ConvertHelpers.ConvertToDouble(valueCustom3) > 0.0 || ConvertHelpers.ConvertToDouble(valueCustom4) > 0.0 || ConvertHelpers.ConvertToDouble(valueCustom5) > 0.0 || ConvertHelpers.ConvertToDouble(valueCustom6) > 0.0))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
						value = 2;
					}
					if (ConvertHelpers.ConvertToDouble(valueCustom5) > ConvertHelpers.ConvertToDouble(valueCustom7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1);
					}
					if (!string.IsNullOrEmpty(text) && ConvertHelpers.ConvertToDouble(valueCustom) == 0.0)
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
						if (value == 0)
						{
							value = 1;
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
					}
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1);
					if (HTKKFormater.isNullString(valColMa) || _dmKSData == null || _dmKSData.Count <= 0 || PeriodInfo.dicVar == null || PeriodInfo.dicVar.Count <= 0 || !PeriodInfo.dicVar.ContainsKey("IS_DM_MOI") || !(valueCustom4.ToNumber() == 0m) || !(valueCustom2.ToNumber() != 0m))
					{
						continue;
					}
					bool flag = true;
					item = _dmKSData.Find((item x) => x.id.Contains(valColMa));
					if (item == null)
					{
						continue;
					}
					if (item.cqtgia.ToNumber() > 0m)
					{
						flag = false;
						if (valueCustom2.ToNumber() < item.cqtgia.ToNumber())
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
							value = 2;
						}
					}
					if (flag && valueCustom2.ToNumber() < item.giatoithieu.ToNumber())
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22")));
						value = 2;
					}
				}
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
		//IL_0396: Unknown result type (might be due to invalid IL or missing references)
		//IL_039d: Expected O, but got Unknown
		//IL_03cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_03d2: Expected O, but got Unknown
		List<item> list = null;
		try
		{
			_isDmMoi = true;
			if (dicDMKhoangSanTmp.Count > 0)
			{
				return dicDMKhoangSanTmp;
			}
			_dmKSData = new List<item>();
			List<item> list2 = new List<item>();
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
							if (!File.Exists(p))
							{
								string str = pathDonGiaTinh.Find((string o) => o.Equals(p));
								if (HTKKFormater.isNullString(str))
								{
									pathDonGiaTinh.Add(p);
								}
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
							list2.Add(item);
						}
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1)
						{
							list2.Add(item);
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
						list2.Add(item3);
						_dmKSData.Add(item3);
					}
				}
			}
			if (list2 != null && list2.Count > 0)
			{
				List<item> dmTaiNguyenForTinh = GeneralFunctions.getDmTaiNguyenForTinh(pathDonGiaTinh, PeriodInfo);
				List<item> list3 = Serializer.DeepCopy(list2);
				foreach (item iT in list2)
				{
					Predicate<item> match = (item obj) => obj.ma.Equals(iT.ma);
					List<item> list4 = dmTaiNguyenForTinh.FindAll(match);
					DateTime dateTime = DateTime.ParseExact(iT.hltu, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					DateTime dateTime2 = DateTime.ParseExact(iT.hlden, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					List<item> list5 = new List<item>();
					if (list4 != null && list4.Count > 0)
					{
						for (int num2 = 0; num2 < list4.Count; num2++)
						{
							DateTime dateTime3 = DateTime.ParseExact(list4[num2].cqthltu, "dd/MM/yyyy", CultureInfo.CurrentCulture);
							DateTime dateTime4 = DateTime.ParseExact(list4[num2].cqthlden, "dd/MM/yyyy", CultureInfo.CurrentCulture);
							if (dateTime2 <= dateTime4)
							{
								list5.Add(list4[num2]);
							}
						}
					}
					if (list5 != null && list5.Count > 0)
					{
						item item4 = list3.Find((item obj) => obj.id == iT.id);
						int num3 = list3.FindIndex((item obj) => obj.id == iT.id);
						for (int num4 = 0; num4 < list5.Count; num4++)
						{
							item item5 = Serializer.DeepCopy(iT);
							item5.cqtgia = list5[num4].cqtgia;
							item5.isCqt = true;
							list3.Insert(num3 + num4 + 1, item5);
						}
						list3.Remove(item4);
					}
				}
				list2 = Serializer.DeepCopy(list3);
				list3 = new List<item>();
				foreach (item item8 in list2)
				{
					if (item8.ten.IndexOf("Đá vôi nguyên khai") > -1)
					{
						int num5 = 0;
					}
					item item6 = Serializer.DeepCopy(item8);
					string ten = item6.ten;
					List<item> list6 = list2.FindAll((item obj) => obj.ten.Equals(ten));
					if (list6 != null && list6.Count > 1)
					{
						item6 = GeneralFunctions.getTaiNguyenForCombobox(item6, list6);
						if (item6 == null && item8.id == list6[0].id)
						{
							list3.Add(Serializer.DeepCopy(item8));
						}
						else if (item6 == null && (item8.id.Contains("V10202") || item8.id.Contains("V10103") || item8.id.Contains("II120201") || item8.id.Contains("X1101")))
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
				for (int num6 = 0; num6 < list3.Count; num6++)
				{
					item item6 = Serializer.DeepCopy(list3[num6]);
					string text = item6.id + "_" + item6.ma + "_" + item6.thuesuat + "_" + item6.giatoithieu + "_" + item6.cqtgia;
					string text2 = item6.ten;
					string text3 = "";
					if (pathDms.Count > 1)
					{
						text3 = "1";
					}
					else if (pathDms.Count == 1 && pathDms[0].pathDonGiaTinh != null && pathDms[0].pathDonGiaTinh.Count > 0)
					{
						text3 = "1";
					}
					if (text3.Equals("1"))
					{
						if (!"".Equals(item6.thuesuat))
						{
							text2 = text2 + " (" + item6.thuesuat + ")";
						}
						if (!"".Equals(item6.giatoithieu))
						{
							text2 = text2 + " (" + item6.giatoithieu + ")";
						}
						if (!"".Equals(item6.cqtgia) && item6.cqtgia != null)
						{
							text2 = text2 + " (" + item6.cqtgia + ")";
						}
						if (item6.cqtgia != null && ConvertHelpers.ConvertToDouble(item6.giatoithieu) > ConvertHelpers.ConvertToDouble(item6.cqtgia))
						{
							continue;
						}
					}
					string[] array = new string[7] { text, item6.ma, text2, item6.dvt, item6.thuesuat, "0", "1" };
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

	private bool kyKeKhaiTrongHL(item itemTinh, HTKKPeriodInfo PeriodInfo)
	{
		bool result = false;
		try
		{
			DateTime[] array = new DateTime[2];
			DateTime[] array2 = new DateTime[2];
			if (HTKKFormater.isNullString(itemTinh.cqthltu))
			{
				return false;
			}
			ref DateTime reference = ref array2[0];
			reference = DateTime.ParseExact(itemTinh.cqthltu, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			if (HTKKFormater.isNullString(itemTinh.cqthlden))
			{
				return false;
			}
			ref DateTime reference2 = ref array2[1];
			reference2 = DateTime.ParseExact(itemTinh.cqthlden, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			string nam = PeriodInfo.Nam;
			string tuNgay = PeriodInfo.TuNgay;
			DateTime dateTime;
			if (!HTKKFormater.isNullString(tuNgay))
			{
				dateTime = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			}
			else
			{
				tuNgay = PeriodInfo.TuThang;
				dateTime = (HTKKFormater.isNullString(tuNgay) ? DateTime.ParseExact("01/01" + nam, "dd/MM/yyyy", CultureInfo.CurrentCulture) : DateTime.ParseExact("01/" + tuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture));
			}
			array[0] = dateTime;
			tuNgay = PeriodInfo.DenNgay;
			if (!HTKKFormater.isNullString(tuNgay))
			{
				dateTime = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			}
			else
			{
				tuNgay = PeriodInfo.DenThang;
				if (!HTKKFormater.isNullString(tuNgay))
				{
					dateTime = DateTime.ParseExact("01/" + tuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					dateTime = new DateTime(dateTime.Year, dateTime.Month, 1).AddMonths(1).AddDays(-1.0);
				}
				else
				{
					dateTime = DateTime.ParseExact("01/12" + nam, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					dateTime = new DateTime(dateTime.Year, dateTime.Month, 1).AddMonths(1).AddDays(-1.0);
				}
			}
			array[1] = dateTime;
			if (array[0] >= array2[0] && array[1] <= array2[1])
			{
				return true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
		return result;
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
				_ = tableInfo.iColumnSTT;
				bool flag = 1 == 0;
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
				convertMa2Id(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/TNKhaiThac/CTietTNKhaiThac/ct1", PeriodInfo);
				convertMa2Id(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/TNThuMuaGom/CTietTNThuMuaGom/ct1", PeriodInfo);
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
		try
		{
			if (PeriodInfo.dicVar != null && PeriodInfo.dicVar.Count > 0 && PeriodInfo.dicVar.ContainsKey("IS_DM_MOI"))
			{
				convertId2Ma(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/TNKhaiThac/CTietTNKhaiThac/ct1");
				convertId2Ma(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/TNThuMuaGom/CTietTNThuMuaGom/ct1");
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
			if (HTKKFormater.isNullString(valColMa) || _dmKSData == null || _dmKSData.Count <= 0 || PeriodInfo.dicVar == null || PeriodInfo.dicVar.Count <= 0 || !PeriodInfo.dicVar.ContainsKey("IS_DM_MOI") || !(valueCustom3.ToNumber() == 0m) || !(valueCustom.ToNumber() != 0m))
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
					HTKKFile.SetNodeBackGroundPink(sView, (int)iRow - 1, num - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23")));
				}
			}
			if (flag && valueCustom.ToNumber() < item.giatoithieu.ToNumber())
			{
				HTKKFile.SetNodeBackGroundPink(sView, (int)iRow - 1, num - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_22")));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
