using System;
using System.Collections.Generic;
using System.Drawing;
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

public class Bvmt02Nckt205
{
	private Dictionary<string, string[]> dicDMKhoangSan = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMKhoangSanTmp = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private string namKK = "";

	private List<item> _dmKSData = new List<item>();

	private bool _isDmMoi = false;

	private int idxCol_C => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_H => HTKKFile.ColumnLettersToIndex("H") - 1;

	private int idxCol_J => HTKKFile.ColumnLettersToIndex("J") - 1;

	private int idxCol_N => HTKKFile.ColumnLettersToIndex("N") - 1;

	private int idxCol_Q => HTKKFile.ColumnLettersToIndex("Q") - 1;

	private int idxCol_V => HTKKFile.ColumnLettersToIndex("V") - 1;

	private int idxCol_Y => HTKKFile.ColumnLettersToIndex("Y") - 1;

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			namKK = PeriodInfo.Nam;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			string text = "ROUND(J{0}*N{0},0)";
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[1];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			int num2 = dataLocationTemplateRowStart2 + hTKKCMCTableInfo2.RowsValues.Count - 1;
			if ((iRow >= dataLocationTemplateRowStart && iRow <= num) || (iRow >= dataLocationTemplateRowStart2 && iRow <= num2))
			{
				if (iCol - 1 == idxCol_J || iCol - 1 == idxCol_N)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_Q].Formula = string.Format(string.Format("ROUND(J{0}*N{0},0)", iRow), iRow.ToString());
					sheetView.Cells[(int)iRow - 1, idxCol_Y].Formula = string.Format(string.Format("Q{0}-V{0}", iRow), iRow.ToString());
				}
				else if (iCol - 1 == idxCol_Q || iCol - 1 == idxCol_V)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_Y].Formula = string.Format(string.Format("Q{0}-V{0}", iRow), iRow.ToString());
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
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15"));
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					array[0] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_C].Value);
					array[1] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_H].Value);
					array[2] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_J].Value);
					array[3] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_N].Value);
					array[4] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_Q].Value);
					array[5] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_V].Value);
					array[6] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_Y].Value);
					if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array))
					{
						HTKKFile.RemoveValidate(sheetView, num + j, idxCol_C);
					}
					else if (string.IsNullOrEmpty(array[0]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxCol_C, message);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + j, idxCol_C);
					}
				}
			}
			switch (value)
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
			winFpSpread.TabStrip[0].BackColor = Color.Pink;
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
		//IL_02f4: Unknown result type (might be due to invalid IL or missing references)
		//IL_02fb: Expected O, but got Unknown
		//IL_0329: Unknown result type (might be due to invalid IL or missing references)
		//IL_0330: Expected O, but got Unknown
		string text = "###0103###020101###020302###020303###0301###0302###0303###0304###0401###0402###0403###0404###0405###0702###07030101###07030102###07030201###07030202###090201###090202###090204###090205###";
		try
		{
			_isDmMoi = true;
			if (dicDMKhoangSanTmp.Count > 0)
			{
				return dicDMKhoangSanTmp;
			}
			List<item> list = new List<item>();
			_dmKSData = new List<item>();
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
					if (!namKK.Equals("2019"))
					{
						list.Add(item3);
						_dmKSData.Add(item3);
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
					string text2 = item5.id + "_" + item5.ma + "_" + item5.thuesuat + "_" + item5.giatoithieu + "_" + item5.cqtgia;
					string[] array = new string[7] { text2, item5.ma, ten, item5.dvt, item5.thuesuat, "0", "1" };
					item5.id = text2;
					if (!dictionary.ContainsKey(text2))
					{
						string value2 = array[2];
						dictionary2.Add(text2, value2);
						dictionary.Add(text2, array);
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
				convertMa2Id(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/KSCSTuKhaiThac/CTietKSCSTuKhaiThac/ct1", PeriodInfo);
				convertMa2Id(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/KSCSThuMua/CTietKSCSThuMua/ct1", PeriodInfo);
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
				convertId2Ma(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/KSCSTuKhaiThac/CTietKSCSTuKhaiThac/ct1");
				convertId2Ma(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/KSCSThuMua/CTietKSCSThuMua/ct1");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
