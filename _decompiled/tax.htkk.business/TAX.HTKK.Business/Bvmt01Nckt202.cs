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

public class Bvmt01Nckt202
{
	private Dictionary<string, string[]> dicDMKhoangSan = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMKhoangSanTmp = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private List<item> _dmKSData = new List<item>();

	private bool _isDmMoi = false;

	private string formulaCol6 = "ROUND((L{0}*Q{0}),0)";

	private string pathTongSo = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/PhiBVMT/tongPhiPNop";

	private string pathTongChu = "PathB55";

	private string pathMaCQT = "PL01_KHBS/NOIDUNGGIAITHICH/coQuanThue";

	private int idxCol_C => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_J => HTKKFile.ColumnLettersToIndex("J") - 1;

	private int idxCol_L => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxCol_Q => HTKKFile.ColumnLettersToIndex("Q") - 1;

	private int idxCol_U => HTKKFile.ColumnLettersToIndex("U") - 1;

	private int[] KHBS_CellId_BE33 => GeneralFunctions.ArrayCellId("BE_33");

	private int[] KHBS_CellId_BI33 => GeneralFunctions.ArrayCellId("BI_33");

	private int[] CellId_E15 => GeneralFunctions.ArrayCellId("E_15");

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[1];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			int num2 = dataLocationTemplateRowStart2 + hTKKCMCTableInfo2.RowsValues.Count - 1;
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTongChu));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTongSo));
			if ((iRow >= dataLocationTemplateRowStart && iRow <= num) || (iRow >= dataLocationTemplateRowStart2 && iRow <= num2))
			{
				if (iCol - 1 == idxCol_L || iCol - 1 == idxCol_Q)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_U].Formula = string.Format(formulaCol6, iRow.ToString());
					string number = ConvertHelpers.ConvertToString(sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = MoneyHelpers.MoneyToText(number);
				}
				else if (iCol - 1 == idxCol_U)
				{
					string number = ConvertHelpers.ConvertToString(sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = MoneyHelpers.MoneyToText(number);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[1];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			int num2 = dataLocationTemplateRowStart2 + hTKKCMCTableInfo2.RowsValues.Count - 1;
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTongChu));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTongSo));
			string number = ConvertHelpers.ConvertToString(sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
			sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = MoneyHelpers.MoneyToText(number);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung)
	{
		try
		{
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTongChu));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTongSo));
			string number = ConvertHelpers.ConvertToString(winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
			winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = MoneyHelpers.MoneyToText(number);
			SheetView sheetView = winFpSpread.Sheets[0];
			int index = winFpSpread.Sheets.Count - 2;
			SheetView sheetView2 = winFpSpread.Sheets[index];
			if (sheetView2 == null || !sheetView2.Visible)
			{
				return;
			}
			HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[index].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathMaCQT));
			Dictionary<string, string> coQuanThueCapCuc = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc;
			string value = hTKKCMCCellInfo3.Value;
			string empty = string.Empty;
			string empty2 = string.Empty;
			if (string.IsNullOrEmpty(value))
			{
				return;
			}
			foreach (string key in coQuanThueCapCuc.Keys)
			{
				empty = coQuanThueCapCuc[key];
				empty2 = key;
				if (empty.Substring(0, 3).Equals(value.Substring(0, 3)))
				{
					sheetView2.Cells[KHBS_CellId_BE33[1], KHBS_CellId_BE33[0]].Value = empty2;
					sheetView2.Cells[KHBS_CellId_BI33[1], KHBS_CellId_BI33[0]].Text = empty;
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
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[1];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			int num2 = dataLocationTemplateRowStart2 + hTKKCMCTableInfo2.RowsValues.Count - 1;
			if (((iRow >= dataLocationTemplateRowStart && iRow <= num) || (iRow >= dataLocationTemplateRowStart2 && iRow <= num2)) && iCol - 1 == idxCol_U)
			{
				sheetView.Cells[(int)iRow - 1, idxCol_U].ResetFormula();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
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
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			string message = HTKKFile.GetMessage(sheetView, CellId_E15);
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
					array[1] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_J].Value);
					array[2] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_L].Value);
					array[3] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_Q].Value);
					array[4] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_U].Value);
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
				convertMa2Id(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/PhiBVMT/KSCSoTuKhaiThac/TTinKhoangSan/KhoangSan/maKhoangSan", PeriodInfo);
				convertMa2Id(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/PhiBVMT/KSanCoSoThuMuaNopThay/TTinKhoangSan/KhoangSan/maKhoangSan", PeriodInfo);
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
				convertId2Ma(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/PhiBVMT/KSCSoTuKhaiThac/TTinKhoangSan/KhoangSan/maKhoangSan");
				convertId2Ma(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/PhiBVMT/KSanCoSoThuMuaNopThay/TTinKhoangSan/KhoangSan/maKhoangSan");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
