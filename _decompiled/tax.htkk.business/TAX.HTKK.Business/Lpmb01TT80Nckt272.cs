using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Lpmb01TT80Nckt272
{
	private string cellId_CT22_ct5 = "Y_47";

	private string cellId_CT22_ct6 = "AA_47";

	private Dictionary<string, string[]> dicKSTmpMB22 = new Dictionary<string, string[]>();

	private Dictionary<string, string> resultDicMB22 = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicKSTmpMB23 = new Dictionary<string, string[]>();

	private Dictionary<string, string> resultDicMB23 = new Dictionary<string, string>();

	private int[] cellTenTinh = GeneralFunctions.ArrayCellId("J_47");

	private int[] cellMaTinh = GeneralFunctions.ArrayCellId("AL_47");

	private int[] cellTenHuyen = GeneralFunctions.ArrayCellId("O_47");

	private int[] cellMaHuyen = GeneralFunctions.ArrayCellId("AM_47");

	private int[] cellKhaiNuaNam = GeneralFunctions.ArrayCellId("J_30");

	private int[] cellCt09 = GeneralFunctions.ArrayCellId("R_28");

	private int[] cellCt2_Bang1 = GeneralFunctions.ArrayCellId("C_47");

	private int[] cellCt4_Bang1 = GeneralFunctions.ArrayCellId("T_47");

	private int[] cellCt5_Bang1 = GeneralFunctions.ArrayCellId("Y_47");

	private int[] cellCt6_Bang1 = GeneralFunctions.ArrayCellId("AA_47");

	private int[] cellCt9_Bang1 = GeneralFunctions.ArrayCellId("AD_47");

	private int[] cellKhoaBang1 = GeneralFunctions.ArrayCellId("Z_33");

	private int[] CellId_CT22_ct5 => GeneralFunctions.ArrayCellId(cellId_CT22_ct5);

	private int[] CellId_CT22_ct6 => GeneralFunctions.ArrayCellId(cellId_CT22_ct6);

	private int idxCT23_ColT => HTKKFile.ColumnLettersToIndex("T") - 1;

	private int col5 => HTKKFile.ColumnLettersToIndex("Y") - 1;

	private int idxCT23_ColAA => HTKKFile.ColumnLettersToIndex("AA") - 1;

	public Dictionary<string, string> GetDMMucLePhiMonBaiCT22(string pathFile)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		//IL_0077: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Expected O, but got Unknown
		try
		{
			if (resultDicMB22 != null && resultDicMB22.Count > 0)
			{
				return resultDicMB22;
			}
			XmlDocument val = new XmlDocument();
			if (!File.Exists(pathFile))
			{
				return null;
			}
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 0 && int.Parse(array[0]) >= 500000)
				{
					string key = array[0];
					string value2 = array[1];
					if (!resultDicMB22.ContainsKey(key))
					{
						resultDicMB22.Add(key, value2);
					}
				}
			}
			return resultDicMB22;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMMucLePhiMonBaiCT23(string pathFile)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_007a: Expected O, but got Unknown
		try
		{
			if (resultDicMB23 != null && resultDicMB22.Count > 0)
			{
				return resultDicMB22;
			}
			XmlDocument val = new XmlDocument();
			if (!File.Exists(pathFile))
			{
				return null;
			}
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 0)
				{
					string key = array[0];
					string value2 = array[1];
					if (!resultDicMB23.ContainsKey(key))
					{
						resultDicMB23.Add(key, value2);
					}
				}
			}
			return resultDicMB23;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			ThayDoiTinhTheoCQT(winFpSpread.Sheets[0], TTinChung, ToKhaiInfo, PeriodInfo);
			LockTheoCQT(winFpSpread.Sheets[0], TTinChung, ToKhaiInfo);
			bool flag = false;
			string text = "";
			if (PeriodInfo.dicVar.Count > 0)
			{
				flag = Convert.ToBoolean(PeriodInfo.dicVar["CoSoMoiThanhLap"]);
				text = (string)PeriodInfo.dicVar["NgayThanhLap"];
				int[] array = GeneralFunctions.ArrayCellId("R_28");
				int[] array2 = GeneralFunctions.ArrayCellId("AA_28");
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Value = flag;
				winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Text = text;
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
				ThayDoiTinhTheoCQT(winFpSpread.Sheets[0], TTinChung, ToKhaiInfo, PeriodInfo);
				LockTheoCQT(winFpSpread.Sheets[0], TTinChung, ToKhaiInfo);
			}
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
			int num = HTKKFile.ColumnLettersToIndex("AD");
			int num2 = HTKKFile.ColumnLettersToIndex("AN");
			int num3 = HTKKFile.ColumnLettersToIndex("AA");
			int num4 = HTKKFile.ColumnLettersToIndex("AO");
			int num5 = HTKKFile.ColumnLettersToIndex("AC");
			int num6 = HTKKFile.ColumnLettersToIndex("C");
			int num7 = HTKKFile.ColumnLettersToIndex("J");
			int num8 = HTKKFile.ColumnLettersToIndex("AL");
			List<int> list = new List<int>();
			list.Add(num5);
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			if (activeSheetIndex != 0)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			if (string.IsNullOrEmpty(sheetView.Cells[(int)iRow - 1, num - 1].Text))
			{
				sheetView.Cells[(int)iRow - 1, num2 - 1].Text = "";
			}
			if (iRow == cellCt09[1] + 1)
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
				ThayDoiTinhTheoCQT(winFpSpread.Sheets[0], TTinChung, ToKhaiInfo, PeriodInfo);
				LockTheoCQT(winFpSpread.Sheets[0], TTinChung, ToKhaiInfo);
			}
			if (iCol == num || iCol == col5 + 1 || iCol == num3)
			{
				string value = sheetView.Cells[(int)iRow - 1, num2 - 1].Text.Trim();
				if (!string.IsNullOrEmpty(value))
				{
					GeneralFunctions.LockUnLockCols(sheetView, (int)iRow, list, locked: true, deleteData: false);
					sheetView.Cells[(int)iRow - 1, num5 - 1].Value = sheetView.Cells[(int)iRow - 1, num4 - 1].Value;
				}
				else
				{
					GeneralFunctions.LockUnLockCols(sheetView, (int)iRow, list, locked: true, deleteData: false);
					sheetView.Cells[(int)iRow - 1, num5 - 1].Value = 0;
				}
			}
			if (iCol != num6)
			{
				return;
			}
			string text = sheetView.Cells[(int)iRow - 1, num6 - 1].Text;
			if (string.IsNullOrEmpty(text))
			{
				return;
			}
			Dictionary<string, string> dicTinhTP = HTKKWinDictionaryDiaDanh.dicTinhTP;
			foreach (KeyValuePair<string, string> item in dicTinhTP)
			{
				if (item.Value.Equals(PeriodInfo.maCQTCapCuc.Substring(0, 3)))
				{
					sheetView.Cells[(int)iRow - 1, num7 - 1].Text = item.Key;
					sheetView.Cells[(int)iRow - 1, num8 - 1].Text = PeriodInfo.maCQTCapCuc.Substring(0, 3);
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			int[] array = GeneralFunctions.ArrayCellId("R_28");
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("AA");
			int num3 = HTKKFile.ColumnLettersToIndex("AO");
			int[] array2 = GeneralFunctions.ArrayCellId("J_47");
			int[] array3 = GeneralFunctions.ArrayCellId("O_47");
			int num4 = HTKKFile.ColumnLettersToIndex("J");
			int num5 = HTKKFile.ColumnLettersToIndex("AL");
			int num6 = HTKKFile.ColumnLettersToIndex("O");
			int num7 = HTKKFile.ColumnLettersToIndex("AM");
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28"));
			double num8 = ConvertHelpers.ConvertToDouble1(sheetView.Cells[CellId_CT22_ct5[1], CellId_CT22_ct5[0]].Value);
			double num9 = ConvertHelpers.ConvertToDouble1(sheetView.Cells[CellId_CT22_ct6[1], CellId_CT22_ct6[0]].Value);
			bool flag = ConvertHelpers.ConvertToBool(sheetView.Cells[array[1], array[0]].Value);
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num10 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			if (string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc))
			{
				return;
			}
			string maTinh = (HTKKWinDictionaryDiaDanh.dicTinhTP.ContainsValue(PeriodInfo.maCQTCapCuc.Substring(0, 3)) ? PeriodInfo.maCQTCapCuc.Substring(0, 3) : "");
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sheetView.Cells[num10 + i, idxCT23_ColT].Text;
				string text2 = sheetView.Cells[num10 + i, num - 1].Text;
				double num11 = Convert.ToDouble(sheetView.Cells[num10 + i, col5].Value);
				double num12 = Convert.ToDouble(sheetView.Cells[num10 + i, num3 - 1].Value);
				if (!string.IsNullOrEmpty(text2) && string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num10 + i, idxCT23_ColT, message);
					value = 2;
				}
				else if (!string.IsNullOrEmpty(text) && text.Length != 5 && !ValidateBase.CheckMaSoThue(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num10 + i, idxCT23_ColT, message3);
					value = 2;
				}
				else if (!string.IsNullOrEmpty(text) && text.Length == 5 && !ValidateBase.CheckMaSoThue5So(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num10 + i, idxCT23_ColT, message3);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num10 + i, idxCT23_ColT);
				}
				if (!(PeriodInfo.maCQTCapCuc != "82500"))
				{
					continue;
				}
				if (!string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy) && !string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc) && PeriodInfo.maCQTQuanLy != PeriodInfo.maCQTCapCuc)
				{
					string text3 = (HTKKWinDictionaryDiaDanh.dicQuanHuyen.ContainsKey(PeriodInfo.maCQTQuanLy) ? PeriodInfo.maCQTQuanLy : "");
					string text4 = sheetView.Cells[num10 + i, num5 - 1].Text.Trim();
					string text5 = sheetView.Cells[num10 + i, num7 - 1].Text.Trim();
					if (!string.IsNullOrEmpty(maTinh) && !string.IsNullOrEmpty(text4))
					{
						KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
						if (!string.IsNullOrEmpty(text4) && text4 != maTinh)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num10 + i, num4 - 1, message4);
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num10 + i, num4 - 1);
						}
					}
				}
				else
				{
					if (string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc))
					{
						continue;
					}
					string text4 = sheetView.Cells[num10 + i, num5 - 1].Text.Trim();
					if (!string.IsNullOrEmpty(maTinh) && !string.IsNullOrEmpty(text4))
					{
						KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
						if (!string.IsNullOrEmpty(text4) && text4 != maTinh)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num10 + i, num4 - 1, message6);
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num10 + i, num4 - 1);
						}
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
		}
	}

	private void XuLyCT22_ct5(SheetView sViewTK, int iRow)
	{
		try
		{
			double num = Convert.ToDouble(sViewTK.Cells[iRow, CellId_CT22_ct5[0]].Value);
			double num2 = 10000000000.0;
			if (iRow > 46)
			{
				if (ConvertHelpers.ConvertToDouble1(sViewTK.Cells[iRow, CellId_CT22_ct6[0]].Value) == 0.0)
				{
					sViewTK.Cells[iRow, CellId_CT22_ct6[0]].Value = 1000000;
				}
			}
			else if (num > num2)
			{
				sViewTK.Cells[iRow, CellId_CT22_ct6[0]].Value = 3000000;
			}
			else
			{
				sViewTK.Cells[iRow, CellId_CT22_ct6[0]].Value = 2000000;
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
			int[] array = GeneralFunctions.ArrayCellId("AA_47");
			int[] array2 = GeneralFunctions.ArrayCellId("AO_47");
			int num = HTKKFile.ColumnLettersToIndex("AA");
			int num2 = HTKKFile.ColumnLettersToIndex("AO");
			SheetView sheetView = winFpSpread.Sheets[0];
			if (ToKhaiInfo.SheetInfos[0].Tables.Count > 0)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					string key = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Value.ToString().Trim();
					string value = (resultDicMB23.ContainsKey(key) ? resultDicMB23[key] : "");
					sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Value = value;
				}
			}
			string key2 = sheetView.Cells[array2[1], array2[0]].Value.ToString().Trim();
			string value2 = (resultDicMB22.ContainsKey(key2) ? resultDicMB22[key2] : "");
			sheetView.Cells[array[1], array[0]].Value = value2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void LockTheoCQT(SheetView sView, HTKKCMCTTinChungInfo TTinChung, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
		bool flag = Convert.ToBoolean(sView.Cells[cellCt09[1], cellCt09[0]].Value);
		if ((!string.IsNullOrEmpty(TTinChung.maCQTQuanLy) && headerInfo.maCQTQuanLy != TTinChung.maCQTQuanLy) || (headerInfo.maCQTQuanLy == TTinChung.maCQTQuanLy && flag))
		{
			foreach (HTKKCMCCellInfo staticCell in ToKhaiInfo.SheetInfos[0].StaticCells)
			{
				if (staticCell.CellTemplate.RowIndex == cellTenTinh[1] + 1)
				{
					sView.Cells[cellTenTinh[1], staticCell.CellTemplate.ColIndex - 1].Value = staticCell.DefaultValue;
					GeneralFunctions.LockUnLockCol(sView, cellTenTinh[1] + 1, staticCell.CellTemplate.ColIndex, locked: true, deleteData: true);
				}
			}
			sView.Cells[cellKhoaBang1[1], cellKhoaBang1[0]].Value = 1;
			return;
		}
		foreach (HTKKCMCCellInfo staticCell2 in ToKhaiInfo.SheetInfos[0].StaticCells)
		{
			if (staticCell2.CellTemplate.RowIndex == cellTenTinh[1] + 1 && (staticCell2.CellTemplate.ColIndex == cellCt5_Bang1[0] + 1 || staticCell2.CellTemplate.ColIndex == cellCt6_Bang1[0] + 1 || staticCell2.CellTemplate.ColIndex == cellCt9_Bang1[0] + 1))
			{
				GeneralFunctions.LockUnLockCol(sView, cellTenTinh[1] + 1, staticCell2.CellTemplate.ColIndex, locked: false, deleteData: false);
			}
		}
		sView.Cells[cellKhoaBang1[1], cellKhoaBang1[0]].Value = 0;
	}

	public bool UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			ThayDoiTinhTheoCQT(winFpSpread.Sheets[0], TTinChung, ToKhaiInfo, PeriodInfo);
			LockTheoCQT(winFpSpread.Sheets[0], TTinChung, ToKhaiInfo);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public void ThayDoiTinhTheoCQT(SheetView sView, HTKKCMCTTinChungInfo TTinChung, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo)
	{
		int num = HTKKFile.ColumnLettersToIndex("J");
		int num2 = HTKKFile.ColumnLettersToIndex("AL");
		int num3 = HTKKFile.ColumnLettersToIndex("O");
		int num4 = HTKKFile.ColumnLettersToIndex("AM");
		HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
		if (PeriodInfo.maCQTCapCuc == "82500")
		{
			GeneralFunctions.LockUnLockCol(sView, cellTenTinh[1] + 1, cellTenTinh[0] + 1, locked: false, deleteData: false);
			GeneralFunctions.LockUnLockCol(sView, cellTenHuyen[1] + 1, cellTenHuyen[0] + 1, locked: false, deleteData: false);
		}
		else
		{
			if (string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc))
			{
				return;
			}
			string maTinh = (HTKKWinDictionaryDiaDanh.dicTinhTP.ContainsValue(PeriodInfo.maCQTCapCuc.Substring(0, 3)) ? PeriodInfo.maCQTCapCuc.Substring(0, 3) : "");
			if (!string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy) && !string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc) && PeriodInfo.maCQTQuanLy != PeriodInfo.maCQTCapCuc)
			{
				string text = (HTKKWinDictionaryDiaDanh.dicQuanHuyen.ContainsKey(PeriodInfo.maCQTQuanLy) ? PeriodInfo.maCQTQuanLy : "");
				if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(maTinh))
				{
					KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
					sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text = keyValuePair.Key;
					sView.Cells[cellMaTinh[1], cellMaTinh[0]].Text = maTinh;
					sView.Cells[cellTenHuyen[1], cellTenHuyen[0]].Text = HTKKWinDictionaryDiaDanh.dicQuanHuyen[text];
					sView.Cells[cellMaHuyen[1], cellMaHuyen[0]].Text = text;
					GeneralFunctions.LockUnLockCol(sView, cellTenTinh[1] + 1, cellTenTinh[0] + 1, locked: true, deleteData: false);
					GeneralFunctions.LockUnLockCol(sView, cellTenHuyen[1] + 1, cellTenHuyen[0] + 1, locked: true, deleteData: false);
				}
			}
			else if (!string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc))
			{
				KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
				sView.Cells[cellTenTinh[1], cellTenTinh[0]].Text = keyValuePair.Key;
				sView.Cells[cellMaTinh[1], cellMaTinh[0]].Text = maTinh;
				GeneralFunctions.LockUnLockCol(sView, cellTenTinh[1] + 1, cellTenTinh[0] + 1, locked: true, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, cellTenHuyen[1] + 1, cellTenHuyen[0] + 1, locked: false, deleteData: false);
			}
		}
	}
}
