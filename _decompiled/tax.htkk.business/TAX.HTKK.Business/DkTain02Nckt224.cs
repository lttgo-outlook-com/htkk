using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkTain02Nckt224
{
	private string cellId_O30 = "O_30";

	private string cellId_P30 = "P_30";

	private string cellId_Q30 = "Q_30";

	private string cellId_ct16_donvi = "P_49";

	private string cellId_ct16_daily = "AA_49";

	private string cellId_ct26 = "AV_57";

	private string cellId_ct28 = "AV_59";

	private string path_ct29 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct29";

	private double preVal_ct29;

	private string cellId_ct31 = "AV_64";

	private string cellId_Tong8_PL2 = "BG_64";

	private string cellId_ct33 = "AV_67";

	private string path_ct36 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct36";

	private double? preVal_ct36;

	private string cellId_PL1_ct9 = "K_19";

	private int[] CellId_O30 => GeneralFunctions.ArrayCellId(cellId_O30);

	private int[] CellId_P30 => GeneralFunctions.ArrayCellId(cellId_P30);

	private int[] CellId_Q30 => GeneralFunctions.ArrayCellId(cellId_Q30);

	private int[] CellId_ct16_donvi => GeneralFunctions.ArrayCellId(cellId_ct16_donvi);

	private int[] CellId_ct16_daily => GeneralFunctions.ArrayCellId(cellId_ct16_daily);

	private int[] CellId_ct26 => GeneralFunctions.ArrayCellId(cellId_ct26);

	private int[] CellId_ct28 => GeneralFunctions.ArrayCellId(cellId_ct28);

	private int[] CellId_ct31 => GeneralFunctions.ArrayCellId(cellId_ct31);

	private int[] CellId_Tong8_PL2 => GeneralFunctions.ArrayCellId(cellId_Tong8_PL2);

	private int[] CellId_ct33 => GeneralFunctions.ArrayCellId(cellId_ct33);

	private int[] CellId_PL1_ct9 => GeneralFunctions.ArrayCellId(cellId_PL1_ct9);

	private int idxPL1_ColD => HTKKFile.ColumnLettersToIndex("D") - 1;

	private int idxPL1_ColL => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxPL1_ColU => HTKKFile.ColumnLettersToIndex("U") - 1;

	private int idxPL1_ColY => HTKKFile.ColumnLettersToIndex("Y") - 1;

	private int idxPL1_ColAB => HTKKFile.ColumnLettersToIndex("AB") - 1;

	private int idxPL1_ColAD => HTKKFile.ColumnLettersToIndex("AD") - 1;

	private int idxPL1_ColAL => HTKKFile.ColumnLettersToIndex("AL") - 1;

	private int idxPL1_ColAJ => HTKKFile.ColumnLettersToIndex("AJ") - 1;

	private int[] CellId_E13 => GeneralFunctions.ArrayCellId("E_13");

	private int[] CellId_E23 => GeneralFunctions.ArrayCellId("E_23");

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId("E_25");

	private int[] CellId_E28 => GeneralFunctions.ArrayCellId("E_28");

	private int[] CellId_E29 => GeneralFunctions.ArrayCellId("E_29");

	private int[] CellId_E35 => GeneralFunctions.ArrayCellId("E_35");

	private int[] CellId_E36 => GeneralFunctions.ArrayCellId("E_36");

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			List<HTKKCMCSheetInfo> sheetInfos = ToKhaiInfo.SheetInfos;
			SheetView sheetView = null;
			if (sheetInfos != null && sheetInfos.Count > 0)
			{
				foreach (HTKKCMCSheetInfo item in sheetInfos)
				{
					if (item == null || !(item.SheetName == "02-1/PL-DK"))
					{
						continue;
					}
					HTKKCMCTableInfo hTKKCMCTableInfo = ((item != null && item.Tables != null && item.Tables.Count > 0) ? item.Tables[0] : null);
					if (hTKKCMCTableInfo != null)
					{
						HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
						if (PeriodInfo.TieuMucHachToanVal != "1" && rowsInfo != null)
						{
							rowsInfo.Cells[0].DefaultValue = "m3";
						}
						HTKKCMCRowValue hTKKCMCRowValue = ((hTKKCMCTableInfo.RowsValues != null && hTKKCMCTableInfo.RowsValues.Count > 0) ? hTKKCMCTableInfo.RowsValues[0] : null);
						if (hTKKCMCRowValue != null && PeriodInfo.TieuMucHachToanVal != "1")
						{
							hTKKCMCRowValue.RowValue[0] = hTKKCMCRowValue.RowValue[0].Replace("Thùng", "m3");
						}
					}
					break;
				}
			}
			sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			LayGiaTriKyTruoc(sheetView, PeriodInfo, hTKKCMCSheetInfo.DataFile, HeaderInfo.mst);
			if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				if (preVal_ct29 != 0.0)
				{
					sheetView.Cells[CellId_ct26[1], CellId_ct26[0]].Value = preVal_ct29;
				}
				if (preVal_ct36.HasValue)
				{
					sheetView.Cells[CellId_ct33[1], CellId_ct33[0]].Value = preVal_ct36;
				}
			}
			if (PeriodInfo.TieuMucHachToanVal != null)
			{
				switch (PeriodInfo.TieuMucHachToanVal)
				{
				case "1":
					sheetView.Cells[CellId_O30[1], CellId_O30[0]].Text = "1";
					sheetView.Cells[CellId_P30[1], CellId_P30[0]].Text = "0";
					sheetView.Cells[CellId_Q30[1], CellId_Q30[0]].Text = "0";
					break;
				case "3":
					sheetView.Cells[CellId_O30[1], CellId_O30[0]].Text = "0";
					sheetView.Cells[CellId_P30[1], CellId_P30[0]].Text = "1";
					sheetView.Cells[CellId_Q30[1], CellId_Q30[0]].Text = "0";
					break;
				case "2":
					sheetView.Cells[CellId_O30[1], CellId_O30[0]].Text = "0";
					sheetView.Cells[CellId_P30[1], CellId_P30[0]].Text = "0";
					sheetView.Cells[CellId_Q30[1], CellId_Q30[0]].Text = "1";
					break;
				}
			}
			if (string.IsNullOrEmpty(TTinChung.tenDLyThue) || string.IsNullOrEmpty(TTinChung.mstDLyThue))
			{
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], locked: true, "", false);
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], locked: true, "", false);
			}
			else if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				sheetView.Cells[CellId_ct16_daily[1], CellId_ct16_daily[0]].Value = true;
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
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (activeSheetIndex == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == "02_2_TAIN_TT36_DK");
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			SheetView sheetView = winFpSpread.Sheets[0];
			if (activeSheetIndex == 1 && iRow - 1 >= num && iRow - 1 <= num2 && (iCol - 1 == idxPL1_ColY || iCol - 1 == idxPL1_ColAD))
			{
				sheetView.Cells[CellId_ct31[1], CellId_ct31[0]].Formula = "BG64";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (winFpSpread.ActiveSheetIndex != 0)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			if (idxRow == CellId_ct16_daily[1] && idxCol == CellId_ct16_daily[0])
			{
				if (Convert.ToBoolean(sheetView.Cells[CellId_ct16_daily[1], CellId_ct16_daily[0]].Value))
				{
					sheetView.Cells[CellId_ct16_donvi[1], CellId_ct16_donvi[0]].Value = false;
				}
			}
			else if (idxRow == CellId_ct16_donvi[1] && idxCol == CellId_ct16_donvi[0] && Convert.ToBoolean(sheetView.Cells[CellId_ct16_donvi[1], CellId_ct16_donvi[0]].Value))
			{
				sheetView.Cells[CellId_ct16_daily[1], CellId_ct16_daily[0]].Value = false;
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			if (winFpSpread.TabStrip[0].BackColor == Color.White && levelWarning.ContainsKey(0))
			{
				levelWarning[0] = 0;
			}
			Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				int value = item.Value;
				if (value == 2 && winFpSpread.TabStrip[key].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
				}
				else if (value == 1 && winFpSpread.TabStrip[key].BackColor == Color.White)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Yellow;
				}
				else
				{
					winFpSpread.TabStrip[key].BackColor = Color.White;
				}
			}
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink || winFpSpread.TabStrip[2].BackColor == Color.Pink)
			{
				return false;
			}
			return true;
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView2 = null;
			bool flag = false;
			foreach (SheetView sheet in winFpSpread.Sheets)
			{
				if (sheet.SheetName == "02-1/PL-DK" && sheet.Visible)
				{
					sheetView2 = sheet;
					flag = true;
					break;
				}
			}
			if (sheetView2 == null || !flag)
			{
				sheetView.Cells[CellId_ct28[1], CellId_ct28[0]].ResetNote();
				sheetView.Cells[CellId_ct28[1], CellId_ct28[0]].BackColor = Color.White;
				winFpSpread.TabStrip[0].BackColor = Color.White;
			}
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, CellId_E36);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E35);
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E25);
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E23);
			string message5 = HTKKFile.GetMessage(sheetView, CellId_E23);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct16_donvi[1], CellId_ct16_donvi[0]].Value);
			bool flag3 = Convert.ToBoolean(sheetView.Cells[CellId_ct16_daily[1], CellId_ct16_daily[0]].Value);
			if (!HTKKFormater.isNullString(TTinChung.mstDLyThue))
			{
				if (flag2 && flag3)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], message3);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], message3);
					num = 2;
				}
				else if (!flag2 && !flag3)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], message4);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], message4);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0]);
					HTKKFile.RemoveValidate(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0]);
			}
			double num2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct26[1], CellId_ct26[0]].Value);
			bool flag4 = 1 == 0;
			if (num2 != preVal_ct29)
			{
				HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct26[1], CellId_ct26[0], message);
				if (num == 0)
				{
					num = 1;
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct26[1], CellId_ct26[0]);
			}
			double num3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct33[1], CellId_ct33[0]].Value);
			if (preVal_ct36.HasValue)
			{
				if (num3 != preVal_ct36)
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct33[1], CellId_ct33[0], message2);
					if (num == 0)
					{
						num = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct33[1], CellId_ct33[0]);
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

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			int num = 0;
			string message = HTKKFile.GetMessage(winFpSpread, CellId_E23[1], CellId_E23[0]);
			string message2 = HTKKFile.GetMessage(winFpSpread, CellId_E13[1], CellId_E13[0]);
			string message3 = HTKKFile.GetMessage(winFpSpread, CellId_E28[1], CellId_E28[0]);
			string message4 = HTKKFile.GetMessage(winFpSpread, CellId_E29[1], CellId_E29[0]);
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
			bool flag = false;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sheetView.Cells[num2 + i, idxPL1_ColD].Text;
				array[1] = sheetView.Cells[num2 + i, idxPL1_ColL].Text;
				array[2] = sheetView.Cells[num2 + i, idxPL1_ColU].Text;
				array[4] = sheetView.Cells[num2 + i, idxPL1_ColAL].Text;
				bool flag2 = GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array);
				string text = array[0] + array[1] + array[2] + array[4];
				if (string.IsNullOrEmpty(text.Replace("0", "")))
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxPL1_ColD);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxPL1_ColL);
					continue;
				}
				if (flag2)
				{
					flag = true;
				}
				if (!flag2)
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxPL1_ColD);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxPL1_ColL);
					continue;
				}
				if (string.IsNullOrEmpty(array[0]) && flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxPL1_ColD, message);
					num = 2;
				}
				else if (!ValidateBase.CheckMaSoThue(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxPL1_ColD, message2);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxPL1_ColD);
				}
				if (string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxPL1_ColL, message);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxPL1_ColL);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PathU29");
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PLuc_01_DK/tongCong");
			double num3 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
			if (num3 != 100.0 && flag)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message3);
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, message4);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
				HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
			}
			switch (num)
			{
			case 2:
				winFpSpread.TabStrip[2].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[2].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[2].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void LayGiaTriKyTruoc(SheetView sViewTK, HTKKPeriodInfo periodInfo, string dataFile, string headerInfo_mst)
	{
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Expected O, but got Unknown
		try
		{
			string preperiodFilename = GetPreperiodFilename(periodInfo, dataFile, headerInfo_mst);
			if (!string.IsNullOrEmpty(preperiodFilename))
			{
				XmlDocument val = new XmlDocument();
				val.Load(preperiodFilename);
				XmlNode val2 = ((XmlNode)val).SelectSingleNode("//Cells/Cell[@Path='" + path_ct29 + "']");
				if (val2 != null)
				{
					preVal_ct29 = Convert.ToDouble(((XmlNode)val2.Attributes["Value"]).Value);
				}
				XmlNode val3 = ((XmlNode)val).SelectSingleNode("//Cells/Cell[@Path='" + path_ct36 + "']");
				if (val3 != null)
				{
					preVal_ct36 = Convert.ToDouble(((XmlNode)val3.Attributes["Value"]).Value);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private string GetPreperiodFilename(HTKKPeriodInfo periodInfo, string dataFile, string headerInfo_mst)
	{
		try
		{
			if (string.IsNullOrEmpty(periodInfo.TuThang))
			{
				return "";
			}
			int num = Convert.ToInt32(periodInfo.TuThang.Split(new char[1] { '/' })[0]);
			int num2 = Convert.ToInt32(periodInfo.TuThang.Split(new char[1] { '/' })[1]);
			int num3 = 12 * num2 + num;
			string pattern = "^" + dataFile + ".*_224.xml$";
			Regex regex = new Regex(pattern);
			string text = HTKKConfig.DATAFILES + "\\" + headerInfo_mst;
			DirectoryInfo directoryInfo = new DirectoryInfo(text);
			FileInfo[] files = directoryInfo.GetFiles();
			if (files.Length > 0)
			{
				FileInfo[] array = files;
				foreach (FileInfo fileInfo in array)
				{
					if (regex.IsMatch(fileInfo.Name))
					{
						string text2 = fileInfo.Name.Split(new char[1] { '_' })[^2];
						if (string.IsNullOrEmpty(text2))
						{
							return "";
						}
						int num4 = Convert.ToInt32(text2.Substring(0, 2));
						int num5 = Convert.ToInt32(text2.Substring(2, 4));
						int num6 = 12 * num5 + num4;
						if (num6 + 1 == num3)
						{
							return text + "\\" + fileInfo.Name;
						}
					}
				}
			}
			return "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void UpdateFormula_1(FpSpread winFpSpread)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			if (winFpSpread.Sheets[1].Visible)
			{
				sheetView.Cells[63, HTKKFile.ColumnLettersToIndex("AV") - 1].Formula = "BG64";
				return;
			}
			sheetView.Cells[63, HTKKFile.ColumnLettersToIndex("AV") - 1].ResetFormula();
			sheetView.Cells[63, HTKKFile.ColumnLettersToIndex("AV") - 1].Value = 0;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void Refresh_1(FpSpread winFpSpread)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.Cells[63, HTKKFile.ColumnLettersToIndex("AV") - 1].ResetFormula();
			sheetView.Cells[63, HTKKFile.ColumnLettersToIndex("AV") - 1].Value = 0;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
