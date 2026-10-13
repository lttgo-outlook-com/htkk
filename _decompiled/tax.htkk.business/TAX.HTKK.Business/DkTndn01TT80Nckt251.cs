using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkTndn01TT80Nckt251
{
	private string cellId_O30 = "O_30";

	private string cellId_P30 = "P_30";

	private string cellId_Q30 = "Q_30";

	private string cellId_AF23 = "AF_23";

	private string cellId_ct08 = "Y_28";

	private string cellId_ct09 = "AS_28";

	private string cellId_ct20_donvi = "S_49";

	private string cellId_ct20_daily = "AC_49";

	private string cellId_ct10 = "AM_29";

	private string cellId_ct38 = "AV_69";

	private string cellId_PL1_ct9 = "K_19";

	private string cellId_E13 = "E_13";

	private string cellId_E19 = "E_19";

	private string cellId_E25 = "E_25";

	private string cellId_E28 = "E_28";

	private string cellId_E29 = "E_29";

	private string cellId_E31 = "E_31";

	private int[] CellId_O30 => GeneralFunctions.ArrayCellId(cellId_O30);

	private int[] CellId_P30 => GeneralFunctions.ArrayCellId(cellId_P30);

	private int[] CellId_Q30 => GeneralFunctions.ArrayCellId(cellId_Q30);

	private int[] CellId_X31 => GeneralFunctions.ArrayCellId("X_31");

	private int[] CellId_AF23 => GeneralFunctions.ArrayCellId(cellId_AF23);

	private int[] CellId_ct08 => GeneralFunctions.ArrayCellId(cellId_ct08);

	private int[] CellId_ct09 => GeneralFunctions.ArrayCellId(cellId_ct09);

	private int[] CellId_ct20_donvi => GeneralFunctions.ArrayCellId(cellId_ct20_donvi);

	private int[] CellId_ct20_daily => GeneralFunctions.ArrayCellId(cellId_ct20_daily);

	private int[] CellId_ct10 => GeneralFunctions.ArrayCellId(cellId_ct10);

	private int[] CellId_ct35 => GeneralFunctions.ArrayCellId("AV_66");

	private int[] CellId_ct35_CThuc => GeneralFunctions.ArrayCellId("BG_66");

	private int[] CellId_ct37 => GeneralFunctions.ArrayCellId("AV_68");

	private int[] CellId_ct37_CThuc => GeneralFunctions.ArrayCellId("BG_68");

	private int[] CellId_ct38 => GeneralFunctions.ArrayCellId(cellId_ct38);

	private int[] CellId_LoaiTien => GeneralFunctions.ArrayCellId("P_54");

	private int[] CellId_MaLoaiTien => GeneralFunctions.ArrayCellId("BG_54");

	private int[] CellId_PL1_ct9 => GeneralFunctions.ArrayCellId(cellId_PL1_ct9);

	private int idxPL1_ColD => HTKKFile.ColumnLettersToIndex("D") - 1;

	private int idxPL1_ColL => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxPL1_ColU => HTKKFile.ColumnLettersToIndex("U") - 1;

	private int idxPL1_ColAB => HTKKFile.ColumnLettersToIndex("AB") - 1;

	private int idxPL1_ColAV => HTKKFile.ColumnLettersToIndex("AV") - 1;

	private int idxPL1_ColAL => HTKKFile.ColumnLettersToIndex("AL") - 1;

	private int[] CellId_E13 => GeneralFunctions.ArrayCellId(cellId_E13);

	private int[] CellId_E19 => GeneralFunctions.ArrayCellId(cellId_E19);

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId(cellId_E25);

	private int[] CellId_E28 => GeneralFunctions.ArrayCellId(cellId_E28);

	private int[] CellId_E29 => GeneralFunctions.ArrayCellId(cellId_E29);

	private int[] CellId_E31 => GeneralFunctions.ArrayCellId(cellId_E31);

	private int[] CellId_E32 => GeneralFunctions.ArrayCellId("E_32");

	private int[] CellId_E33 => GeneralFunctions.ArrayCellId("E_33");

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			sheetView.Cells[CellId_AF23[1], CellId_AF23[0]].Text = PeriodInfo.LanXuatBan;
			if (PeriodInfo.TieuMucHachToanVal != null)
			{
				switch (PeriodInfo.TieuMucHachToanVal)
				{
				case "1":
					sheetView.Cells[CellId_O30[1], CellId_O30[0]].Value = true;
					sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3752";
					break;
				case "3":
					sheetView.Cells[CellId_P30[1], CellId_P30[0]].Value = true;
					sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3952";
					break;
				case "2":
					sheetView.Cells[CellId_Q30[1], CellId_Q30[0]].Value = true;
					sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3802";
					break;
				}
			}
			if (Convert.ToBoolean(sheetView.Cells[CellId_ct09[1], CellId_ct09[0]].Value))
			{
				sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].ResetFormula();
				sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = false;
			}
			else
			{
				sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = true;
			}
			if (string.IsNullOrEmpty(TTinChung.tenDLyThue) || string.IsNullOrEmpty(TTinChung.mstDLyThue))
			{
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0], locked: true, "", false);
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0], locked: true, "", false);
			}
			else if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				sheetView.Cells[CellId_ct20_daily[1], CellId_ct20_daily[0]].Value = true;
			}
			sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text = PeriodInfo.DonViTinhMa;
			sheetView.Cells[CellId_LoaiTien[1], CellId_LoaiTien[0]].Text = PeriodInfo.DonViTinhTen;
			sheetView.Cells[CellId_LoaiTien[1], CellId_LoaiTien[0]].Locked = true;
			string cellFormat = "AV_62###AV_64###AV_66###AV_67###AV_68";
			string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
			if (PeriodInfo.LoaiToKhai == "B")
			{
				sheetView.Cells[CellId_LoaiTien[1], CellId_LoaiTien[0]].Locked = true;
			}
			GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0], 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo != null)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				string text = "";
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					text = text + "AB_" + (dataLocationTemplateRowStart + i) + "###";
				}
				text = text.Substring(0, text.Length - 3);
				string donViTinhMa = PeriodInfo.DonViTinhMa;
				GeneralFunctions.changFormatCell(sView, text, donViTinhMa);
				text = "";
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PLuc_01_DK/tongCong");
				text = "AB_" + hTKKCMCCellInfo.CellTemplate.RowIndex;
				GeneralFunctions.changFormatCell(sView, text, donViTinhMa);
				text = "K_19";
				GeneralFunctions.changFormatCell(sView, text, donViTinhMa);
				GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], 1);
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
				string formula = "ROUND(AV62*AV63/100,2) + ROUND(AV64*AV65/100,2)";
				string formula2 = "IF(AV66-AV67>0,ROUND(AV66-AV67,2),0)";
				if (CellId_LoaiTien[1] + 1 == iRow && CellId_LoaiTien[0] + 1 == iCol)
				{
					string cellFormat = "AV_62###AV_64###AV_66###AV_67###AV_68";
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
					if (text == "VND")
					{
						formula = "ROUND(AV62*AV63/100,0) + ROUND(AV64*AV65/100,0)";
						formula2 = "IF(AV66-AV67>0,ROUND(AV66-AV67,0),0)";
					}
					sheetView.Cells[CellId_ct35[1], CellId_ct35[0]].Formula = formula;
					sheetView.Cells[CellId_ct37[1], CellId_ct37[0]].Formula = formula2;
					SheetView sView = winFpSpread.Sheets[1];
					HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
					if (hTKKCMCSheetInfo2 != null)
					{
						HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables[0];
						int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
						cellFormat = "";
						for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
						{
							cellFormat = cellFormat + "AB_" + (dataLocationTemplateRowStart + i) + "###";
						}
						cellFormat = cellFormat.Substring(0, cellFormat.Length - 3);
						GeneralFunctions.changFormatCell(sView, cellFormat, text);
						cellFormat = "";
						HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PLuc_01_DK/tongCong");
						cellFormat = "AB_" + hTKKCMCCellInfo.CellTemplate.RowIndex;
						GeneralFunctions.changFormatCell(sView, cellFormat, text);
						cellFormat = "K_19";
						GeneralFunctions.changFormatCell(sView, cellFormat, text);
						tinhCt5(winFpSpread, ToKhaiInfo);
					}
				}
				if (iRow >= 62 && iRow <= 65 && CellId_ct35[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula = "ROUND(AV62*AV63/100,0) + ROUND(AV64*AV65/100,0)";
					}
					sheetView.Cells[CellId_ct35[1], CellId_ct35[0]].Formula = formula;
				}
				if (iRow >= 62 && iRow <= 66 && CellId_ct37[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula2 = "IF(AV66-AV67>0,ROUND(AV66-AV67,0),0)";
					}
					sheetView.Cells[CellId_ct37[1], CellId_ct37[0]].Formula = formula2;
				}
				if (iRow >= 62 && iRow <= 67 && CellId_ct37[0] + 1 == iCol)
				{
					tinhCt5(winFpSpread, ToKhaiInfo);
				}
			}
			GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0], 0);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (activeSheetIndex == 1)
			{
				SheetView sheetView = winFpSpread.Sheets[1];
				if (idxPL1_ColU + 1 == iCol)
				{
					sheetView.Cells[(int)iRow - 1, idxPL1_ColAB].Formula = sheetView.Cells[(int)iRow - 1, idxPL1_ColAV].Formula;
				}
			}
			GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], 1);
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
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_ct08[1], CellId_ct08[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct09[1], CellId_ct09[0]].Value);
			bool flag3 = flag || flag2;
			if (idxRow == CellId_ct08[1] && idxCol == CellId_ct08[0])
			{
				if (!flag)
				{
					sheetView.Cells[CellId_ct09[1], CellId_ct09[0]].Value = true;
					sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = false;
				}
				else
				{
					sheetView.Cells[CellId_ct09[1], CellId_ct09[0]].Value = false;
					sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = true;
					sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Value = "";
				}
			}
			else if (idxRow == CellId_ct09[1] && idxCol == CellId_ct09[0])
			{
				if (!flag2)
				{
					sheetView.Cells[CellId_ct08[1], CellId_ct08[0]].Value = true;
					sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = true;
				}
				else
				{
					sheetView.Cells[CellId_ct08[1], CellId_ct08[0]].Value = false;
					sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = false;
				}
			}
			if (idxRow == CellId_ct20_daily[1] && idxCol == CellId_ct20_daily[0])
			{
				if (Convert.ToBoolean(sheetView.Cells[CellId_ct20_daily[1], CellId_ct20_daily[0]].Value))
				{
					sheetView.Cells[CellId_ct20_donvi[1], CellId_ct20_donvi[0]].Value = false;
				}
			}
			else if (idxRow == CellId_ct20_donvi[1] && idxCol == CellId_ct20_donvi[0] && Convert.ToBoolean(sheetView.Cells[CellId_ct20_donvi[1], CellId_ct20_donvi[0]].Value))
			{
				sheetView.Cells[CellId_ct20_daily[1], CellId_ct20_daily[0]].Value = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void tinhCt5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo != null)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					sheetView.Cells[num + i, idxPL1_ColAB].Formula = sheetView.Cells[num + i, idxPL1_ColAV].Formula;
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
			if (ToKhaiInfo.SheetInfos[0] != null)
			{
				Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			}
			if (ToKhaiInfo.SheetInfos[1] != null)
			{
				Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			}
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink || winFpSpread.TabStrip[1].BackColor == Color.Pink)
			{
				return false;
			}
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
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, CellId_E19);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E25);
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E31);
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E32);
			string message5 = HTKKFile.GetMessage(sheetView, CellId_E33);
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_ct08[1], CellId_ct08[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct09[1], CellId_ct09[0]].Value);
			bool flag3 = Convert.ToBoolean(sheetView.Cells[CellId_ct20_donvi[1], CellId_ct20_donvi[0]].Value);
			bool flag4 = Convert.ToBoolean(sheetView.Cells[CellId_ct20_daily[1], CellId_ct20_daily[0]].Value);
			if (!flag && !flag2)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct08[1], CellId_ct08[0], message);
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct09[1], CellId_ct09[0], message);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct08[1], CellId_ct08[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct09[1], CellId_ct09[0]);
			}
			HTKKFile.RemoveValidate(sheetView, CellId_ct10[1], CellId_ct10[0]);
			if (flag2)
			{
				string text = sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Text;
				DateTime dtDate;
				if (string.IsNullOrEmpty(text) || text.Equals("../../...."))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct10[1], CellId_ct10[0], message);
					num = 2;
				}
				else if (!HTKKFormater.IsDateTime(text, "dd/MM/yyyy", out dtDate))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct10[1], CellId_ct10[0], HTKKFile.GetMessage("0071"));
					num = 2;
				}
			}
			if (!HTKKFormater.isNullString(TTinChung.mstDLyThue))
			{
				if (flag3 && flag4)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0], message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0], message2);
					num = 2;
				}
				else if (!flag3 && !flag4)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0], message);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0], message);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0]);
					HTKKFile.RemoveValidate(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0]);
			}
			string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			double value = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct35_CThuc[1], CellId_ct35_CThuc[0]].Value);
			double value2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct35[1], CellId_ct35[0]].Value);
			double num2 = 0.0;
			double value3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct37_CThuc[1], CellId_ct37_CThuc[0]].Value);
			double value4 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct37[1], CellId_ct37[0]].Value);
			double num3 = 0.0;
			if (text2 != "VND")
			{
				num2 = ((Math.Abs(value) * 0.001 < 5.0) ? (Math.Abs(value) * 0.001) : 5.0);
				num3 = ((Math.Abs(value3) * 0.001 < 5.0) ? (Math.Abs(value3) * 0.001) : 5.0);
			}
			else
			{
				num2 = ((Math.Abs(value) * 0.001 < 100000.0) ? (Math.Abs(value) * 0.001) : 100000.0);
				num3 = ((Math.Abs(value3) * 0.001 < 100000.0) ? (Math.Abs(value3) * 0.001) : 100000.0);
			}
			HTKKFile.RemoveValidateBackGround(sheetView, CellId_ct35[1], CellId_ct35[0]);
			HTKKFile.RemoveValidateBackGround(sheetView, CellId_ct37[1], CellId_ct37[0]);
			if (Math.Abs(value2) > Math.Abs(value) + num2 || Math.Abs(value2) < Math.Abs(value) - num2)
			{
				if (text2 != "VND")
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct35[1], CellId_ct35[0], message4);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct35[1], CellId_ct35[0], message5);
				}
				num = 2;
			}
			if (Math.Abs(value4) > Math.Abs(value3) + num3 || Math.Abs(value4) < Math.Abs(value3) - num3)
			{
				if (text2 != "VND")
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct37[1], CellId_ct37[0], message4);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct37[1], CellId_ct37[0], message5);
				}
				num = 2;
			}
			int[] array = new int[4] { 58, 59, 60, 61 };
			for (int i = 0; i < array.Length; i++)
			{
				int[] array2 = GeneralFunctions.ArrayCellId("AN_" + array[i]);
				HTKKFile.RemoveValidate(sheetView, array2[1], array2[0]);
				string text3 = sheetView.Cells[array2[1], array2[0]].Text;
				if (string.IsNullOrEmpty(text3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message);
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

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, CellId_E19);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E13);
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
			bool flag = false;
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "pathU29");
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PLuc_01_DK/tongCong");
			for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsValues.Count; num3++)
			{
				array[0] = sheetView.Cells[num2 + num3, idxPL1_ColD].Text;
				array[1] = sheetView.Cells[num2 + num3, idxPL1_ColL].Text;
				array[2] = sheetView.Cells[num2 + num3, idxPL1_ColU].Text;
				array[4] = sheetView.Cells[num2 + num3, idxPL1_ColAL].Text;
				bool flag2 = GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array);
				if (flag2)
				{
					flag = true;
				}
				if (!flag2)
				{
					HTKKFile.RemoveValidate(sheetView, num2 + num3, idxPL1_ColD);
					HTKKFile.RemoveValidate(sheetView, num2 + num3, idxPL1_ColL);
					continue;
				}
				if (string.IsNullOrEmpty(array[0]) && string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.RemoveValidate(sheetView, num2 + num3, idxPL1_ColD);
					HTKKFile.RemoveValidate(sheetView, num2 + num3, idxPL1_ColL);
					HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
					return;
				}
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + num3, idxPL1_ColD, message);
					num = 2;
				}
				else if (!ValidateBase.CheckMaSoThue(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + num3, idxPL1_ColD, message2);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num2 + num3, idxPL1_ColD);
				}
				if (string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + num3, idxPL1_ColL, message);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num2 + num3, idxPL1_ColL);
				}
			}
			double num4 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
			double num5 = Math.Round(Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value), 2);
			double num6 = Math.Round(Convert.ToDouble(sheetView.Cells[CellId_PL1_ct9[1], CellId_PL1_ct9[0]].Value), 2);
			if (num4 != 100.0 && flag)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, CellId_E28));
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
			}
			if (num5 != num6 && flag)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, CellId_E29));
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
			}
			switch (num)
			{
			case 2:
				winFpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[1].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
