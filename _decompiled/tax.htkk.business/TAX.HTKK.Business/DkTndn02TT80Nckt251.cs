using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkTndn02TT80Nckt251
{
	private string cellId_O30 = "O_30";

	private string cellId_P30 = "P_30";

	private string cellId_Q30 = "Q_30";

	private int[] cellTenTHQT_TK = GeneralFunctions.ArrayCellId("U_30");

	private int[] cellMaTHQT_TK = GeneralFunctions.ArrayCellId("T_30");

	private string cellId_ct16_donvi = "Q_55";

	private string cellId_ct16_daily = "Y_55";

	private string cellId_ct41 = "AV_94";

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

	private int[] CellId_ct16_donvi => GeneralFunctions.ArrayCellId(cellId_ct16_donvi);

	private int[] CellId_ct16_daily => GeneralFunctions.ArrayCellId(cellId_ct16_daily);

	private int[] CellId_ct27 => GeneralFunctions.ArrayCellId("AV_76");

	private int[] CellId_ct27_CThuc => GeneralFunctions.ArrayCellId("BG_76");

	private int[] CellId_ct29 => GeneralFunctions.ArrayCellId("AV_80");

	private int[] CellId_ct29_CThuc => GeneralFunctions.ArrayCellId("BG_80");

	private int[] CellId_ct30 => GeneralFunctions.ArrayCellId("AV_81");

	private int[] CellId_ct30_CThuc => GeneralFunctions.ArrayCellId("BG_81");

	private int[] CellId_ct32 => GeneralFunctions.ArrayCellId("AV_83");

	private int[] CellId_ct32_CThuc => GeneralFunctions.ArrayCellId("BG_83");

	private int[] CellId_ct36 => GeneralFunctions.ArrayCellId("AV_89");

	private int[] CellId_ct36_CThuc => GeneralFunctions.ArrayCellId("BG_89");

	private int[] CellId_ct38 => GeneralFunctions.ArrayCellId("AV_91");

	private int[] CellId_ct38_CThuc => GeneralFunctions.ArrayCellId("BG_91");

	private int[] CellId_ct40 => GeneralFunctions.ArrayCellId("AV_93");

	private int[] CellId_ct40_CThuc => GeneralFunctions.ArrayCellId("BG_93");

	private int[] CellId_ct41 => GeneralFunctions.ArrayCellId(cellId_ct41);

	private int[] CellId_LoaiTien => GeneralFunctions.ArrayCellId("P_69");

	private int[] CellId_MaLoaiTien => GeneralFunctions.ArrayCellId("BG_69");

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
			if (string.IsNullOrEmpty(TTinChung.tenDLyThue) || string.IsNullOrEmpty(TTinChung.mstDLyThue))
			{
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], locked: true, "", false);
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], locked: true, "", false);
			}
			else if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				sheetView.Cells[CellId_ct16_daily[1], CellId_ct16_daily[0]].Value = true;
			}
			sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text = PeriodInfo.DonViTinhMa;
			sheetView.Cells[CellId_LoaiTien[1], CellId_LoaiTien[0]].Text = PeriodInfo.DonViTinhTen;
			sheetView.Cells[CellId_LoaiTien[1], CellId_LoaiTien[0]].Locked = true;
			string cellFormat = "AV_74###AV_75###AV_76###AV_77###AV_78###AV_79###AV_80###AV_81###AV_82###AV_83###AV_86###AV_89###AV_90###AV_91###AV_92###AV_93";
			string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
			if (PeriodInfo.LoaiToKhai == "B")
			{
				sheetView.Cells[CellId_LoaiTien[1], CellId_LoaiTien[0]].Locked = true;
			}
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sheetView.Cells[cellTenTHQT_TK[1], cellTenTHQT_TK[0]].Value = PeriodInfo.TenQuyetToan;
				sheetView.Cells[cellMaTHQT_TK[1], cellMaTHQT_TK[0]].Value = PeriodInfo.MaQuyetToan;
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
				string formula = "BG76";
				string formula2 = "BG80";
				string formula3 = "BG81";
				string formula4 = "BG83";
				string formula5 = "BG89";
				string formula6 = "BG91";
				string formula7 = "BG93";
				if (CellId_LoaiTien[1] + 1 == iRow && CellId_LoaiTien[0] + 1 == iCol)
				{
					string cellFormat = "AV_74###AV_75###AV_76###AV_77###AV_78###AV_79###AV_80###AV_81###AV_82###AV_83###AV_86###AV_89###AV_90###AV_91###AV_92###AV_93";
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
					if (text == "VND")
					{
						formula = "BH76";
						formula2 = "BH80";
						formula3 = "BH81";
						formula4 = "BH83";
						formula5 = "BH89";
						formula6 = "BH91";
						formula7 = "BH93";
					}
					sheetView.Cells[CellId_ct27[1], CellId_ct27[0]].Formula = formula;
					sheetView.Cells[CellId_ct29[1], CellId_ct29[0]].Formula = formula2;
					sheetView.Cells[CellId_ct30[1], CellId_ct30[0]].Formula = formula3;
					sheetView.Cells[CellId_ct32[1], CellId_ct32[0]].Formula = formula4;
					sheetView.Cells[CellId_ct36[1], CellId_ct36[0]].Formula = formula5;
					sheetView.Cells[CellId_ct38[1], CellId_ct38[0]].Formula = formula6;
					sheetView.Cells[CellId_ct40[1], CellId_ct40[0]].Formula = formula7;
					SheetView sView = winFpSpread.Sheets[1];
					HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
					if (hTKKCMCSheetInfo2 != null)
					{
						HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables[0];
						int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
						cellFormat = "";
						for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
						{
							cellFormat = "AB_" + (dataLocationTemplateRowStart + i) + "###";
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
				if (iRow >= 77 && iRow <= 78 && CellId_ct27[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula = "BH76";
					}
					sheetView.Cells[CellId_ct27[1], CellId_ct27[0]].Formula = formula;
				}
				if ((iRow == 75 || iRow == 79) && CellId_ct27[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula2 = "BH80";
					}
					sheetView.Cells[CellId_ct29[1], CellId_ct29[0]].Formula = formula2;
				}
				switch (iRow)
				{
				case 74L:
				case 76L:
				case 77L:
				case 78L:
					if (CellId_ct27[0] + 1 == iCol)
					{
						string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
						if (text == "VND")
						{
							formula3 = "BH81";
						}
						sheetView.Cells[CellId_ct30[1], CellId_ct30[0]].Formula = formula3;
					}
					break;
				}
				if ((iRow == 74 || (iRow >= 76 && iRow <= 78) || iRow == 81 || iRow == 82) && CellId_ct27[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula4 = "BH83";
					}
					sheetView.Cells[CellId_ct32[1], CellId_ct32[0]].Formula = formula4;
				}
				if ((iRow == 74 || (iRow >= 76 && iRow <= 78) || (iRow >= 81 && iRow <= 84) || iRow == 86 || iRow == 87) && CellId_ct27[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula5 = "BH89";
					}
					sheetView.Cells[CellId_ct36[1], CellId_ct36[0]].Formula = formula5;
				}
				if ((iRow == 74 || (iRow >= 76 && iRow <= 78) || (iRow >= 81 && iRow <= 84) || iRow == 86 || iRow == 87 || iRow == 89 || iRow == 90) && CellId_ct27[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula6 = "BH91";
					}
					sheetView.Cells[CellId_ct38[1], CellId_ct38[0]].Formula = formula6;
				}
				if ((iRow == 74 || (iRow >= 76 && iRow <= 78) || (iRow >= 81 && iRow <= 84) || iRow == 86 || iRow == 87 || (iRow >= 89 && iRow <= 92)) && CellId_ct27[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula7 = "BH93";
					}
					sheetView.Cells[CellId_ct40[1], CellId_ct40[0]].Formula = formula7;
				}
				if ((iRow == 74 || (iRow >= 76 && iRow <= 78) || (iRow >= 81 && iRow <= 84) || iRow == 86 || iRow == 87 || iRow == 89 || iRow == 90 || iRow == 91) && CellId_ct29[0] + 1 == iCol)
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
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_ct16_donvi[1], CellId_ct16_donvi[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct16_daily[1], CellId_ct16_daily[0]].Value);
			HTKKFile.RemoveValidate(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0]);
			HTKKFile.RemoveValidate(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0]);
			if (!HTKKFormater.isNullString(TTinChung.mstDLyThue))
			{
				if (flag && flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], message2);
					num = 2;
				}
				else if (!flag && !flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], message);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], message);
					num = 2;
				}
			}
			string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			int[] array = new int[7] { 76, 80, 81, 83, 89, 91, 93 };
			for (int i = 0; i < array.Length; i++)
			{
				int[] array2 = GeneralFunctions.ArrayCellId("AV_" + array[i]);
				int[] array3 = GeneralFunctions.ArrayCellId("BG_" + array[i]);
				double value = ConvertHelpers.ConvertToDouble(sheetView.Cells[array3[1], array3[0]].Value);
				double value2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array2[1], array2[0]].Value);
				double num2 = 0.0;
				HTKKFile.RemoveValidateBackGround(sheetView, array2[1], array2[0]);
				num2 = ((!(text != "VND")) ? ((Math.Abs(value) * 0.001 < 100000.0) ? (Math.Abs(value) * 0.001) : 100000.0) : ((Math.Abs(value) * 0.001 < 5.0) ? (Math.Abs(value) * 0.001) : 5.0));
				if (Math.Abs(value2) > Math.Abs(value) + num2 || Math.Abs(value2) < Math.Abs(value) - num2)
				{
					if (text != "VND")
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message3);
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message4);
					}
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
			int[] array2 = GeneralFunctions.ArrayCellId("K_19");
			double num3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array2[1], array2[0]].Value);
			for (int num4 = 0; num4 < hTKKCMCTableInfo.RowsValues.Count; num4++)
			{
				array[0] = sheetView.Cells[num2 + num4, idxPL1_ColD].Text;
				array[1] = sheetView.Cells[num2 + num4, idxPL1_ColL].Text;
				array[2] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + num4, idxPL1_ColU].Value);
				array[3] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + num4, idxPL1_ColAB].Value);
				array[4] = sheetView.Cells[num2 + num4, idxPL1_ColAL].Text;
				if (GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array))
				{
					flag = true;
				}
				HTKKFile.RemoveValidate(sheetView, num2 + num4, idxPL1_ColD);
				HTKKFile.RemoveValidate(sheetView, num2 + num4, idxPL1_ColL);
				HTKKFile.RemoveValidate(sheetView, num2 + num4, idxPL1_ColU);
				HTKKFile.RemoveValidateBackGround(sheetView, num2 + num4, idxPL1_ColAB);
				if (!string.IsNullOrEmpty(array[0]) && !ValidateBase.CheckMaSoThue(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + num4, idxPL1_ColD, message2);
					num = 2;
				}
				if (string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + num4, idxPL1_ColL, message);
					num = 2;
				}
				if (ConvertHelpers.ConvertToDouble(array[2]) == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + num4, idxPL1_ColU, message);
					num = 2;
				}
				if (ConvertHelpers.ConvertToDouble(array[3]) == 0.0 && num3 != 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + num4, idxPL1_ColAB, message);
					num = 2;
				}
			}
			double num5 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
			double num6 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
			double num7 = Convert.ToDouble(sheetView.Cells[CellId_PL1_ct9[1], CellId_PL1_ct9[0]].Value);
			if (num5 != 100.0 && flag)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, CellId_E28));
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
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
