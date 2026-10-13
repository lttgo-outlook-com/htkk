using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkTain02TT80Nckt284
{
	private string cellId_O30 = "O_30";

	private string cellId_P30 = "P_30";

	private string cellId_Q30 = "Q_30";

	private string cellId_W31 = "W_31";

	private string cellId_O31 = "O_31";

	private string cellId_R30 = "R_30";

	private string cellId_W30 = "W_30";

	private string cellId_ct16_donvi = "Q_54";

	private string cellId_ct16_daily = "Y_54";

	private string cellId_ct28 = "AV_75";

	private string cellId_ct29 = "AV_78";

	private string cellId_ct30 = "AV_79";

	private string cellId_ct34 = "AV_84";

	private string cellId_ct35 = "AV_85";

	private string cellId_ct36 = "AV_86";

	private string cellId_ct37 = "AV_87";

	private string cellId_ct38 = "AV_88";

	private string cellId_ct39 = "AV_89";

	private string cellId_ct40 = "AV_90";

	private string cellId_ct42 = "AV_92";

	private int[] cellTenTHQT_TK = GeneralFunctions.ArrayCellId("U_30");

	private int[] cellMaTHQT_TK = GeneralFunctions.ArrayCellId("T_30");

	private string cellId_PL1_ct9 = "K_19";

	private string cellFormat = "AV_80###AV_88###AV_89###AV_90###AV_91###AV_92";

	private int[] CellId_O30 => GeneralFunctions.ArrayCellId(cellId_O30);

	private int[] CellId_P30 => GeneralFunctions.ArrayCellId(cellId_P30);

	private int[] CellId_Q30 => GeneralFunctions.ArrayCellId(cellId_Q30);

	private int[] CellId_W31 => GeneralFunctions.ArrayCellId(cellId_W31);

	private int[] CellId_O31 => GeneralFunctions.ArrayCellId(cellId_O31);

	private int[] CellId_X31 => GeneralFunctions.ArrayCellId("X_31");

	private int[] CellId_R30 => GeneralFunctions.ArrayCellId(cellId_R30);

	private int[] CellId_W30 => GeneralFunctions.ArrayCellId(cellId_W30);

	private int[] CellId_ct16_donvi => GeneralFunctions.ArrayCellId(cellId_ct16_donvi);

	private int[] CellId_ct16_daily => GeneralFunctions.ArrayCellId(cellId_ct16_daily);

	private int[] CellId_ct28 => GeneralFunctions.ArrayCellId(cellId_ct28);

	private int[] CellId_ct29 => GeneralFunctions.ArrayCellId(cellId_ct29);

	private int[] CellId_ct30 => GeneralFunctions.ArrayCellId(cellId_ct30);

	private int[] CellId_ct34 => GeneralFunctions.ArrayCellId(cellId_ct34);

	private int[] CellId_ct35 => GeneralFunctions.ArrayCellId(cellId_ct35);

	private int[] CellId_ct36 => GeneralFunctions.ArrayCellId(cellId_ct36);

	private int[] CellId_ct37 => GeneralFunctions.ArrayCellId(cellId_ct37);

	private int[] CellId_ct38 => GeneralFunctions.ArrayCellId(cellId_ct38);

	private int[] CellId_ct39 => GeneralFunctions.ArrayCellId(cellId_ct39);

	private int[] CellId_ct40 => GeneralFunctions.ArrayCellId(cellId_ct40);

	private int[] CellId_ct42 => GeneralFunctions.ArrayCellId(cellId_ct42);

	private int[] CellId_LoaiTien => GeneralFunctions.ArrayCellId("P_58");

	private int[] CellId_MaLoaiTien => GeneralFunctions.ArrayCellId("BG_58");

	private int[] CellId_PL1_ct9 => GeneralFunctions.ArrayCellId(cellId_PL1_ct9);

	private int idxPL1_ColD => HTKKFile.ColumnLettersToIndex("D") - 1;

	private int idxPL1_ColL => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxPL1_ColU => HTKKFile.ColumnLettersToIndex("U") - 1;

	private int idxPL1_ColAB => HTKKFile.ColumnLettersToIndex("AB") - 1;

	private int idxPL1_ColAV => HTKKFile.ColumnLettersToIndex("AV") - 1;

	private int idxPL1_ColY => HTKKFile.ColumnLettersToIndex("Y") - 1;

	private int idxPL1_ColAD => HTKKFile.ColumnLettersToIndex("AD") - 1;

	private int idxPL1_ColAJ => HTKKFile.ColumnLettersToIndex("AJ") - 1;

	private int idxPL1_ColAL => HTKKFile.ColumnLettersToIndex("AL") - 1;

	private int[] CellId_E13 => GeneralFunctions.ArrayCellId("E_13");

	private int[] CellId_E19 => GeneralFunctions.ArrayCellId("E_19");

	private int[] CellId_E22 => GeneralFunctions.ArrayCellId("E_22");

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId("E_25");

	private int[] CellId_E28 => GeneralFunctions.ArrayCellId("E_28");

	private int[] CellId_E29 => GeneralFunctions.ArrayCellId("E_29");

	private int[] CellId_E30 => GeneralFunctions.ArrayCellId("E_30");

	private int[] CellId_E31 => GeneralFunctions.ArrayCellId("E_31");

	private int[] CellId_E32 => GeneralFunctions.ArrayCellId("E_32");

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
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
					sheetView.Cells[CellId_P30[1], CellId_P30[0]].Value = false;
					sheetView.Cells[CellId_Q30[1], CellId_Q30[0]].Value = false;
					sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3751";
					break;
				case "3":
					sheetView.Cells[CellId_O30[1], CellId_O30[0]].Value = false;
					sheetView.Cells[CellId_P30[1], CellId_P30[0]].Value = true;
					sheetView.Cells[CellId_Q30[1], CellId_Q30[0]].Value = false;
					sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3951";
					break;
				case "2":
					sheetView.Cells[CellId_O30[1], CellId_O30[0]].Value = false;
					sheetView.Cells[CellId_P30[1], CellId_P30[0]].Value = false;
					sheetView.Cells[CellId_Q30[1], CellId_Q30[0]].Value = true;
					sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3801";
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
				sheetView.Cells[CellId_ct16_donvi[1], CellId_ct16_donvi[0]].Value = false;
			}
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sheetView.Cells[cellTenTHQT_TK[1], cellTenTHQT_TK[0]].Value = PeriodInfo.TenQuyetToan;
				sheetView.Cells[cellMaTHQT_TK[1], cellMaTHQT_TK[0]].Value = PeriodInfo.MaQuyetToan;
			}
			sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text = PeriodInfo.DonViTinhMa;
			sheetView.Cells[CellId_LoaiTien[1], CellId_LoaiTien[0]].Text = PeriodInfo.DonViTinhTen;
			sheetView.Cells[CellId_LoaiTien[1], CellId_LoaiTien[0]].Locked = true;
			string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
			if (PeriodInfo.LoaiToKhai == "B")
			{
				sheetView.Cells[CellId_LoaiTien[1], CellId_LoaiTien[0]].Locked = true;
			}
			GeneralFunctions.SaveAllCellTypeOfSheet(sheetView, hTKKCMCSheetInfo, 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
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
				GeneralFunctions.SaveAllCellTypeOfSheet(sView, hTKKCMCSheetInfo, 1);
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
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			if (activeSheetIndex == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				string text = "BG75";
				string text2 = "BG78";
				string text3 = "BG79";
				string text4 = "BG84";
				string text5 = "BG85";
				string text6 = "BG86";
				string text7 = "BG87";
				string text8 = "BG88";
				string text9 = "BG89";
				string text10 = "BG90";
				string text11 = "BG92";
				if (CellId_LoaiTien[1] + 1 == iRow && CellId_LoaiTien[0] + 1 == iCol)
				{
					string text12 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					GeneralFunctions.changFormatCell(sheetView, cellFormat, text12);
					if (text12 == "VND")
					{
						text8 = "BH88";
						text9 = "BH89";
						text10 = "BH90";
						text11 = "BH92";
					}
					sheetView.Cells[CellId_ct38[1], CellId_ct38[0]].Formula = text8;
					sheetView.Cells[CellId_ct39[1], CellId_ct39[0]].Formula = text9;
					sheetView.Cells[CellId_ct40[1], CellId_ct40[0]].Formula = text10;
					sheetView.Cells[CellId_ct42[1], CellId_ct42[0]].Formula = text11;
					SheetView sView = winFpSpread.Sheets[2];
					HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
					if (hTKKCMCSheetInfo != null)
					{
						HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
						int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
						cellFormat = "";
						for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
						{
							cellFormat = cellFormat + "AB_" + (dataLocationTemplateRowStart + i) + "###";
						}
						cellFormat = cellFormat.Substring(0, cellFormat.Length - 3);
						GeneralFunctions.changFormatCell(sView, cellFormat, text12);
						cellFormat = "";
						HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PLuc_01_DK/tongCong");
						cellFormat = "AB_" + hTKKCMCCellInfo.CellTemplate.RowIndex;
						GeneralFunctions.changFormatCell(sView, cellFormat, text12);
						cellFormat = "K_19";
						GeneralFunctions.changFormatCell(sView, cellFormat, text12);
						tinhCt5(winFpSpread, ToKhaiInfo);
					}
					sView = winFpSpread.Sheets[1];
					hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
					if (hTKKCMCSheetInfo != null)
					{
						HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
						int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
						cellFormat = "";
						for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
						{
							cellFormat = cellFormat + "AJ_" + (dataLocationTemplateRowStart + i) + "###";
						}
						cellFormat = cellFormat.Substring(0, cellFormat.Length - 3);
						GeneralFunctions.changFormatCell(sView, cellFormat, text12);
						cellFormat = "";
						HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL02_1_PL_DK/BangSLuongVaDThu/tongCot8");
						cellFormat = "AJ_" + hTKKCMCCellInfo.CellTemplate.RowIndex;
						GeneralFunctions.changFormatCell(sView, cellFormat, text12);
					}
				}
				if (iRow >= 76 && iRow <= 77 && CellId_ct28[0] + 1 == iCol)
				{
					string text12 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text12 == "VND")
					{
						text = text.Replace("BG", "BH");
					}
					sheetView.Cells[CellId_ct28[1], CellId_ct28[0]].Formula = text;
				}
				if (iRow >= 73 && iRow <= 77 && CellId_ct28[0] + 1 == iCol)
				{
					string text12 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text12 == "VND")
					{
						text2 = text2.Replace("BG", "BH");
					}
					sheetView.Cells[CellId_ct29[1], CellId_ct29[0]].Formula = text2;
				}
				if ((iRow == 71 || iRow == 74) && CellId_ct30[0] + 1 == iCol)
				{
					string text12 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text12 == "VND")
					{
						text3 = text3.Replace("BG", "BH");
					}
					sheetView.Cells[CellId_ct30[1], CellId_ct30[0]].Formula = text3;
				}
				if ((iRow == 74 || iRow == 81) && CellId_ct34[0] + 1 == iCol)
				{
					string text12 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text12 == "VND")
					{
						text4 = text4.Replace("BG", "BH");
					}
					sheetView.Cells[CellId_ct34[1], CellId_ct34[0]].Formula = text4;
				}
				switch (iRow)
				{
				case 73L:
				case 76L:
				case 77L:
				case 81L:
				case 83L:
					if (CellId_ct35[0] + 1 == iCol)
					{
						string text12 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
						if (text12 == "VND")
						{
							text5 = text5.Replace("BG", "BH");
						}
						sheetView.Cells[CellId_ct35[1], CellId_ct35[0]].Formula = text5;
					}
					break;
				}
				switch (iRow)
				{
				case 73L:
				case 76L:
				case 77L:
				case 81L:
				case 83L:
				case 84L:
				case 85L:
					if (CellId_ct35[0] + 1 == iCol)
					{
						string text12 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
						if (text12 == "VND")
						{
							text6 = text6.Replace("BG", "BH");
						}
						sheetView.Cells[CellId_ct36[1], CellId_ct36[0]].Formula = text6;
					}
					break;
				}
				switch (iRow)
				{
				case 75L:
				case 76L:
				case 77L:
				case 80L:
					if (CellId_ct35[0] + 1 == iCol)
					{
						string text12 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
						if (text12 == "VND")
						{
							text7 = text7.Replace("BG", "BH");
						}
						sheetView.Cells[CellId_ct37[1], CellId_ct37[0]].Formula = text7;
					}
					break;
				}
				switch (iRow)
				{
				case 75L:
				case 76L:
				case 77L:
				case 80L:
				case 83L:
					if (CellId_ct35[0] + 1 == iCol)
					{
						string text12 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
						if (text12 == "VND")
						{
							text9 = text9.Replace("BG", "BH");
						}
						sheetView.Cells[CellId_ct39[1], CellId_ct39[0]].Formula = text9;
					}
					break;
				}
				switch (iRow)
				{
				case 73L:
				case 75L:
				case 76L:
				case 77L:
				case 80L:
				case 81L:
				case 83L:
				case 85L:
				case 87L:
					if (CellId_ct35[0] + 1 == iCol)
					{
						string text12 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
						if (text12 == "VND")
						{
							text10 = text10.Replace("BG", "BH");
						}
						sheetView.Cells[CellId_ct40[1], CellId_ct40[0]].Formula = text10;
					}
					break;
				}
				switch (iRow)
				{
				case 75L:
				case 76L:
				case 77L:
				case 80L:
				case 83L:
				case 89L:
				case 90L:
					if (CellId_ct35[0] + 1 == iCol)
					{
						string text12 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
						if (text12 == "VND")
						{
							text8 = text8.Replace("BG", "BH");
						}
						sheetView.Cells[CellId_ct38[1], CellId_ct38[0]].Formula = text8;
					}
					break;
				}
				switch (iRow)
				{
				case 75L:
				case 76L:
				case 77L:
				case 80L:
				case 83L:
				case 88L:
				case 89L:
				case 90L:
				case 91L:
					if (CellId_ct35[0] + 1 == iCol)
					{
						string text12 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
						if (text12 == "VND")
						{
							text11 = text11.Replace("BG", "BH");
						}
						sheetView.Cells[CellId_ct42[1], CellId_ct42[0]].Formula = text11;
					}
					break;
				}
				switch (iRow)
				{
				case 75L:
				case 76L:
				case 77L:
				case 80L:
				case 83L:
				case 88L:
				case 89L:
				case 90L:
					if (CellId_ct34[0] + 1 == iCol)
					{
						tinhCt5(winFpSpread, ToKhaiInfo);
					}
					break;
				}
			}
			GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[0], sInfo, 0);
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
				if (idxPL1_ColY + 1 == iCol || idxPL1_ColAD + 1 == iCol)
				{
					sheetView.Cells[(int)iRow - 1, idxPL1_ColAJ].Formula = sheetView.Cells[(int)iRow - 1, idxPL1_ColAV].Formula;
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

	public void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (activeSheetIndex == 2)
			{
				SheetView sheetView = winFpSpread.Sheets[2];
				if (idxPL1_ColU + 1 == iCol)
				{
					sheetView.Cells[(int)iRow - 1, idxPL1_ColAB].Formula = sheetView.Cells[(int)iRow - 1, idxPL1_ColAV].Formula;
				}
			}
			GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], 1);
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
			SheetView sheetView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
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
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, CellId_E19);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E25);
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E29);
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E31);
			string message5 = HTKKFile.GetMessage(sheetView, CellId_E32);
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
			else if (flag && flag2)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], message3);
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], message3);
				num = 2;
			}
			string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			int[] array = new int[5] { 87, 88, 89, 90, 92 };
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
						HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message4);
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message5);
					}
					num = 2;
				}
			}
			int[] array4 = new int[12]
			{
				73, 74, 75, 76, 77, 78, 79, 83, 84, 85,
				86, 87
			};
			for (int i = 0; i < array4.Length; i++)
			{
				int[] array5 = GeneralFunctions.ArrayCellId("AN_" + array4[i]);
				HTKKFile.RemoveValidate(sheetView, array5[1], array5[0]);
				string text2 = sheetView.Cells[array5[1], array5[0]].Text;
				if (string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array5[1], array5[0], message);
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

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sheetView2 = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView2, CellId_E19);
			string message2 = HTKKFile.GetMessage(sheetView2, CellId_E13);
			string message3 = HTKKFile.GetMessage(sheetView2, GeneralFunctions.ArrayCellId("E_36"));
			string message4 = HTKKFile.GetMessage(sheetView2, CellId_E30);
			string message5 = HTKKFile.GetMessage(sheetView, CellId_E31);
			string message6 = HTKKFile.GetMessage(sheetView, CellId_E32);
			HTKKCMCTableInfo hTKKCMCTableInfo = ((hTKKCMCSheetInfo != null && hTKKCMCSheetInfo.Tables != null && hTKKCMCSheetInfo.Tables.Count > 0) ? hTKKCMCSheetInfo.Tables[0] : null);
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "pathU29");
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PLuc_01_DK/tongCong");
			bool flag = false;
			if (hTKKCMCTableInfo != null)
			{
				int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
				for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsValues.Count; num3++)
				{
					array[0] = sheetView2.Cells[num2 + num3, idxPL1_ColD].Text;
					array[1] = sheetView2.Cells[num2 + num3, idxPL1_ColL].Text;
					array[2] = sheetView2.Cells[num2 + num3, idxPL1_ColU].Text;
					array[4] = sheetView2.Cells[num2 + num3, idxPL1_ColAL].Text;
					string text = array[0] + array[1] + array[2] + array[4];
					if (string.IsNullOrEmpty(text) || text.Equals("0,0000"))
					{
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxPL1_ColD);
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxPL1_ColL);
						HTKKFile.RemoveValidateBackGround(sheetView2, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
						return;
					}
					bool flag2 = GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array);
					if (flag2)
					{
						flag = true;
					}
					if (!flag2)
					{
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxPL1_ColD);
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxPL1_ColL);
						HTKKFile.RemoveValidateBackGround(sheetView2, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
						continue;
					}
					HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxPL1_ColD);
					HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxPL1_ColL);
					if (!ValidateBase.CheckMaSoThue(array[0]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView2, num2 + num3, idxPL1_ColD, message2);
						num = 2;
					}
					if (string.IsNullOrEmpty(array[1]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView2, num2 + num3, idxPL1_ColL, message);
						num = 2;
					}
					string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					int[] array2 = GeneralFunctions.ArrayCellId("AB_" + (num2 + num3 + 1));
					int[] array3 = GeneralFunctions.ArrayCellId("AV_" + (num2 + num3 + 1));
					double value = ConvertHelpers.ConvertToDouble(sheetView2.Cells[array3[1], array3[0]].Value);
					double value2 = ConvertHelpers.ConvertToDouble(sheetView2.Cells[array2[1], array2[0]].Value);
					double num4 = 0.0;
					HTKKFile.RemoveValidateBackGround(sheetView2, array2[1], array2[0]);
					num4 = ((!(text2 != "VND")) ? ((Math.Abs(value) * 0.001 < 100000.0) ? (Math.Abs(value) * 0.001) : 100000.0) : ((Math.Abs(value) * 0.001 < 5.0) ? (Math.Abs(value) * 0.001) : 5.0));
					if (Math.Abs(value2) > Math.Abs(value) + num4 || Math.Abs(value2) < Math.Abs(value) - num4)
					{
						if (text2 != "VND")
						{
							HTKKFile.SetNodeBackGroundPink(sheetView2, array2[1], array2[0], message5);
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView2, array2[1], array2[0], message6);
						}
						num = 2;
					}
				}
				double num5 = Convert.ToDouble(sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
				double num6 = Convert.ToDouble(sheetView2.Cells[CellId_PL1_ct9[1], CellId_PL1_ct9[0]].Value);
				if (num5 != 100.0 && flag)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView2, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message3);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidateBackGround(sheetView2, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
				}
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

	public static void UpdateFormula_1(FpSpread winFpSpread)
	{
		try
		{
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
