using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkTndn02TT80Nckt294
{
	private int idxPL03_01GDLK_ND132 = 2;

	private int idxPL03_02GDLK_ND132 = 3;

	private int idxPL03_03GDLK_ND132 = 4;

	private int idxPL03_04GDLK_ND132 = 5;

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

	public void AutoBindData_2(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			int index = idxPL03_01GDLK_ND132;
			SheetView sheetView = winFpSpread.Sheets[index];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			int count = hTKKCMCSheetInfo.Tables[0].RowsValues.Count;
			int num = count + hTKKCMCSheetInfo.Tables[1].RowsValues.Count - 1;
			int num2 = num + hTKKCMCSheetInfo.Tables[2].RowsValues.Count - 1;
			int num3 = num2 + hTKKCMCSheetInfo.Tables[3].RowsValues.Count - 1;
			int num4 = num3 + hTKKCMCSheetInfo.Tables[4].RowsValues.Count - 1;
			int num5 = num4 + hTKKCMCSheetInfo.Tables[5].RowsValues.Count - 1;
			int num6 = num5 + hTKKCMCSheetInfo.Tables[6].RowsValues.Count - 1;
			int num7 = num6 + hTKKCMCSheetInfo.Tables[7].RowsValues.Count - 1;
			int num8 = num7 + hTKKCMCSheetInfo.Tables[8].RowsValues.Count - 1;
			int num9 = num8 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count - 1;
			int num10 = num9 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count - 1;
			int num11 = num10 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count - 1;
			List<int> list = new List<int>();
			list.Add(HTKKFile.ColumnLettersToIndex("AF"));
			if (Convert.ToInt32(PeriodInfo.Nam) >= 2024)
			{
				GeneralFunctions.LockUnLockColsInTable(sheetView, hTKKCMCSheetInfo.Tables[0], list, locked: false, deleteData: false);
			}
			else
			{
				GeneralFunctions.LockUnLockColsInTable(sheetView, hTKKCMCSheetInfo.Tables[0], list, locked: true, deleteData: true);
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value))
			{
				for (int i = 51 + count; i < sheetView.NonEmptyRowCount; i++)
				{
					if (sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("AS") - 1].Text == "x")
					{
						sheetView.Rows[i].Visible = false;
					}
				}
			}
			else
			{
				for (int i = 51 + count; i < sheetView.NonEmptyRowCount; i++)
				{
					if (sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("AS") - 1].Text == "x")
					{
						sheetView.Rows[i].Visible = true;
					}
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12a"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12a"));
				for (int i = 1; i < 9; i++)
				{
					for (int num12 = hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[i].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("J") - 1].Locked = true;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("Y") - 1].Formula = "";
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AB") - 1].Locked = true;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AF") - 1].Formula = "";
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AG") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AH") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AI") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AJ") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AK") - 1].Locked = false;
					}
				}
				bool flag = ConvertHelpers.ConvertToBool(sheetView.Cells[100 + num8, HTKKFile.ColumnLettersToIndex("U") - 1].Value);
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					for (int i = 106 + num8; i < 123 + num8; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					}
					for (int i = 106 + num8; i < 132 + num8; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
					}
					for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
				}
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value))
				{
					for (int i = 106 + num8; i < 123 + num8; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					}
					for (int i = 106 + num8; i < 132 + num8; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
					}
					for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
					}
				}
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					for (int i = 220 + num10; i < 234 + num10; i++)
					{
						if (i != 224 + num10 && i != 225 + num10 && i != 226 + num10 && i != 233 + num10)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						}
					}
					for (int i = 175 + num10; i < 234 + num10; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
					}
					for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					}
				}
				if (!ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					for (int i = 220 + num10; i < 234 + num10; i++)
					{
						if (i != 224 + num10 && i != 225 + num10 && i != 226 + num10 && i != 233 + num10)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
						}
					}
					for (int i = 175 + num10; i < 234 + num10; i++)
					{
						if (i == 175 + num10 || i == 176 + num10 || i == 190 + num10 || i == 191 + num10 || i == 206 + num10 || i == 211 + num10 || i == 219 + num10 || i == 225 + num10 || i == 226 + num10 || i == 227 + num10 || i == 233 + num10)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "AN" + (i + 1);
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					}
				}
			}
			else
			{
				sheetView.Cells[59 + count, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
				sheetView.Cells[59 + count, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12a"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12a"));
				for (int i = 1; i < 9; i++)
				{
					for (int num12 = hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[i].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("J") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = string.Format("J{0}-H{0}", (num12 + 1).ToString());
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("Y") - 1].Formula = string.Format("V{0}-R{0}", (num12 + 1).ToString());
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AB") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AF") - 1].Formula = string.Format("ABS(M{0})+ABS(Y{0})", (num12 + 1).ToString());
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AG") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AH") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AI") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AJ") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AK") - 1].Locked = false;
					}
				}
				for (int i = 106 + num8; i < 114 + num8; i++)
				{
					if (i != 109 + num8 && i != 111 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
					}
				}
				string pathCt6 = "PL_GDLK2-01/Muc_IV/IV_1/Dong_5/ct06";
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathCt6);
				for (int i = 106 + num8; i < 114 + num8; i++)
				{
					if (i == 109 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (i - 2).ToString()).Replace("@@@", i.ToString());
						if (i != hTKKCMCCellInfo3.CellTemplate.RowIndex - 1)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					else if (i == 111 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (i - 1).ToString()).Replace("@@@", i.ToString());
						if (i != hTKKCMCCellInfo3.CellTemplate.RowIndex - 1)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					else
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "M###+N###+R###".Replace("###", (i + 1).ToString());
						if (i != hTKKCMCCellInfo3.CellTemplate.RowIndex - 1)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
				}
				for (int i = 121 + num8; i < 133 + num8; i++)
				{
					if (i == 121 + num8 || i == 122 + num8 || i == 123 + num8 || i == 124 + num8 || i == 125 + num8 || i == 131 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "AN" + (i + 1);
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
				}
				for (int i = 220 + num10; i < 234 + num10; i++)
				{
					if (i != 224 + num10 && i != 225 + num10 && i != 226 + num10 && i != 233 + num10)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
					}
				}
				for (int i = 175 + num10; i < 234 + num10; i++)
				{
					if (i == 175 + num10 || i == 176 + num10 || i == 190 + num10 || i == 191 + num10 || i == 206 + num10 || i == 211 + num10 || i == 219 + num10 || i == 225 + num10 || i == 226 + num10 || i == 227 + num10 || i == 233 + num10)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "AN" + (i + 1);
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value))
			{
				for (int i = 106 + num8; i < 114 + num8; i++)
				{
					if (i != 109 + num8 && i != 111 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				for (int i = 126 + num8; i < 132 + num8; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
			{
				for (int i = 106 + num8; i < 114 + num8; i++)
				{
					if (i != 109 + num8 && i != 111 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else
			{
				for (int i = 106 + num8; i < 114 + num8; i++)
				{
					if (i != 109 + num8 && i != 111 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
				}
				for (int i = 126 + num8; i < 132 + num8; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
				for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[139 + num9, GeneralFunctions.ArrayCellId("X_141")[0]].Value))
			{
				for (int i = 145 + num9; i < 160 + num9; i++)
				{
					if (i != 147 + num9 && i != 150 + num9 && i != 156 + num9)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else
			{
				for (int i = 145 + num9; i < 160 + num9; i++)
				{
					if (i != 147 + num9 && i != 150 + num9 && i != 156 + num9)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[169 + num10, GeneralFunctions.ArrayCellId("X_171")[0]].Value))
			{
				for (int i = 220 + num10; i < 224 + num10; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int i = 228 + num10; i < 233 + num10; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
			{
				for (int i = 220 + num10; i < 224 + num10; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int i = 228 + num10; i < 233 + num10; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else
			{
				for (int i = 220 + num10; i < 224 + num10; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
				for (int i = 228 + num10; i < 233 + num10; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
				for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
			string text = "";
			string[] array = new string[11]
			{
				"H_", "J_", "M_", "R_", "V_", "Y_", "AF_", "AG_", "AH_", "AI_",
				"AJ_"
			};
			for (int i = 1; i < 9; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int num12 = 0; num12 < hTKKCMCTableInfo.RowsValues.Count; num12++)
				{
					for (int num13 = 0; num13 < array.Length; num13++)
					{
						text = text + array[num13] + (dataLocationTemplateRowStart + num12) + "###";
					}
				}
			}
			text = text.Substring(0, text.Length - 3);
			string donViTinhMa = PeriodInfo.DonViTinhMa;
			GeneralFunctions.changFormatCell(sheetView, text, donViTinhMa);
			text = "";
			string[] array2 = new string[11]
			{
				"H_", "J_", "M_", "R_", "V_", "Y_", "AF_", "AG_", "AH_", "AI_",
				"AJ_"
			};
			HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL_GDLK2-01/Muc_III/Dong_I/ct3");
			text = "H_" + hTKKCMCCellInfo4.CellTemplate.RowIndex;
			hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL_GDLK2-01/Muc_III/Dong_I/ct7");
			text = text + "###R_" + hTKKCMCCellInfo4.CellTemplate.RowIndex + "###";
			for (int num12 = 0; num12 < array2.Length; num12++)
			{
				text = text + array2[num12] + (hTKKCMCCellInfo4.CellTemplate.RowIndex + 1) + "###";
				text = text + array2[num12] + (hTKKCMCCellInfo4.CellTemplate.RowIndex + 2) + "###";
				text = text + array2[num12] + (hTKKCMCCellInfo4.CellTemplate.RowIndex + 3) + "###";
			}
			text = text.Substring(0, text.Length - 3);
			GeneralFunctions.changFormatCell(sheetView, text, donViTinhMa);
			hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL_GDLK2-01/Muc_III/Dong_II_1_1.2/ct3");
			text = "";
			for (int num12 = 0; num12 < array2.Length; num12++)
			{
				text = text + array2[num12] + hTKKCMCCellInfo4.CellTemplate.RowIndex + "###";
			}
			text = text.Substring(0, text.Length - 3);
			GeneralFunctions.changFormatCell(sheetView, text, donViTinhMa);
			hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL_GDLK2-01/Muc_III/Dong_II_2/ct3");
			text = "";
			for (int num12 = 0; num12 < array2.Length; num12++)
			{
				text = text + array2[num12] + hTKKCMCCellInfo4.CellTemplate.RowIndex + "###";
				text = text + array2[num12] + (hTKKCMCCellInfo4.CellTemplate.RowIndex + 1) + "###";
			}
			text = text.Substring(0, text.Length - 3);
			GeneralFunctions.changFormatCell(sheetView, text, donViTinhMa);
			hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL_GDLK2-01/Muc_III/Dong_II_2_2.2/ct3");
			text = "";
			for (int num12 = 0; num12 < array2.Length; num12++)
			{
				text = text + array2[num12] + hTKKCMCCellInfo4.CellTemplate.RowIndex + "###";
			}
			text = text.Substring(0, text.Length - 3);
			GeneralFunctions.changFormatCell(sheetView, text, donViTinhMa);
			hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL_GDLK2-01/Muc_III/Dong_II_2_2.3/ct3");
			text = "";
			for (int num12 = 0; num12 < array2.Length; num12++)
			{
				text = text + array2[num12] + hTKKCMCCellInfo4.CellTemplate.RowIndex + "###";
			}
			text = text.Substring(0, text.Length - 3);
			GeneralFunctions.changFormatCell(sheetView, text, donViTinhMa);
			hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL_GDLK2-01/Muc_III/Dong_II_2_2.4/ct3");
			text = "";
			for (int num12 = 0; num12 < array2.Length; num12++)
			{
				text = text + array2[num12] + hTKKCMCCellInfo4.CellTemplate.RowIndex + "###";
				text = text + array2[num12] + (hTKKCMCCellInfo4.CellTemplate.RowIndex + 1) + "###";
			}
			text = text.Substring(0, text.Length - 3);
			GeneralFunctions.changFormatCell(sheetView, text, donViTinhMa);
			hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.2/ct3");
			text = "";
			for (int num12 = 0; num12 < array2.Length; num12++)
			{
				text = text + array2[num12] + hTKKCMCCellInfo4.CellTemplate.RowIndex + "###";
			}
			text = text.Substring(0, text.Length - 3);
			GeneralFunctions.changFormatCell(sheetView, text, donViTinhMa);
			hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL_GDLK2-01/Muc_III/Dong_II_2_2.5/ct3");
			text = "";
			for (int num12 = 0; num12 < array2.Length; num12++)
			{
				text = text + array2[num12] + hTKKCMCCellInfo4.CellTemplate.RowIndex + "###";
			}
			text = text.Substring(0, text.Length - 3);
			GeneralFunctions.changFormatCell(sheetView, text, donViTinhMa);
			hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL_GDLK2-01/Muc_IV/IV_1/Dong_1/ct03");
			text = "";
			string[] array3 = new string[4] { "M_", "N_", "R_", "V_" };
			for (int num12 = 0; num12 < array3.Length; num12++)
			{
				for (int num14 = 0; num14 < 26; num14++)
				{
					if (num14 <= 7 || num14 == 16 || (num14 >= 19 && num14 <= 24))
					{
						text = text + array3[num12] + (hTKKCMCCellInfo4.CellTemplate.RowIndex + num14) + "###";
					}
					if (((num14 >= 8 && num14 <= 15) || num14 == 17) && array3[num12] == "V_")
					{
						text = text + array3[num12] + (hTKKCMCCellInfo4.CellTemplate.RowIndex + num14) + "###";
					}
					if (num14 == 25 && array3[num12] != "V_")
					{
						text = text + array3[num12] + (hTKKCMCCellInfo4.CellTemplate.RowIndex + num14) + "###";
					}
				}
			}
			text = text.Substring(0, text.Length - 3);
			GeneralFunctions.changFormatCell(sheetView, text, donViTinhMa);
			hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL_GDLK2-01/Muc_IV/IV_2/Dong_1/ct03");
			text = "";
			string[] array4 = new string[4] { "M_", "N_", "R_", "V_" };
			for (int num12 = 0; num12 < array4.Length; num12++)
			{
				for (int num14 = 0; num14 < 17; num14++)
				{
					text = text + array4[num12] + (hTKKCMCCellInfo4.CellTemplate.RowIndex + num14) + "###";
				}
			}
			text = text.Substring(0, text.Length - 3);
			GeneralFunctions.changFormatCell(sheetView, text, donViTinhMa);
			hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL_GDLK2-01/Muc_IV/IV_3/Dong_1/ct06");
			text = "";
			string[] array5 = new string[4] { "M_", "N_", "R_", "V_" };
			for (int num12 = 0; num12 < array5.Length; num12++)
			{
				for (int num14 = 0; num14 < 58; num14++)
				{
					if (array5[num12] == "V_" && num14 != 51)
					{
						text = text + array5[num12] + (hTKKCMCCellInfo4.CellTemplate.RowIndex + num14) + "###";
					}
					if (array5[num12] != "V_" && ((num14 >= 45 && num14 <= 48) || (num14 >= 52 && num14 <= 57)))
					{
						text = text + array5[num12] + (hTKKCMCCellInfo4.CellTemplate.RowIndex + num14) + "###";
					}
				}
			}
			text = text.Substring(0, text.Length - 3);
			GeneralFunctions.changFormatCell(sheetView, text, donViTinhMa);
			GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[idxPL03_01GDLK_ND132], ToKhaiInfo.SheetInfos[idxPL03_01GDLK_ND132], 1);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[idxPL03_04GDLK_ND132];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_04GDLK_ND132];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			string[] array = new string[9] { "D_", "G_", "J_", "M_", "O_", "Q_", "S_", "U_", "Y_" };
			string text = "";
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				for (int j = 0; j < array.Length; j++)
				{
					text = text + array[j] + (dataLocationTemplateRowStart + i) + "###";
				}
			}
			text = text.Substring(0, text.Length - 3);
			string donViTinhMa = PeriodInfo.DonViTinhMa;
			GeneralFunctions.changFormatCell(sView, text, donViTinhMa);
			GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[idxPL03_04GDLK_ND132], ToKhaiInfo.SheetInfos[idxPL03_04GDLK_ND132], 1);
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

	public void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxPL03_01GDLK_ND132];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_01GDLK_ND132];
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.1/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.1/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo11 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.1/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.1/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo13 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.2/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo14 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.2/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo15 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.2/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo16 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.2/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo17 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo18 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo19 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo20 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo21 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.1/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo22 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.1/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo23 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.1/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo24 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.1/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo25 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.2/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo26 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.2/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo27 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.2/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo28 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.2/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo29 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.3/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo30 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.3/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo31 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.3/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo32 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.3/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo33 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo34 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo35 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo36 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo37 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.1/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo38 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.1/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo39 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.1/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo40 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.1/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo41 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.2/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo42 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.2/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo43 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.2/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo44 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.2/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo45 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.5/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo46 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.5/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo47 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.5/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo48 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.5/ct12d"));
			if (iRow == hTKKCMCCellInfo5.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo17.CellTemplate.RowIndex)
			{
				if (iCol == hTKKCMCCellInfo5.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo6.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo7.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo8.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
				}
			}
			else if (iRow == hTKKCMCCellInfo9.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo13.CellTemplate.RowIndex)
			{
				if (iCol == hTKKCMCCellInfo9.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo9.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo13.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo10.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo10.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo14.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo11.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo11.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo15.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo12.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo12.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo16.CellTemplate.RowIndex}";
				}
			}
			else if (iRow == hTKKCMCCellInfo21.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo25.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo29.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo33.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo45.CellTemplate.RowIndex)
			{
				if (iCol == hTKKCMCCellInfo21.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo22.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo23.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo24.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
				}
			}
			else if (iRow == hTKKCMCCellInfo37.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo41.CellTemplate.RowIndex)
			{
				if (iCol == hTKKCMCCellInfo41.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo33.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo37.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo41.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo42.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo34.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo38.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo42.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo43.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo35.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo39.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo43.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo44.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo36.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo40.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo44.CellTemplate.RowIndex}";
				}
			}
			for (int num = 1; num < 9; num++)
			{
				if (iRow < hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart || iRow > hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count)
				{
					continue;
				}
				if (num == 1 || num == 2)
				{
					if (iCol == hTKKCMCCellInfo9.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo9.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo13.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo10.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo10.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo14.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo11.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo11.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo15.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo12.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo12.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo16.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AJ{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AJ{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
				}
				else if (num == 3 || num == 4 || num == 5 || num == 8)
				{
					if (iCol == hTKKCMCCellInfo21.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo22.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo23.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo24.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AJ{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AJ{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
				}
				else if (num == 6 || num == 7)
				{
					if (iCol == hTKKCMCCellInfo41.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo33.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo37.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo41.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo42.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo34.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo38.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo42.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo43.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo35.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo39.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo43.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo44.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo36.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo40.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo44.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AJ{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AJ{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
				}
			}
			string pathCt6 = "PL_GDLK2-01/Muc_IV/IV_1/Dong_5/ct06";
			HTKKCMCCellInfo hTKKCMCCellInfo49 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathCt6);
			if (iRow < hTKKCMCCellInfo49.CellTemplate.RowIndex && iRow >= hTKKCMCCellInfo49.CellTemplate.RowIndex - 11)
			{
				sheetView.Cells[hTKKCMCCellInfo49.CellTemplate.RowIndex - 1, hTKKCMCCellInfo49.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo49.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("AN") - 1].Value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
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

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			int index = idxPL03_01GDLK_ND132;
			winFpSpread.SuspendLayout();
			SheetView sheetView = winFpSpread.Sheets[index];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			int count = hTKKCMCSheetInfo.Tables[0].RowsValues.Count;
			int num = count + hTKKCMCSheetInfo.Tables[1].RowsValues.Count - 1;
			int num2 = num + hTKKCMCSheetInfo.Tables[2].RowsValues.Count - 1;
			int num3 = num2 + hTKKCMCSheetInfo.Tables[3].RowsValues.Count - 1;
			int num4 = num3 + hTKKCMCSheetInfo.Tables[4].RowsValues.Count - 1;
			int num5 = num4 + hTKKCMCSheetInfo.Tables[5].RowsValues.Count - 1;
			int num6 = num5 + hTKKCMCSheetInfo.Tables[6].RowsValues.Count - 1;
			int num7 = num6 + hTKKCMCSheetInfo.Tables[7].RowsValues.Count - 1;
			int num8 = num7 + hTKKCMCSheetInfo.Tables[8].RowsValues.Count - 1;
			int num9 = num8 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count - 1;
			int num10 = num9 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count - 1;
			int num11 = num10 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count - 1;
			if (idxRow == 38 + count && idxCol == GeneralFunctions.ArrayCellId("N_40")[0])
			{
				bool flag = ConvertHelpers.ConvertToBool(sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value);
				sheetView.Cells[37 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value = flag;
			}
			if ((idxRow == 42 + count || idxRow == 41 + count || idxRow == 40 + count || idxRow == 43 + count || idxRow == 44 + count || idxRow == 45 + count) && idxCol == GeneralFunctions.ArrayCellId("N_44")[0])
			{
				bool flag2 = ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value);
				bool flag3 = ConvertHelpers.ConvertToBool(sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value);
				bool flag4 = ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value);
				bool flag5 = ConvertHelpers.ConvertToBool(sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value);
				bool flag6 = ConvertHelpers.ConvertToBool(sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value);
				bool flag7 = ConvertHelpers.ConvertToBool(sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value);
				sheetView.Cells[39 + count, GeneralFunctions.ArrayCellId("N_41")[0]].Value = flag2 || flag3 || flag4 || flag5 || flag6 || flag7;
			}
			if (idxRow == 38 + count && idxCol == GeneralFunctions.ArrayCellId("N_40")[0])
			{
				AnMucIII_IV_ND132(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
			}
			if (idxRow == 38 + count && idxCol == GeneralFunctions.ArrayCellId("N_40")[0] && ConvertHelpers.ConvertToBool(sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value))
			{
				sheetView.Cells[39 + count, GeneralFunctions.ArrayCellId("N_41")[0]].Value = false;
				sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value = false;
				sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value = false;
				sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value = false;
				sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value = false;
				sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value = false;
				sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value = false;
			}
			if (idxRow == 40 + count && idxCol == GeneralFunctions.ArrayCellId("N_42")[0] && ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
			{
				sheetView.Cells[37 + count, GeneralFunctions.ArrayCellId("N_39")[0]].Value = false;
				sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value = false;
				sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value = false;
				sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value = false;
				sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value = false;
				sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value = false;
				sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value = false;
				AnMucIII_IV_ND132(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
			}
			if (idxRow == 41 + count && idxCol == GeneralFunctions.ArrayCellId("N_43")[0] && ConvertHelpers.ConvertToBool(sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value))
			{
				sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value = false;
				sheetView.Cells[37 + count, GeneralFunctions.ArrayCellId("N_39")[0]].Value = false;
				sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value = false;
				sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value = false;
				sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value = false;
				sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value = false;
				sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value = false;
				AnMucIII_IV_ND132(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
			}
			if (idxRow == 42 + count && idxCol == GeneralFunctions.ArrayCellId("N_44")[0])
			{
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value))
				{
					sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value = false;
					sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value = false;
					sheetView.Cells[37 + count, GeneralFunctions.ArrayCellId("N_39")[0]].Value = false;
					sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value = false;
					AnMucIII_IV_ND132(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
				}
				else
				{
					sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value = false;
					sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value = false;
					sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value = false;
				}
			}
			if ((idxRow == 43 + count || idxRow == 44 + count || idxRow == 45 + count) && idxCol == GeneralFunctions.ArrayCellId("N_44")[0] && (ConvertHelpers.ConvertToBool(sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value)))
			{
				sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value = false;
				sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value = false;
				sheetView.Cells[37 + count, GeneralFunctions.ArrayCellId("N_39")[0]].Value = false;
				sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value = false;
				sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value = true;
				AnMucIII_IV_ND132(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
			}
			if (idxRow >= 38 + count && idxRow <= 45 + count && idxCol == GeneralFunctions.ArrayCellId("N_44")[0])
			{
				XoaDLIII_IV_ND132(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count, ToKhaiInfo);
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12a"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12a"));
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AG") - 1].Formula = $"AG{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AG{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AH") - 1].Formula = $"AH{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AH{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AI") - 1].Formula = $"AI{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AI{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AJ") - 1].Formula = $"AJ{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AJ{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					for (int num12 = 1; num12 < 9; num12++)
					{
						for (int num13 = hTKKCMCSheetInfo.Tables[num12].Location.DataLocationTemplateRowStart - 1; num13 < hTKKCMCSheetInfo.Tables[num12].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[num12].RowsValues.Count; num13++)
						{
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("J") - 1].Locked = true;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("Y") - 1].Formula = "";
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AB") - 1].Locked = true;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AF") - 1].Formula = "";
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AG") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AH") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AI") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AJ") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AK") - 1].Locked = false;
						}
					}
					bool flag8 = ConvertHelpers.ConvertToBool(sheetView.Cells[100 + num8, HTKKFile.ColumnLettersToIndex("U") - 1].Value);
					if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) && idxRow == 40 + count)
					{
						for (int num12 = 106 + num8; num12 < 123 + num8; num12++)
						{
							if (num12 != 109 + num8 && num12 != 111 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "";
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							}
						}
						for (int num12 = 106 + num8; num12 < 132 + num8; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
						for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Value = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					if (ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value) && (idxRow == 42 + count || idxRow == 43 + count || idxRow == 44 + count || idxRow == 45 + count))
					{
						for (int num12 = 106 + num8; num12 < 123 + num8; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						}
						for (int num12 = 106 + num8; num12 < 132 + num8; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
						for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
					}
					if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) && idxRow == 40 + count)
					{
						for (int num12 = 220 + num10; num12 < 234 + num10; num12++)
						{
							if (num12 != 224 + num10 && num12 != 225 + num10 && num12 != 226 + num10 && num12 != 233 + num10)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							}
						}
						for (int num12 = 175 + num10; num12 < 234 + num10; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
						for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Value = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						}
					}
					if (!ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
					{
						for (int num12 = 220 + num10; num12 < 234 + num10; num12++)
						{
							if (num12 != 224 + num10 && num12 != 225 + num10 && num12 != 226 + num10 && num12 != 233 + num10)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
							}
						}
						for (int num12 = 175 + num10; num12 < 234 + num10; num12++)
						{
							if (num12 == 175 + num10 || num12 == 176 + num10 || num12 == 190 + num10 || num12 == 191 + num10 || num12 == 206 + num10 || num12 == 211 + num10 || num12 == 219 + num10 || num12 == 225 + num10 || num12 == 226 + num10 || num12 == 227 + num10 || num12 == 233 + num10)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "AN" + (num12 + 1);
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
							}
						}
						for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						}
					}
				}
				else
				{
					sheetView.Cells[59 + count, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
					sheetView.Cells[59 + count, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12a"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12a"));
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AG") - 1].Formula = $"AG{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AG{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AH") - 1].Formula = $"AH{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AH{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AI") - 1].Formula = $"AI{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AI{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AJ") - 1].Formula = $"AJ{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AJ{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					for (int num12 = 1; num12 < 9; num12++)
					{
						for (int num13 = hTKKCMCSheetInfo.Tables[num12].Location.DataLocationTemplateRowStart - 1; num13 < hTKKCMCSheetInfo.Tables[num12].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[num12].RowsValues.Count; num13++)
						{
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("J") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = string.Format("J{0}-H{0}", (num13 + 1).ToString());
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("Y") - 1].Formula = string.Format("V{0}-R{0}", (num13 + 1).ToString());
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AB") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AF") - 1].Formula = string.Format("ABS(M{0})+ABS(Y{0})", (num13 + 1).ToString());
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AG") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AH") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AI") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AJ") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AK") - 1].Locked = false;
						}
					}
					for (int num12 = 106 + num8; num12 < 114 + num8; num12++)
					{
						if (num12 != 109 + num8 && num12 != 111 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
						}
					}
					for (int num12 = 106 + num8; num12 < 114 + num8; num12++)
					{
						if (num12 == 109 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (num12 - 2).ToString()).Replace("@@@", num12.ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (num12 == 111 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (num12 - 1).ToString()).Replace("@@@", num12.ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "M###+N###+R###".Replace("###", (num12 + 1).ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					for (int num12 = 121 + num8; num12 < 133 + num8; num12++)
					{
						if (num12 == 121 + num8 || num12 == 122 + num8 || num12 == 123 + num8 || num12 == 124 + num8 || num12 == 125 + num8 || num12 == 131 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "AN" + (num12 + 1);
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
					for (int num12 = 220 + num10; num12 < 234 + num10; num12++)
					{
						if (num12 != 224 + num10 && num12 != 225 + num10 && num12 != 226 + num10 && num12 != 233 + num10)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
						}
					}
					for (int num12 = 175 + num10; num12 < 234 + num10; num12++)
					{
						if (num12 == 175 + num10 || num12 == 176 + num10 || num12 == 190 + num10 || num12 == 191 + num10 || num12 == 206 + num10 || num12 == 211 + num10 || num12 == 219 + num10 || num12 == 225 + num10 || num12 == 226 + num10 || num12 == 227 + num10 || num12 == 233 + num10)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "AN" + (num12 + 1);
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Value = "";
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					}
				}
			}
			if (idxRow == 100 + num8 && (idxCol == GeneralFunctions.ArrayCellId("X_102")[0] || idxCol == GeneralFunctions.ArrayCellId("U_102")[0]))
			{
				if (idxCol == GeneralFunctions.ArrayCellId("X_102")[0] && sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("U_102")[0]].Value = false;
				}
				if (idxCol == GeneralFunctions.ArrayCellId("U_102")[0] && sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("U_102")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value = false;
				}
			}
			if (idxRow == 139 + num9 && (idxCol == GeneralFunctions.ArrayCellId("X_141")[0] || idxCol == GeneralFunctions.ArrayCellId("U_141")[0]))
			{
				if (idxCol == GeneralFunctions.ArrayCellId("X_141")[0] && sheetView.Cells[139 + num9, GeneralFunctions.ArrayCellId("X_141")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[139 + num9, GeneralFunctions.ArrayCellId("U_141")[0]].Value = false;
				}
				if (idxCol == GeneralFunctions.ArrayCellId("U_141")[0] && sheetView.Cells[139 + num9, GeneralFunctions.ArrayCellId("U_141")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[139 + num9, GeneralFunctions.ArrayCellId("X_141")[0]].Value = false;
				}
			}
			if (idxRow == 169 + num10 && (idxCol == GeneralFunctions.ArrayCellId("X_171")[0] || idxCol == GeneralFunctions.ArrayCellId("U_171")[0]))
			{
				if (idxCol == GeneralFunctions.ArrayCellId("X_171")[0] && sheetView.Cells[169 + num10, GeneralFunctions.ArrayCellId("X_171")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[169 + num10, GeneralFunctions.ArrayCellId("U_171")[0]].Value = false;
				}
				if (idxCol == GeneralFunctions.ArrayCellId("U_171")[0] && sheetView.Cells[169 + num10, GeneralFunctions.ArrayCellId("U_171")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[169 + num10, GeneralFunctions.ArrayCellId("X_171")[0]].Value = false;
				}
			}
			if (idxRow == 100 + num8 && (idxCol == GeneralFunctions.ArrayCellId("X_102")[0] || idxCol == GeneralFunctions.ArrayCellId("U_102")[0]))
			{
				if (idxCol == GeneralFunctions.ArrayCellId("X_102")[0] && sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("U_102")[0]].Value = false;
				}
				if (idxCol == GeneralFunctions.ArrayCellId("U_102")[0] && sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("U_102")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value = false;
				}
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value))
				{
					for (int num12 = 106 + num8; num12 < 114 + num8; num12++)
					{
						if (num12 != 109 + num8 && num12 != 111 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						}
					}
					for (int num12 = 126 + num8; num12 < 132 + num8; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				else if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					for (int num12 = 106 + num8; num12 < 114 + num8; num12++)
					{
						if (num12 != 109 + num8 && num12 != 111 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				else
				{
					for (int num12 = 106 + num8; num12 < 114 + num8; num12++)
					{
						if (num12 != 109 + num8 && num12 != 111 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						}
					}
					for (int num12 = 126 + num8; num12 < 132 + num8; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
				}
			}
			if (idxRow == 139 + num9 && (idxCol == GeneralFunctions.ArrayCellId("X_141")[0] || idxCol == GeneralFunctions.ArrayCellId("U_141")[0]))
			{
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[139 + num9, GeneralFunctions.ArrayCellId("X_141")[0]].Value))
				{
					for (int num12 = 145 + num9; num12 < 160 + num9; num12++)
					{
						if (num12 != 147 + num9 && num12 != 150 + num9 && num12 != 156 + num9)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				else
				{
					for (int num12 = 145 + num9; num12 < 160 + num9; num12++)
					{
						if (num12 != 147 + num9 && num12 != 150 + num9 && num12 != 156 + num9)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
				}
			}
			if (idxRow != 169 + num10 || (idxCol != GeneralFunctions.ArrayCellId("X_171")[0] && idxCol != GeneralFunctions.ArrayCellId("U_171")[0]))
			{
				return;
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[169 + num10, GeneralFunctions.ArrayCellId("X_171")[0]].Value))
			{
				for (int num12 = 220 + num10; num12 < 225 + num10; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int num12 = 228 + num10; num12 < 233 + num10; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
			{
				for (int num12 = 220 + num10; num12 < 224 + num10; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int num12 = 228 + num10; num12 < 233 + num10; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else
			{
				for (int num12 = 220 + num10; num12 < 224 + num10; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
				for (int num12 = 228 + num10; num12 < 233 + num10; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
				for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		finally
		{
			winFpSpread.ResumeLayout();
		}
	}

	private static void AnMucIII_IV_ND132(int idxRow, int idxCol, SheetView sView, int rowCountTable1)
	{
		try
		{
			if (ConvertHelpers.ConvertToBool(sView.Cells[idxRow, idxCol].Value))
			{
				for (int i = 51 + rowCountTable1; i < sView.NonEmptyRowCount; i++)
				{
					if (!(sView.Cells[i, HTKKFile.ColumnLettersToIndex("AS") - 1].Text == "x"))
					{
						continue;
					}
					for (int j = 0; j < sView.NonEmptyColumnCount; j++)
					{
						Cell cell = sView.Cells[i, j];
						if (!cell.Locked && cell.CellType != null)
						{
							if (cell.CellType.ToString() == "NumberCellType")
							{
								cell.Value = 0;
							}
							else if (cell.CellType.ToString() == "CheckBoxCellType")
							{
								cell.Value = false;
							}
							else
							{
								cell.Value = "";
							}
						}
					}
					sView.Rows[i].Visible = false;
				}
				return;
			}
			for (int i = 51 + rowCountTable1; i < sView.NonEmptyRowCount; i++)
			{
				if (sView.Cells[i, HTKKFile.ColumnLettersToIndex("AS") - 1].Text == "x")
				{
					sView.Rows[i].Visible = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private static void XoaDLIII_IV_ND132(int idxRow, int idxCol, SheetView sView, int rowCountTable1, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			for (int i = 51 + rowCountTable1; i < sView.NonEmptyRowCount; i++)
			{
				if (!(sView.Cells[i, HTKKFile.ColumnLettersToIndex("AS") - 1].Text == "x"))
				{
					continue;
				}
				for (int j = 0; j < sView.NonEmptyColumnCount; j++)
				{
					Cell cell = sView.Cells[i, j];
					if (cell.Locked && !cell.Note.IsNulOrEmpty())
					{
						HTKKFile.RemoveValidateBackGround(sView, i, j);
					}
					else if (!cell.Locked && !cell.Note.IsNulOrEmpty())
					{
						HTKKFile.RemoveValidate(sView, i, j);
					}
					if (!cell.Locked && cell.CellType != null)
					{
						if (cell.CellType.ToString() == "NumberCellType")
						{
							cell.Value = 0;
						}
						else if (cell.CellType.ToString() == "CheckBoxCellType")
						{
							cell.Value = false;
						}
						else
						{
							cell.Value = "";
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
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
			if (ToKhaiInfo.SheetInfos[2] != null)
			{
				checkPL3_1_ND132(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
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

	private void checkPL3_1_ND132(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int num = idxPL03_01GDLK_ND132;
			if (winFpSpread.Sheets[num] == null || !winFpSpread.Sheets[num].Visible)
			{
				return;
			}
			int value = 0;
			levelWarning.TryGetValue(num, out value);
			bool flag = false;
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[num];
			int count = hTKKCMCSheetInfo.Tables[0].RowsValues.Count;
			int num2 = count + hTKKCMCSheetInfo.Tables[1].RowsValues.Count - 1;
			int num3 = num2 + hTKKCMCSheetInfo.Tables[2].RowsValues.Count - 1;
			int num4 = num3 + hTKKCMCSheetInfo.Tables[3].RowsValues.Count - 1;
			int num5 = num4 + hTKKCMCSheetInfo.Tables[4].RowsValues.Count - 1;
			int num6 = num5 + hTKKCMCSheetInfo.Tables[5].RowsValues.Count - 1;
			int num7 = num6 + hTKKCMCSheetInfo.Tables[6].RowsValues.Count - 1;
			int num8 = num7 + hTKKCMCSheetInfo.Tables[7].RowsValues.Count - 1;
			int num9 = num8 + hTKKCMCSheetInfo.Tables[8].RowsValues.Count - 1;
			int num10 = num9 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count - 1;
			int num11 = num10 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count - 1;
			int num12 = num11 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count - 1;
			if (!sheetView.Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[num];
			int num13 = ToKhaiInfo.SheetInfos[num].Tables[0].RowsValues.Count - 1;
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables[0];
			int num14 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("O") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("D") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("I") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("O") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("U") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("V") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("W") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("X") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("Y") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("Z") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AA") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AC") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AD") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AE") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AF") - 1);
				string[] array = hTKKCMCTableInfo.RowsValues[i].RowValue.ToArray();
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					value = 2;
				}
				if (string.IsNullOrEmpty(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("O") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					value = 2;
				}
				if (string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("I") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					value = 2;
				}
				else
				{
					if (string.IsNullOrEmpty(array[13]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("I") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_40")));
						value = 2;
					}
					if (array[1] == "Việt Nam" && !string.IsNullOrEmpty(array[2]) && !ValidateBase.CheckMaSoThue(array[2]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("O") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_41")));
						value = 2;
					}
				}
				if (!ConvertHelpers.ConvertToBool(array[3]) && !ConvertHelpers.ConvertToBool(array[4]) && !ConvertHelpers.ConvertToBool(array[5]) && !ConvertHelpers.ConvertToBool(array[6]) && !ConvertHelpers.ConvertToBool(array[7]) && !ConvertHelpers.ConvertToBool(array[8]) && !ConvertHelpers.ConvertToBool(array[9]) && !ConvertHelpers.ConvertToBool(array[10]) && !ConvertHelpers.ConvertToBool(array[11]) && !ConvertHelpers.ConvertToBool(array[12]) && !ConvertHelpers.ConvertToBool(array[13]) && !ConvertHelpers.ConvertToBool(array[14]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("W") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("X") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("Y") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AA") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AC") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AE") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AF") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					value = 2;
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[42 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1].Value))
			{
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[43 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[44 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[45 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1].Value))
				{
					HTKKFile.RemoveValidate(sheetView, 43 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
					HTKKFile.RemoveValidate(sheetView, 44 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
					HTKKFile.RemoveValidate(sheetView, 45 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 43 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					HTKKFile.SetNodeBackGroundPink(sheetView, 44 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					HTKKFile.SetNodeBackGroundPink(sheetView, 45 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					value = 2;
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, 43 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
				HTKKFile.RemoveValidate(sheetView, 44 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
				HTKKFile.RemoveValidate(sheetView, 45 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[num].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_II/ct1"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[num].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_II/ct2a"));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[num].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_II/ct2c"));
			if (!ConvertHelpers.ConvertToBool(hTKKCMCCellInfo.Value))
			{
				for (int i = 1; i < 9; i++)
				{
					for (int num15 = hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1; num15 < hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[i].RowsValues.Count; num15++)
					{
						string[] array = hTKKCMCSheetInfo.Tables[i].RowsValues[num15 - hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + 1].RowValue.ToArray();
						if (!GeneralFunctions.HaveToValidateRow(hTKKCMCSheetInfo.Tables[i].RowsInfo, array))
						{
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("AG") - 1);
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("AH") - 1);
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("AI") - 1);
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("AJ") - 1);
							continue;
						}
						if (array[3].ToIntOrDefault() < 0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num15, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_42")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("M") - 1);
						}
						if (array[7].ToIntOrDefault() > 0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num15, HTKKFile.ColumnLettersToIndex("Y") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("Y") - 1);
						}
						if (array[1].ToIntOrDefault() > 0 && array[4].IsNullOrEmpty())
						{
							if (!hTKKCMCCellInfo2.Value.ToBoolOrDefault() && !hTKKCMCCellInfo3.Value.ToBoolOrDefault())
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num15, HTKKFile.ColumnLettersToIndex("N") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
								value = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("N") - 1);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("N") - 1);
						}
						if (array[5].ToIntOrDefault() > 0 && array[8].IsNullOrEmpty())
						{
							if (!hTKKCMCCellInfo2.Value.ToBoolOrDefault() && !hTKKCMCCellInfo3.Value.ToBoolOrDefault())
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num15, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
								value = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("AB") - 1);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("AB") - 1);
						}
					}
				}
				bool flag2 = hTKKCMCSheetInfo.StaticCells[3].Value.ToBoolOrDefault();
				bool flag3 = hTKKCMCSheetInfo.StaticCells[4].Value.ToBoolOrDefault();
				bool flag4 = hTKKCMCSheetInfo.StaticCells[5].Value.ToBoolOrDefault();
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_1/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_1/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_2/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_4/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_6/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_7/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_9/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo11 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/Dong_1/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/Dong_2/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo13 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/Dong_1/ct06"));
				HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1, testColor: true);
				if (flag2)
				{
					value = CheckTongCot6MucIVP1P3_ND132(value, sheetView, hTKKCMCSheetInfo);
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo12.CellTemplate.RowIndex - 1, hTKKCMCCellInfo12.CellTemplate.ColIndex - 1, testColor: true);
				}
				if (flag3)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo14 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/co"));
					HTKKCMCCellInfo hTKKCMCCellInfo15 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/khong"));
					HTKKCMCCellInfo hTKKCMCCellInfo16 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_2/co"));
					HTKKCMCCellInfo hTKKCMCCellInfo17 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_2/khong"));
					HTKKCMCCellInfo hTKKCMCCellInfo18 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/co"));
					HTKKCMCCellInfo hTKKCMCCellInfo19 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/khong"));
					if (!hTKKCMCCellInfo14.Value.ToBoolOrDefault() && !hTKKCMCCellInfo15.Value.ToBoolOrDefault() && !hTKKCMCCellInfo16.Value.ToBoolOrDefault() && !hTKKCMCCellInfo17.Value.ToBoolOrDefault() && !hTKKCMCCellInfo18.Value.ToBoolOrDefault() && !hTKKCMCCellInfo19.Value.ToBoolOrDefault())
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo14.CellTemplate.RowIndex - 1, hTKKCMCCellInfo14.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo16.CellTemplate.RowIndex - 1, hTKKCMCCellInfo16.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, hTKKCMCCellInfo17.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo18.CellTemplate.RowIndex - 1, hTKKCMCCellInfo18.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo19.CellTemplate.RowIndex - 1, hTKKCMCCellInfo19.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo14.CellTemplate.RowIndex - 1, hTKKCMCCellInfo14.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo16.CellTemplate.RowIndex - 1, hTKKCMCCellInfo16.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, hTKKCMCCellInfo17.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo18.CellTemplate.RowIndex - 1, hTKKCMCCellInfo18.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo19.CellTemplate.RowIndex - 1, hTKKCMCCellInfo19.CellTemplate.ColIndex - 1, testColor: true);
					}
				}
				for (int i = 9; i < 12; i++)
				{
					for (int num15 = hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1; num15 < hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[i].RowsValues.Count; num15++)
					{
						string[] array = hTKKCMCSheetInfo.Tables[i].RowsValues[num15 - hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + 1].RowValue.ToArray();
						if (!GeneralFunctions.HaveToValidateRow(hTKKCMCSheetInfo.Tables[i].RowsInfo, array) && (!flag4 || i != 9))
						{
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("D") - 1);
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("V") - 1);
						}
						else if (string.IsNullOrEmpty(sheetView.Cells[num15, HTKKFile.ColumnLettersToIndex("D") - 1].Text) && !sheetView.Cells[num15, HTKKFile.ColumnLettersToIndex("D") - 1].Locked)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num15, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("D") - 1);
						}
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[num].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[num].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[num].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private static int CheckTongCot6MucIVP1P3_ND132(int iwarningColor, SheetView sView, HTKKCMCSheetInfo sheetInfo)
	{
		try
		{
			HTKKCMCCellInfo hTKKCMCCellInfo = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_1/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/Dong_1/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_1/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo4 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_4/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo5 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_6/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo6 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_7/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo7 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_9/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo8 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/Dong_1/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo9 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/Dong_2/ct06"));
			bool flag = false;
			if (sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m)
			{
				flag = true;
			}
			if (!flag)
			{
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_44")));
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_44")));
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_44")));
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_44")));
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_44")));
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_44")));
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_44")));
				iwarningColor = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, testColor: true);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, testColor: true);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, testColor: true);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1, testColor: true);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1, testColor: true);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1, testColor: true);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1, testColor: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return iwarningColor;
	}
}
