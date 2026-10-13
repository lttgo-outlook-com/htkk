using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkTk01Nd373Nckt293
{
	private int[] cellTenLoaiTien = GeneralFunctions.ArrayCellId("P_65");

	private int[] cellMaLoaiTien = GeneralFunctions.ArrayCellId("BG_65");

	private int colGiaTri = HTKKFile.ColumnLettersToIndex("AV");

	private int colCongThucUsd = HTKKFile.ColumnLettersToIndex("BI");

	private int colCongThucVnd = HTKKFile.ColumnLettersToIndex("BJ");

	private int[] ct01 = GeneralFunctions.ArrayCellId("BQ_14");

	private int[] ct02 = GeneralFunctions.ArrayCellId("BR_14");

	private int[] ct03 = GeneralFunctions.ArrayCellId("BS_14");

	private int[] ct05 = GeneralFunctions.ArrayCellId("AF_33");

	private int[] cellVsp = GeneralFunctions.ArrayCellId("Q_67");

	private int[] cellPsc = GeneralFunctions.ArrayCellId("Q_69");

	private int[] cell33 = GeneralFunctions.ArrayCellId("AN_87");

	private int[] ct08 = GeneralFunctions.ArrayCellId("AF_38");

	private string cellId_ct20_donvi = "Q_59";

	private string cellId_ct20_daily = "Y_59";

	private string cellId_E30 = "E_30";

	private string cellId_E29 = "E_29";

	private string USD = "USD";

	private string VND = "VND";

	private int[] CellId_ct20_donvi => GeneralFunctions.ArrayCellId(cellId_ct20_donvi);

	private int[] CellId_ct20_daily => GeneralFunctions.ArrayCellId(cellId_ct20_daily);

	private int[] CellId_E30 => GeneralFunctions.ArrayCellId(cellId_E30);

	private int[] CellId_E29 => GeneralFunctions.ArrayCellId(cellId_E29);

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			if (PeriodInfo.LanXuatBan.IndexOf("---") >= 0)
			{
				PeriodInfo.LanXuatBan = PeriodInfo.LanXuatBan.Replace("---", "/");
			}
			if (PeriodInfo.LanXuatBan.IndexOf("###") >= 0)
			{
				PeriodInfo.LanXuatBan = PeriodInfo.LanXuatBan.Replace("###", "\\");
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			sheetView.Cells[cellMaLoaiTien[1], cellMaLoaiTien[0]].Text = PeriodInfo.DonViTinhMa;
			sheetView.Cells[cellTenLoaiTien[1], cellTenLoaiTien[0]].Text = PeriodInfo.DonViTinhTen;
			sheetView.Cells[ct05[1], ct05[0]].Value = PeriodInfo.LanXuatBan;
			ThayDoiDonViTien_Tk(sheetView, hTKKCMCSheetInfo);
			sheetView.Cells[ct01[1], ct01[0]].Value = PeriodInfo.TieuMucHachToanVal == "1";
			sheetView.Cells[ct02[1], ct02[0]].Value = PeriodInfo.TieuMucHachToanVal == "3";
			sheetView.Cells[ct03[1], ct03[0]].Value = PeriodInfo.TieuMucHachToanVal == "2";
			if (PeriodInfo.LoaiToKhai == "B")
			{
				sheetView.Cells[cellVsp[1], cellVsp[0]].Locked = true;
				sheetView.Cells[cellPsc[1], cellPsc[0]].Locked = true;
			}
			LockUnLockTheoVspPsc(sheetView);
			if (string.IsNullOrEmpty(TTinChung.tenDLyThue) || string.IsNullOrEmpty(TTinChung.mstDLyThue))
			{
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0], locked: true, "", false);
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0], locked: true, "", false);
			}
			else if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				sheetView.Cells[CellId_ct20_daily[1], CellId_ct20_daily[0]].Value = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[1];
			ThayDoiDonViTien_Pl(winFpSpread, sInfo);
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
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			sheetView.AutoCalculation = true;
			string text = sheetView.Cells[cellMaLoaiTien[1], cellMaLoaiTien[0]].Text;
			if (iCol == HTKKFile.ColumnLettersToIndex("P"))
			{
				ThayDoiDonViTien_Tk(sheetView, sInfo);
				ThayDoiDonViTien_Pl(winFpSpread, ToKhaiInfo.SheetInfos[1]);
			}
			bool flag = Convert.ToBoolean(sheetView.Cells[cellPsc[1], cellPsc[0]].Value);
			if (iRow >= 92 && iRow <= 93 && flag)
			{
				if (text == USD)
				{
					sheetView.Cells[93, colGiaTri - 1].Value = sheetView.Cells[93, colCongThucUsd - 1].Value;
				}
				else
				{
					sheetView.Cells[93, colGiaTri - 1].Value = sheetView.Cells[93, colCongThucVnd - 1].Value;
				}
			}
			if (flag && (iRow == 92 || (iRow >= 96 && iRow <= 101)))
			{
				if (text == USD)
				{
					sheetView.Cells[101, colGiaTri - 1].Value = sheetView.Cells[101, colCongThucUsd - 1].Value;
				}
				else
				{
					sheetView.Cells[101, colGiaTri - 1].Value = sheetView.Cells[101, colCongThucVnd - 1].Value;
				}
			}
			if (flag && (iRow == 92 || (iRow >= 96 && iRow <= 106)))
			{
				if (text == USD)
				{
					sheetView.Cells[106, colGiaTri - 1].Value = sheetView.Cells[106, colCongThucUsd - 1].Value;
				}
				else
				{
					sheetView.Cells[106, colGiaTri - 1].Value = sheetView.Cells[106, colCongThucVnd - 1].Value;
				}
			}
			if (iRow >= 86 && iRow <= 88)
			{
				if (text == USD)
				{
					sheetView.Cells[111, colGiaTri - 1].Value = sheetView.Cells[111, colCongThucUsd - 1].Value;
				}
				else
				{
					sheetView.Cells[111, colGiaTri - 1].Value = sheetView.Cells[111, colCongThucVnd - 1].Value;
				}
			}
			if (iRow == 86 || iRow == 88)
			{
				if (text == USD)
				{
					sheetView.Cells[113, colGiaTri - 1].Value = sheetView.Cells[113, colCongThucUsd - 1].Value;
				}
				else
				{
					sheetView.Cells[113, colGiaTri - 1].Value = sheetView.Cells[113, colCongThucVnd - 1].Value;
				}
			}
			if (iRow >= 86 && iRow <= 88)
			{
				if (text == USD)
				{
					sheetView.Cells[114, colGiaTri - 1].Value = sheetView.Cells[114, colCongThucUsd - 1].Value;
				}
				else
				{
					sheetView.Cells[114, colGiaTri - 1].Value = sheetView.Cells[114, colCongThucVnd - 1].Value;
				}
			}
			if ((iRow >= 86 && iRow <= 88) || (iRow >= 114 && iRow <= 115))
			{
				if (text == USD)
				{
					sheetView.Cells[112, colGiaTri - 1].Value = sheetView.Cells[112, colCongThucUsd - 1].Value;
				}
				else
				{
					sheetView.Cells[112, colGiaTri - 1].Value = sheetView.Cells[112, colCongThucVnd - 1].Value;
				}
			}
			if ((iRow >= 86 && iRow <= 88) || (iRow >= 112 && iRow <= 115))
			{
				if (text == USD)
				{
					sheetView.Cells[107, colGiaTri - 1].Value = sheetView.Cells[107, colCongThucUsd - 1].Value;
				}
				else
				{
					sheetView.Cells[107, colGiaTri - 1].Value = sheetView.Cells[107, colCongThucVnd - 1].Value;
				}
			}
			if (iRow >= 92 && iRow <= 94)
			{
				suaCongThucCt5Pl(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], -1);
			}
			if ((iRow >= 96 && iRow <= 107) || iRow == 92)
			{
				suaCongThucCt7Pl(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], -1);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sViewPl = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo sInfoPl = ToKhaiInfo.SheetInfos[1];
			int num = HTKKFile.ColumnLettersToIndex("U");
			int num2 = HTKKFile.ColumnLettersToIndex("Y");
			int num3 = HTKKFile.ColumnLettersToIndex("AK");
			int num4 = HTKKFile.ColumnLettersToIndex("AQ");
			int num5 = HTKKFile.ColumnLettersToIndex("BQ");
			int num6 = HTKKFile.ColumnLettersToIndex("BR");
			if (iCol == num)
			{
				suaCongThucCt5Pl(sViewPl, sInfoPl, (int)iRow);
			}
			if (iCol == num3)
			{
				suaCongThucCt7Pl(sViewPl, sInfoPl, (int)iRow);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void suaCongThucCt5Pl(SheetView sViewPl, HTKKCMCSheetInfo sInfoPl, int iRowPl)
	{
		if (sInfoPl == null)
		{
			return;
		}
		int num = HTKKFile.ColumnLettersToIndex("Y");
		int num2 = HTKKFile.ColumnLettersToIndex("BQ");
		if (iRowPl >= 0)
		{
			sViewPl.Cells[iRowPl - 1, num - 1].Value = sViewPl.Cells[iRowPl - 1, num2 - 1].Value;
			return;
		}
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfoPl.Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			sViewPl.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Value = sViewPl.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Value;
		}
	}

	public void suaCongThucCt7Pl(SheetView sViewPl, HTKKCMCSheetInfo sInfoPl, int iRowPl)
	{
		if (sInfoPl == null)
		{
			return;
		}
		int num = HTKKFile.ColumnLettersToIndex("AQ");
		int num2 = HTKKFile.ColumnLettersToIndex("BR");
		if (iRowPl >= 0)
		{
			sViewPl.Cells[iRowPl - 1, num - 1].Value = sViewPl.Cells[iRowPl - 1, num2 - 1].Value;
			return;
		}
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfoPl.Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			sViewPl.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Value = sViewPl.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Value;
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
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			int colClick = winFpSpread.ActiveSheet.ActiveColumnIndex + 1;
			int num = winFpSpread.ActiveSheet.ActiveRowIndex + 1;
			GeneralFunctions.resetOtherCheckboxStatic(sheetView, sInfo, cellVsp[1] + 1, cellPsc[1] + 1, num, colClick);
			GeneralFunctions.resetOtherCheckboxStatic(sheetView, sInfo, ct08[1] + 1, ct08[1] + 1, num, colClick);
			if (idxRow == CellId_ct20_daily[1] && idxCol == CellId_ct20_daily[0])
			{
				if (Convert.ToBoolean(sheetView.Cells[CellId_ct20_daily[1], CellId_ct20_daily[0]].Value))
				{
					sheetView.Cells[CellId_ct20_donvi[1], CellId_ct20_donvi[0]].Value = false;
				}
			}
			else if (idxRow == CellId_ct20_donvi[1] && idxCol == CellId_ct20_donvi[0])
			{
				if (Convert.ToBoolean(sheetView.Cells[CellId_ct20_donvi[1], CellId_ct20_donvi[0]].Value))
				{
					sheetView.Cells[CellId_ct20_daily[1], CellId_ct20_daily[0]].Value = false;
				}
			}
			else if (num >= 67 && num <= 69)
			{
				LockUnLockTheoVspPsc(sheetView);
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink)
			{
				return false;
			}
			if (winFpSpread.TabStrip[1].BackColor == Color.Pink)
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
			return false;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, CellId_E30);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E29);
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_ct20_donvi[1], CellId_ct20_donvi[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct20_daily[1], CellId_ct20_daily[0]].Value);
			if (!HTKKFormater.isNullString(TTinChung.mstDLyThue))
			{
				if (flag && flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0], message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0], message2);
					num = 2;
				}
				else if (!flag && !flag2)
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
			if (Convert.ToBoolean(sheetView.Cells[cellPsc[1], cellPsc[0]].Value) && ToKhaiInfo.SheetInfos[1] == null)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, cellPsc[1], cellPsc[0], "Bắt buộc kê khai tờ PL 01-1/TK-DK");
				num = 2;
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
		}
	}

	public void ThayDoiDonViTien_Tk(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		string text = sView.Cells[cellMaLoaiTien[1], cellMaLoaiTien[0]].Text;
		List<string> list = new List<string>();
		list.Add("AV_92");
		list.Add("AV_94");
		list.Add("AV_95");
		list.Add("AV_100");
		list.Add("AV_102");
		list.Add("AV_106");
		list.Add("AV_107");
		list.Add("AV_108");
		list.Add("AV_112");
		list.Add("AV_113");
		list.Add("AV_114");
		list.Add("AV_115");
		foreach (string item in list)
		{
			int[] array = GeneralFunctions.ArrayCellId(item);
			NumberCellType numberCellType = (NumberCellType)sView.Cells[array[1], array[0]].CellType;
			numberCellType.DecimalPlaces = ((text == USD) ? 2 : 0);
			sView.Cells[array[1], array[0]].CellType = numberCellType;
		}
	}

	public void ThayDoiDonViTien_Pl(FpSpread winFpSpread, HTKKCMCSheetInfo sInfo)
	{
		if (sInfo == null)
		{
			return;
		}
		SheetView sheetView = winFpSpread.Sheets[0];
		SheetView sheetView2 = winFpSpread.Sheets[1];
		string text = sheetView.Cells[cellMaLoaiTien[1], cellMaLoaiTien[0]].Text;
		List<string> list = new List<string>();
		list.Add("T_69");
		list.Add("W_71");
		foreach (string item in list)
		{
			int[] array = GeneralFunctions.ArrayCellId(item);
			NumberCellType numberCellType = (NumberCellType)sheetView2.Cells[array[1], array[0]].CellType;
			numberCellType.DecimalPlaces = ((text == USD) ? 2 : 0);
			sheetView2.Cells[array[1], array[0]].CellType = numberCellType;
		}
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "PLuc_01_DK/tongCongct5");
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "PLuc_01_DK/tongCongct7");
		NumberCellType numberCellType2 = (NumberCellType)sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].CellType;
		numberCellType2.DecimalPlaces = ((text == USD) ? 2 : 0);
		sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].CellType = numberCellType2;
		NumberCellType numberCellType3 = (NumberCellType)sheetView2.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].CellType;
		numberCellType3.DecimalPlaces = ((text == USD) ? 2 : 0);
		sheetView2.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].CellType = numberCellType3;
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[0];
		int num = HTKKFile.ColumnLettersToIndex("Y");
		int num2 = HTKKFile.ColumnLettersToIndex("AQ");
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsValues.Count; num3++)
		{
			NumberCellType numberCellType = (NumberCellType)sheetView2.Cells[dataLocationTemplateRowStart + num3 - 1, num - 1].CellType;
			numberCellType.DecimalPlaces = ((text == USD) ? 2 : 0);
			sheetView2.Cells[dataLocationTemplateRowStart + num3 - 1, num - 1].CellType = numberCellType;
			numberCellType = (NumberCellType)sheetView2.Cells[dataLocationTemplateRowStart + num3 - 1, num2 - 1].CellType;
			numberCellType.DecimalPlaces = ((text == USD) ? 2 : 0);
			sheetView2.Cells[dataLocationTemplateRowStart + num3 - 1, num2 - 1].CellType = numberCellType;
		}
	}

	public void LockUnLockTheoVspPsc(SheetView sView)
	{
		bool flag = Convert.ToBoolean(sView.Cells[cellVsp[1], cellVsp[0]].Value);
		bool flag2 = Convert.ToBoolean(sView.Cells[cellPsc[1], cellPsc[0]].Value);
		List<int[]> list = new List<int[]>();
		list.Add(GeneralFunctions.ArrayCellId("AV_85"));
		list.Add(GeneralFunctions.ArrayCellId("AV_95"));
		List<int[]> list2 = new List<int[]>();
		list2.Add(GeneralFunctions.ArrayCellId("AV_93"));
		list2.Add(GeneralFunctions.ArrayCellId("AV_96"));
		list2.Add(GeneralFunctions.ArrayCellId("AV_100"));
		list2.Add(GeneralFunctions.ArrayCellId("AV_101"));
		list2.Add(GeneralFunctions.ArrayCellId("AV_106"));
		if (flag)
		{
			foreach (int[] item in list)
			{
				GeneralFunctions.LockUnLockCol(sView, item[1] + 1, item[0] + 1, locked: false, deleteData: false);
			}
		}
		else
		{
			foreach (int[] item2 in list)
			{
				GeneralFunctions.LockUnLockCol(sView, item2[1] + 1, item2[0] + 1, locked: true, deleteData: true);
			}
		}
		if (flag2)
		{
			foreach (int[] item3 in list2)
			{
				GeneralFunctions.LockUnLockCol(sView, item3[1] + 1, item3[0] + 1, locked: false, deleteData: false);
			}
			return;
		}
		foreach (int[] item4 in list2)
		{
			GeneralFunctions.LockUnLockCol(sView, item4[1] + 1, item4[0] + 1, locked: true, deleteData: true);
		}
	}
}
