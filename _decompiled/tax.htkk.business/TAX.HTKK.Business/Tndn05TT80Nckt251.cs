using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tndn05TT80Nckt251
{
	private string cellId_ct22 = "BH_54";

	private string cellId_ct23 = "BH_55";

	private int[] cellNgayHopDong = GeneralFunctions.ArrayCellId("AG_72");

	private int[] CellId_ct22 => GeneralFunctions.ArrayCellId(cellId_ct22);

	private int[] CellId_ct23 => GeneralFunctions.ArrayCellId(cellId_ct23);

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("AB_39");
			int[] array2 = GeneralFunctions.ArrayCellId("I_51");
			int[] array3 = GeneralFunctions.ArrayCellId("S_53");
			int[] array4 = GeneralFunctions.ArrayCellId("S_56");
			int[] array5 = GeneralFunctions.ArrayCellId("AB_68");
			int[] array6 = GeneralFunctions.ArrayCellId("S_70");
			int[] array7 = GeneralFunctions.ArrayCellId("BP_51");
			int[] array8 = GeneralFunctions.ArrayCellId("AA_22");
			int num = HTKKFile.ColumnLettersToIndex("BP");
			int num2 = HTKKFile.ColumnLettersToIndex("BQ");
			int num3 = HTKKFile.ColumnLettersToIndex("BR");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			sView.Cells[array8[1], array8[0]].Value = PeriodInfo.TieuMucHachToanVal;
			if (PeriodInfo.TieuMucHachToanVal == "1")
			{
				if (!ToKhaiInfo.SheetInfos[0].IsReadDataFile && PeriodInfo.LoaiToKhai == "C")
				{
					sView.Cells[array5[1], array5[0]].Value = "";
					sView.Cells[array6[1], array6[0]].Value = "";
				}
				sView.Cells[array[1], array[0]].Value = sView.Cells[array[1], num - 1].Value;
				sView.Cells[array2[1], array2[0]].Value = sView.Cells[array2[1], num2 - 1].Value;
				sView.Cells[array7[1], array7[0]].Value = sView.Cells[array7[1], num3 - 1].Value;
				sView.Cells[array3[1], array3[0]].Value = sView.Cells[array3[1], num - 1].Value;
				sView.Cells[array4[1], array4[0]].Value = sView.Cells[array4[1], num - 1].Value;
				GeneralFunctions.LockUnLockCol(sView, array[1] + 1, array[0] + 1, locked: true, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, array2[1] + 1, array2[0] + 1, locked: true, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, array3[1] + 1, array3[0] + 1, locked: true, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, array4[1] + 1, array4[0] + 1, locked: true, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, array5[1] + 1, array5[0] + 1, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, array6[1] + 1, array6[0] + 1, locked: false, deleteData: false);
			}
			else
			{
				sView.Cells[array5[1], array5[0]].Value = sView.Cells[array[1], num - 1].Value;
				sView.Cells[array6[1], array6[0]].Value = sView.Cells[array4[1], num - 1].Value;
			}
			DateTime dateTime = ConvertHelpers.ConvertToDateTime2("15/12/2025", "dd/MM/yyyy");
			DateTime dateTime2 = ConvertHelpers.ConvertToDateTime2(PeriodInfo.kyKKhai, "dd/MM/yyyy");
			int[] array9 = GeneralFunctions.ArrayCellId("AW_98");
			int[] array10 = GeneralFunctions.ArrayCellId("AW_99");
			int[] array11 = GeneralFunctions.ArrayCellId("AW_101");
			if (string.IsNullOrEmpty(sView.Cells[cellNgayHopDong[1], cellNgayHopDong[0]].Text))
			{
				sView.Cells[cellNgayHopDong[1], cellNgayHopDong[0]].Text = PeriodInfo.TuNgay;
			}
			else
			{
				PeriodInfo.TuNgay = sView.Cells[cellNgayHopDong[1], cellNgayHopDong[0]].Text;
			}
			DateTime dateTime3 = ConvertHelpers.ConvertToDateTime2(PeriodInfo.TuNgay, "dd/MM/yyyy");
			if (dateTime3 >= dateTime)
			{
				sView.Cells[array9[1], array9[0]].Locked = true;
				sView.Cells[array9[1], array9[0]].Value = 0;
				sView.Cells[array10[1], array10[0]].Locked = true;
				sView.Cells[array10[1], array10[0]].Value = 0;
				sView.Cells[array11[1], array11[0]].Value = 2;
			}
			if (dateTime3 >= dateTime)
			{
				sView.Cells[array11[1], array11[0]].Value = 2;
			}
			else
			{
				sView.Cells[array11[1], array11[0]].Value = 20;
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				switch (item.Value)
				{
				case 2:
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
					continue;
				case 1:
					if (winFpSpread.TabStrip[key].BackColor == Color.White)
					{
						winFpSpread.TabStrip[key].BackColor = Color.Yellow;
						continue;
					}
					break;
				}
				winFpSpread.TabStrip[key].BackColor = Color.White;
			}
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
		int value = 0;
		levelWarning.TryGetValue(0, out value);
		int[] array = GeneralFunctions.ArrayCellId("S_56");
		int[] array2 = GeneralFunctions.ArrayCellId("S_70");
		int[] array3 = GeneralFunctions.ArrayCellId("AB_68");
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[12, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = "S_32~S_36~AB_39~I_51~S_53";
			string[] array4 = text3.Split(new char[1] { '~' });
			if (array4.Length > 0)
			{
				for (int i = 0; i < array4.Length; i++)
				{
					int[] array5 = GeneralFunctions.ArrayCellId(array4[i]);
					if (PeriodInfo.TieuMucHachToanVal == "1")
					{
						if (i >= 0 && i <= 1)
						{
							HTKKFile.RemoveValidate(sheetView, array5[1], array5[0]);
						}
						else
						{
							HTKKFile.RemoveValidateBackGround(sheetView, array5[1], array5[0]);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, array5[1], array5[0]);
					}
				}
			}
			if (array4.Length > 0)
			{
				for (int i = 0; i < array4.Length; i++)
				{
					int[] array5 = GeneralFunctions.ArrayCellId(array4[i]);
					string value2 = sheetView.Cells[array5[1], array5[0]].Text.Trim();
					if (string.IsNullOrEmpty(value2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array5[1], array5[0], text);
						value = 2;
					}
				}
			}
			if (PeriodInfo.TieuMucHachToanVal == "2")
			{
				HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, array[1], array[0]);
			}
			string text4 = sheetView.Cells[array[1], array[0]].Text.Trim();
			if (!string.IsNullOrEmpty(text4) && !ValidateBase.CheckMaSoThue(text4))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], text2);
				value = 2;
			}
			if (PeriodInfo.TieuMucHachToanVal == "1")
			{
				HTKKFile.RemoveValidate(sheetView, array2[1], array2[0]);
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, array2[1], array2[0]);
			}
			string text5 = sheetView.Cells[array2[1], array2[0]].Text.Trim();
			if (string.IsNullOrEmpty(text5))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], text);
				value = 2;
			}
			if (!string.IsNullOrEmpty(text5) && !ValidateBase.CheckMaSoThue(text5))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], text2);
				value = 2;
			}
			string value3 = sheetView.Cells[array3[1], array3[0]].Text.Trim();
			if (PeriodInfo.TieuMucHachToanVal == "1")
			{
				HTKKFile.RemoveValidate(sheetView, array3[1], array3[0]);
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, array3[1], array3[0]);
			}
			if (string.IsNullOrEmpty(value3))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], text);
				value = 2;
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
		}
	}
}
