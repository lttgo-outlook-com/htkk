using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tndn02TT80Nckt251
{
	private string cellId_ct09d = "I_69";

	private string cellId_ct09dMa = "BL_69";

	private string cellId_ct09c = "I_71";

	private string cellId_ct09cMa = "BL_71";

	private string cellId_ct09b = "I_73";

	private string cellId_ct09bMa = "BL_73";

	private string cellId_ct09a = "P_75";

	private int[] CellId_ct09d => GeneralFunctions.ArrayCellId(cellId_ct09d);

	private int[] CellId_ct09dMa => GeneralFunctions.ArrayCellId(cellId_ct09dMa);

	private int[] CellId_ct09c => GeneralFunctions.ArrayCellId(cellId_ct09c);

	private int[] CellId_ct09cMa => GeneralFunctions.ArrayCellId(cellId_ct09cMa);

	private int[] CellId_ct09b => GeneralFunctions.ArrayCellId(cellId_ct09b);

	private int[] CellId_ct09bMa => GeneralFunctions.ArrayCellId(cellId_ct09bMa);

	private int[] CellId_ct09a => GeneralFunctions.ArrayCellId(cellId_ct09a);

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			GeneralFunctions.ThayDoiTinhTheoCQT2Cap(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sView, PeriodInfo, CellId_ct09d, CellId_ct09dMa, CellId_ct09c, CellId_ct09cMa, CellId_ct09b, CellId_ct09bMa, CellId_ct09a, null, null, PhuThuocTruSoChinh: false);
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
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			int[] array = GeneralFunctions.ArrayCellId("V_46");
			int[] array2 = GeneralFunctions.ArrayCellId("AP_48");
			int[] array3 = GeneralFunctions.ArrayCellId("I_52");
			int[] array4 = GeneralFunctions.ArrayCellId("I_54");
			int[] array5 = GeneralFunctions.ArrayCellId("I_56");
			int[] array6 = GeneralFunctions.ArrayCellId("I_58");
			int[] array7 = GeneralFunctions.ArrayCellId("I_71");
			int[] array8 = GeneralFunctions.ArrayCellId("I_73");
			int[] array9 = GeneralFunctions.ArrayCellId("P_75");
			int[] array10 = GeneralFunctions.ArrayCellId("U_77");
			int[] array11 = GeneralFunctions.ArrayCellId("AC_77");
			int[] array12 = GeneralFunctions.ArrayCellId("AC_79");
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_35"));
			HTKKFile.RemoveValidate(sheetView, array2[1], array2[0]);
			HTKKFile.RemoveValidate(sheetView, array11[1], array11[0]);
			HTKKFile.RemoveValidate(sheetView, array12[1], array12[0]);
			string text = "V_46~I_52~I_56~I_58~I_73~P_75~U_77";
			string[] array13 = text.Split(new char[1] { '~' });
			if (array13.Length > 0)
			{
				for (int i = 0; i < array13.Length; i++)
				{
					int[] array14 = GeneralFunctions.ArrayCellId(array13[i]);
					HTKKFile.RemoveValidate(sheetView, array14[1], array14[0]);
				}
			}
			if (array13.Length > 0)
			{
				for (int i = 0; i < array13.Length; i++)
				{
					int[] array14 = GeneralFunctions.ArrayCellId(array13[i]);
					string value2 = sheetView.Cells[array14[1], array14[0]].Text.Trim();
					if (string.IsNullOrEmpty(value2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array14[1], array14[0], message);
						value = 2;
					}
				}
			}
			DateTime now = DateTime.Now;
			string text2 = sheetView.Cells[array11[1], array11[0]].Text.Trim();
			if (!string.IsNullOrEmpty(text2))
			{
				if (!HTKKFormater.IsDateTime(text2, "dd/MM/yyyy", out var dtDate))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array11[1], array11[0], HTKKFile.GetMessage("0071"));
					value = 2;
				}
				else if (DateTime.Compare(dtDate, now) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array11[1], array11[0], message4);
					value = 2;
				}
			}
			else
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array11[1], array11[0], message);
				value = 2;
			}
			string text3 = sheetView.Cells[array12[1], array12[0]].Text.Trim();
			if (!string.IsNullOrEmpty(text3))
			{
				if (!HTKKFormater.IsDateTime(text3, "dd/MM/yyyy", out var dtDate2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array12[1], array12[0], HTKKFile.GetMessage("0071"));
					value = 2;
				}
				else if (DateTime.Compare(dtDate2, now) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array12[1], array12[0], message4);
					value = 2;
				}
			}
			else
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array12[1], array12[0], message);
				value = 2;
			}
			string text4 = sheetView.Cells[array2[1], array2[0]].Text.Trim();
			if (!string.IsNullOrEmpty(text4))
			{
				if (text4.Length == 10 || text4.Length == 13)
				{
					if (!ValidateBase.CheckMaSoThue(text4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message3);
						value = 2;
					}
				}
				else if (text4.Length != 9 && text4.Length != 12)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message5);
					value = 2;
				}
			}
			else
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message);
				value = 2;
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
}
