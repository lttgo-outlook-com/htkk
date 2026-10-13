using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkPthuVsp011TT80Nckt284
{
	private string cellId_ct01Val = "O_30";

	private string cellId_ct03Val = "Q_30";

	private string cellId_ct16_donvi = "Q_51";

	private string cellId_ct16_daily = "Y_51";

	private int[] CellId_ct01Val => GeneralFunctions.ArrayCellId(cellId_ct01Val);

	private int[] CellId_ct03Val => GeneralFunctions.ArrayCellId(cellId_ct03Val);

	private int[] CellId_ct16_donvi => GeneralFunctions.ArrayCellId(cellId_ct16_donvi);

	private int[] CellId_ct16_daily => GeneralFunctions.ArrayCellId(cellId_ct16_daily);

	private int[] CellId_ctTuNgay => GeneralFunctions.ArrayCellId("Y_23");

	private int[] CellId_ctDenNgay => GeneralFunctions.ArrayCellId("AM_23");

	private int[] CellId_ct26 => GeneralFunctions.ArrayCellId("AV_59");

	private int[] CellId_ct27 => GeneralFunctions.ArrayCellId("AV_60");

	private int[] CellId_ct28 => GeneralFunctions.ArrayCellId("AV_61");

	private int[] CellId_ct29 => GeneralFunctions.ArrayCellId("AV_62");

	private int[] CellId_ct30 => GeneralFunctions.ArrayCellId("AV_63");

	private int[] CellId_ct31 => GeneralFunctions.ArrayCellId("AV_64");

	private int[] CellId_ct32 => GeneralFunctions.ArrayCellId("AV_65");

	private int[] CellId_ct33 => GeneralFunctions.ArrayCellId("AV_66");

	private int[] CellId_ct34 => GeneralFunctions.ArrayCellId("AV_67");

	private int[] CellId_ct35 => GeneralFunctions.ArrayCellId("AV_68");

	private int[] CellId_ct36 => GeneralFunctions.ArrayCellId("AV_69");

	private int[] CellId_MaSoThueHDDK => GeneralFunctions.ArrayCellId("X_49");

	private int[] CellId_E13 => GeneralFunctions.ArrayCellId("E_13");

	private int[] CellId_E19 => GeneralFunctions.ArrayCellId("E_19");

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId("E_25");

	private int[] CellId_E26 => GeneralFunctions.ArrayCellId("E_26");

	private int[] CellId_E27 => GeneralFunctions.ArrayCellId("E_27");

	private int[] CellId_E28 => GeneralFunctions.ArrayCellId("E_28");

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			switch (PeriodInfo.TieuMucHachToanVal)
			{
			case "1":
				sheetView.Cells[CellId_ct01Val[1], CellId_ct01Val[0]].Value = true;
				sheetView.Cells[CellId_ct03Val[1], CellId_ct03Val[0]].Value = false;
				break;
			case "3":
				sheetView.Cells[CellId_ct01Val[1], CellId_ct01Val[0]].Value = false;
				sheetView.Cells[CellId_ct03Val[1], CellId_ct03Val[0]].Value = true;
				break;
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
			GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0], 0);
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
				string formula = "AV59-AV60";
				string formula2 = "IF(AV62<>0,ROUND(AV61/AV62,2),0)";
				string formula3 = "IF(AV66>0,AV66,AV67)";
				string formula4 = "IF(AND(1.2*AV64<AV62,AV62<=1.5*AV64),ROUND(0.5*(AV62-1.2*AV64)*AV63,2),0)";
				string formula5 = "IF(AV62>1.5*AV64,AV68+AV69,0)";
				string formula6 = "IF(AV62>1.5*AV64,ROUND(0.5*(1.5*AV64-1.2*AV64)*AV63,2),0)";
				string formula7 = "IF(AV62>1.5*AV64,ROUND(0.6*(AV62-1.5*AV64)*AV63,2),0)";
				if (iRow >= 59 && iRow <= 60 && CellId_ct26[0] + 1 == iCol)
				{
					sheetView.Cells[CellId_ct28[1], CellId_ct28[0]].Formula = formula;
				}
				if (iRow >= 59 && iRow <= 62 && CellId_ct26[0] + 1 == iCol)
				{
					sheetView.Cells[CellId_ct30[1], CellId_ct30[0]].Formula = formula2;
				}
				if (((iRow >= 59 && iRow <= 64) || (iRow >= 66 && iRow <= 69)) && CellId_ct26[0] + 1 == iCol)
				{
					sheetView.Cells[CellId_ct32[1], CellId_ct32[0]].Formula = formula3;
				}
				if (iRow >= 59 && iRow <= 64 && CellId_ct26[0] + 1 == iCol)
				{
					sheetView.Cells[CellId_ct33[1], CellId_ct33[0]].Formula = formula4;
				}
				if (((iRow >= 59 && iRow <= 64) || iRow == 68 || iRow == 69) && CellId_ct26[0] + 1 == iCol)
				{
					sheetView.Cells[CellId_ct34[1], CellId_ct34[0]].Formula = formula5;
				}
				if (iRow >= 59 && iRow <= 64 && CellId_ct26[0] + 1 == iCol)
				{
					sheetView.Cells[CellId_ct35[1], CellId_ct35[0]].Formula = formula6;
				}
				if (iRow >= 59 && iRow <= 64 && CellId_ct26[0] + 1 == iCol)
				{
					sheetView.Cells[CellId_ct36[1], CellId_ct36[0]].Formula = formula7;
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
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink || winFpSpread.TabStrip[1].BackColor == Color.Pink)
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
			string message = HTKKFile.GetMessage(sheetView, CellId_E13);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E19);
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E25);
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E26);
			string message5 = HTKKFile.GetMessage(sheetView, CellId_E27);
			string message6 = HTKKFile.GetMessage(sheetView, CellId_E28);
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_ct16_donvi[1], CellId_ct16_donvi[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct16_daily[1], CellId_ct16_daily[0]].Value);
			if (!HTKKFormater.isNullString(TTinChung.mstDLyThue))
			{
				if (flag && flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], message3);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], message3);
					num = 2;
				}
				else if (!flag && !flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], message2);
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
			string text = sheetView.Cells[CellId_ctTuNgay[1], CellId_ctTuNgay[0]].Text;
			string text2 = sheetView.Cells[CellId_ctDenNgay[1], CellId_ctDenNgay[0]].Text;
			bool flag3 = false;
			bool flag4 = false;
			HTKKFile.RemoveValidate(sheetView, CellId_ctTuNgay[1], CellId_ctTuNgay[0]);
			HTKKFile.RemoveValidate(sheetView, CellId_ctDenNgay[1], CellId_ctDenNgay[0]);
			DateTime dtDate;
			if (!string.IsNullOrEmpty(text))
			{
				if (HTKKFormater.IsDateTime(text, "dd/MM/yyyy", out dtDate) && dtDate >= DateTime.ParseExact("01/01/1900", "dd/MM/yyyy", CultureInfo.InvariantCulture))
				{
					flag3 = true;
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ctTuNgay[1], CellId_ctTuNgay[0], message6);
					num = 2;
				}
			}
			else
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ctTuNgay[1], CellId_ctTuNgay[0], message2);
				num = 2;
			}
			DateTime dtDate2;
			if (!string.IsNullOrEmpty(text2))
			{
				if (HTKKFormater.IsDateTime(text2, "dd/MM/yyyy", out dtDate2) && dtDate2 >= DateTime.ParseExact("01/01/1900", "dd/MM/yyyy", CultureInfo.InvariantCulture))
				{
					flag4 = true;
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ctDenNgay[1], CellId_ctDenNgay[0], message6);
					num = 2;
				}
			}
			else
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ctDenNgay[1], CellId_ctDenNgay[0], message2);
				num = 2;
			}
			if (!string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text) && flag3 && flag4)
			{
				HTKKFormater.IsDateTime(text, "dd/MM/yyyy", out dtDate);
				HTKKFormater.IsDateTime(text2, "dd/MM/yyyy", out dtDate2);
				if (dtDate > dtDate2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ctTuNgay[1], CellId_ctTuNgay[0], message5);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ctDenNgay[1], CellId_ctDenNgay[0], message5);
					num = 2;
				}
			}
			string text3 = sheetView.Cells[CellId_MaSoThueHDDK[1], CellId_MaSoThueHDDK[0]].Text;
			HTKKFile.RemoveValidate(sheetView, CellId_MaSoThueHDDK[1], CellId_MaSoThueHDDK[0]);
			if (!string.IsNullOrEmpty(text3) && !ValidateBase.CheckMaSoThue(text3))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_MaSoThueHDDK[1], CellId_MaSoThueHDDK[0], message);
				num = 2;
			}
			int[] array = new int[6] { 61, 65, 66, 67, 68, 69 };
			for (int i = 0; i < array.Length; i++)
			{
				int[] array2 = GeneralFunctions.ArrayCellId("BG_" + array[i]);
				int[] array3 = GeneralFunctions.ArrayCellId("AV_" + array[i]);
				double value = ConvertHelpers.ConvertToDouble(sheetView.Cells[array2[1], array2[0]].Value);
				double value2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array3[1], array3[0]].Value);
				double num2 = 0.0;
				num2 = ((Math.Abs(value) * 0.001 < 5.0) ? (Math.Abs(value) * 0.001) : 5.0);
				HTKKFile.RemoveValidateBackGround(sheetView, array3[1], array3[0]);
				if (Math.Abs(value2) > Math.Abs(value) + num2 || Math.Abs(value2) < Math.Abs(value) - num2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message4);
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

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/ten_DonVi");
			XmlNode val2 = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/ma_DonVi");
			if (val != null)
			{
				val.InnerText = "Đô la Mỹ";
			}
			if (val2 != null)
			{
				val2.InnerText = "USD";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
