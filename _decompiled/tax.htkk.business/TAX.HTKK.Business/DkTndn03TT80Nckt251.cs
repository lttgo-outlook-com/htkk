using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkTndn03TT80Nckt251
{
	private string cellId_ct09_ngay = "AI_32";

	private string cellId_ct09_so = "P_32";

	private string cellId_ct14_donVi = "Q_47";

	private string cellId_ct14_daiLy = "AB_47";

	private string cellId_ct06 = "P_26";

	private string cellId_ct12 = "Y_42";

	private string cellId_ct15 = "J_49";

	private string cellId_ct24 = "P_53";

	private string cellId_ngayKy = "AP_93";

	private int[] CellId_ct09_ngay => GeneralFunctions.ArrayCellId(cellId_ct09_ngay);

	private int[] CellId_ct09_so => GeneralFunctions.ArrayCellId(cellId_ct09_so);

	private int[] CellId_ct14_donVi => GeneralFunctions.ArrayCellId(cellId_ct14_donVi);

	private int[] CellId_ct14_daiLy => GeneralFunctions.ArrayCellId(cellId_ct14_daiLy);

	private int[] CellId_ct06 => GeneralFunctions.ArrayCellId(cellId_ct06);

	private int[] CellId_ct12 => GeneralFunctions.ArrayCellId(cellId_ct12);

	private int[] CellId_ct15 => GeneralFunctions.ArrayCellId(cellId_ct15);

	private int[] CellId_ct24 => GeneralFunctions.ArrayCellId(cellId_ct24);

	private int[] CellId_ngayKy => GeneralFunctions.ArrayCellId(cellId_ngayKy);

	private int[] CellId_E16 => GeneralFunctions.ArrayCellId("E_16");

	private int[] CellId_E17 => GeneralFunctions.ArrayCellId("E_17");

	private int[] CellId_E18 => GeneralFunctions.ArrayCellId("E_18");

	private int[] CellId_E19 => GeneralFunctions.ArrayCellId("E_19");

	private int[] CellId_E20 => GeneralFunctions.ArrayCellId("E_20");

	private int[] CellId_E21 => GeneralFunctions.ArrayCellId("E_21");

	private int[] CellId_E22 => GeneralFunctions.ArrayCellId("E_22");

	private int[] CellId_E23 => GeneralFunctions.ArrayCellId("E_23");

	private int[] CellId_E24 => GeneralFunctions.ArrayCellId("E_24");

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId("E_25");

	private int[] CellId_E26 => GeneralFunctions.ArrayCellId("E_26");

	private int[] CellId_MaLoaiTien => GeneralFunctions.ArrayCellId("BF_60");

	private int[] CellId_LoaiTien => GeneralFunctions.ArrayCellId("P_60");

	private int[] CellId_Ct30 => GeneralFunctions.ArrayCellId("AW_73");

	private int[] CellId_Ct32 => GeneralFunctions.ArrayCellId("AW_75");

	private int[] CellId_Ct30_CThuc => GeneralFunctions.ArrayCellId("BF_73");

	private int[] CellId_Ct32_CThuc => GeneralFunctions.ArrayCellId("BF_75");

	private int[] CellId_E15 => GeneralFunctions.ArrayCellId("E_15");

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (string.IsNullOrEmpty(TTinChung.tenDLyThue) || string.IsNullOrEmpty(TTinChung.mstDLyThue))
			{
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct14_donVi[1], CellId_ct14_donVi[0], locked: true, "", false);
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct14_daiLy[1], CellId_ct14_daiLy[0], locked: true, "", false);
			}
			else if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				sheetView.Cells[CellId_ct14_daiLy[1], CellId_ct14_daiLy[0]].Value = true;
			}
			string cellFormat = "AW_70###AW_71###AW_72###AW_73###AW_75";
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

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (activeSheetIndex == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				string formula = "ROUND(AW73*AW74/100,2)";
				if (CellId_LoaiTien[1] + 1 == iRow && CellId_LoaiTien[0] + 1 == iCol)
				{
					string cellFormat = "AW_70###AW_71###AW_72###AW_73###AW_75";
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
					if (text == "VND")
					{
						formula = "ROUND(AW73*AW74/100,0)";
					}
					sheetView.Cells[CellId_Ct32[1], CellId_Ct32[0]].Formula = formula;
				}
				if (iRow >= 70 && iRow <= 74 && CellId_Ct32[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula = "ROUND(AW73*AW74/100,0)";
					}
					sheetView.Cells[CellId_Ct32[1], CellId_Ct32[0]].Formula = formula;
					if (iRow != 73)
					{
						sheetView.Cells[CellId_Ct30[1], CellId_Ct30[0]].Formula = "IF(AW70-AW71-AW72>0,AW70-AW71-AW72,0)";
					}
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
			if (idxRow == CellId_ct14_daiLy[1] && idxCol == CellId_ct14_daiLy[0])
			{
				if (Convert.ToBoolean(sheetView.Cells[CellId_ct14_daiLy[1], CellId_ct14_daiLy[0]].Value))
				{
					sheetView.Cells[CellId_ct14_donVi[1], CellId_ct14_donVi[0]].Value = false;
				}
			}
			else if (idxRow == CellId_ct14_donVi[1] && idxCol == CellId_ct14_donVi[0] && Convert.ToBoolean(sheetView.Cells[CellId_ct14_donVi[1], CellId_ct14_donVi[0]].Value))
			{
				sheetView.Cells[CellId_ct14_daiLy[1], CellId_ct14_daiLy[0]].Value = false;
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
			string message = HTKKFile.GetMessage(winFpSpread, CellId_E15[1], CellId_E15[0]);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E17);
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E16);
			string message4 = HTKKFile.GetMessage(winFpSpread, CellId_E18[1], CellId_E18[0]);
			string message5 = HTKKFile.GetMessage(winFpSpread, CellId_E19[1], CellId_E19[0]);
			string message6 = HTKKFile.GetMessage(winFpSpread, CellId_E20[1], CellId_E20[0]);
			string message7 = HTKKFile.GetMessage(winFpSpread, CellId_E21[1], CellId_E21[0]);
			string message8 = HTKKFile.GetMessage(winFpSpread, CellId_E22[1], CellId_E22[0]);
			string message9 = HTKKFile.GetMessage(winFpSpread, CellId_E23[1], CellId_E23[0]);
			string message10 = HTKKFile.GetMessage(winFpSpread, CellId_E24[1], CellId_E24[0]);
			string message11 = HTKKFile.GetMessage(winFpSpread, CellId_E25[1], CellId_E25[0]);
			string message12 = HTKKFile.GetMessage(winFpSpread, CellId_E26[1], CellId_E26[0]);
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_ct14_donVi[1], CellId_ct14_donVi[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct14_daiLy[1], CellId_ct14_daiLy[0]].Value);
			string text = sheetView.Cells[CellId_ct06[1], CellId_ct06[0]].Text.Trim();
			string text2 = sheetView.Cells[CellId_ct12[1], CellId_ct12[0]].Text.Trim();
			string text3 = sheetView.Cells[CellId_ct15[1], CellId_ct15[0]].Text.Trim();
			string text4 = sheetView.Cells[CellId_ct24[1], CellId_ct24[0]].Text.Trim();
			string value2 = sheetView.Cells[CellId_ct09_so[1], CellId_ct09_so[0]].Text.Trim();
			string text5 = sheetView.Cells[CellId_ct09_ngay[1], CellId_ct09_ngay[0]].Text.Trim();
			string text6 = "";
			HTKKFile.RemoveValidate(sheetView, CellId_ct06[1], CellId_ct06[0]);
			HTKKFile.RemoveValidate(sheetView, CellId_ct09_so[1], CellId_ct09_so[0]);
			HTKKFile.RemoveValidate(sheetView, CellId_ct09_ngay[1], CellId_ct09_ngay[0]);
			HTKKFile.RemoveValidate(sheetView, CellId_ct12[1], CellId_ct12[0]);
			HTKKFile.RemoveValidateBackGround1(sheetView, CellId_ct15[1], CellId_ct15[0]);
			HTKKFile.RemoveValidate(sheetView, CellId_ct24[1], CellId_ct24[0]);
			text6 = "~" + text2 + "~" + text3 + "~" + text4 + "~";
			if (!string.IsNullOrEmpty(text) && text6.IndexOf("~" + text + "~") > -1)
			{
				HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct06[1], CellId_ct06[0], message6);
				value = 1;
			}
			text6 = "~" + text + "~" + text3 + "~" + text4 + "~";
			if (!string.IsNullOrEmpty(text2) && text6.IndexOf("~" + text2 + "~") > -1)
			{
				HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct12[1], CellId_ct12[0], message6);
				value = 1;
			}
			text6 = "~" + text2 + "~" + text + "~" + text4 + "~";
			if (!string.IsNullOrEmpty(text3) && text6.IndexOf("~" + text3 + "~") > -1)
			{
				HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct15[1], CellId_ct15[0], message6);
				value = 1;
			}
			text6 = "~" + text2 + "~" + text + "~" + text3 + "~";
			if (!string.IsNullOrEmpty(text4) && text6.IndexOf("~" + text4 + "~") > -1)
			{
				HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct24[1], CellId_ct24[0], message6);
				value = 1;
			}
			if (!string.IsNullOrEmpty(text))
			{
				if (!ValidateBase.CheckMaSoThue(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct06[1], CellId_ct06[0], message);
					value = 2;
				}
				else if (text.Replace("-", "") != TTinChung.mst)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct06[1], CellId_ct06[0], message7);
					value = 2;
				}
				if (string.IsNullOrEmpty(value2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct09_so[1], CellId_ct09_so[0], message2);
					value = 2;
				}
				if (string.IsNullOrEmpty(text5))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct09_ngay[1], CellId_ct09_ngay[0], message2);
					value = 2;
				}
			}
			if (!HTKKFormater.isNullString(text5))
			{
				DateTime t = new DateTime(ConvertHelpers.ConvertToInt(PeriodInfo.Nam), ConvertHelpers.ConvertToInt(PeriodInfo.Thang), ConvertHelpers.ConvertToInt(PeriodInfo.Ngay));
				bool success = false;
				DateTime t2 = HTKKFormater.ConvertToDateTime(text5, out success);
				if (success)
				{
					if (DateTime.Compare(t2, t) > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct09_ngay[1], CellId_ct09_ngay[0], message5);
						value = 2;
					}
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct09_ngay[1], CellId_ct09_ngay[0], HTKKFile.GetMessage("0071"));
					value = 2;
				}
			}
			if (!string.IsNullOrEmpty(text2) && !ValidateBase.CheckMaSoThue(text2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct12[1], CellId_ct12[0], message);
				value = 2;
			}
			if (!string.IsNullOrEmpty(text4))
			{
				if (!ValidateBase.CheckMaSoThue(text4))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct24[1], CellId_ct24[0], message);
					value = 2;
				}
				else if (text4.Replace("-", "").Replace(" ", "") != TTinChung.mst.Replace("-", "").Replace(" ", ""))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct24[1], CellId_ct24[0], message8);
					value = 2;
				}
			}
			if (!HTKKFormater.isNullString(TTinChung.mstDLyThue))
			{
				if (flag && flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct14_daiLy[1], CellId_ct14_daiLy[0], message3);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct14_donVi[1], CellId_ct14_donVi[0], message3);
					value = 2;
				}
				else if (!flag && !flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct14_daiLy[1], CellId_ct14_daiLy[0], message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct14_donVi[1], CellId_ct14_donVi[0], message2);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct14_daiLy[1], CellId_ct14_daiLy[0]);
					HTKKFile.RemoveValidate(sheetView, CellId_ct14_donVi[1], CellId_ct14_donVi[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct14_daiLy[1], CellId_ct14_daiLy[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct14_donVi[1], CellId_ct14_donVi[0]);
			}
			if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text4))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct06[1], CellId_ct06[0], message4);
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct24[1], CellId_ct24[0], message4);
				value = 2;
			}
			else if (string.IsNullOrEmpty(text) && string.IsNullOrEmpty(text4))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct06[1], CellId_ct06[0], message2);
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct24[1], CellId_ct24[0], message2);
				value = 2;
			}
			string text7 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			double value3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_Ct30_CThuc[1], CellId_Ct30_CThuc[0]].Value);
			double value4 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_Ct32_CThuc[1], CellId_Ct32_CThuc[0]].Value);
			double value5 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_Ct30[1], CellId_Ct30[0]].Value);
			double value6 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_Ct32[1], CellId_Ct32[0]].Value);
			double num = 0.0;
			double num2 = 0.0;
			if (text7 != "VND")
			{
				num = ((Math.Abs(value3) * 0.001 < 5.0) ? (Math.Abs(value3) * 0.001) : 5.0);
				num2 = ((Math.Abs(value4) * 0.001 < 5.0) ? (Math.Abs(value4) * 0.001) : 5.0);
			}
			else
			{
				num = ((Math.Abs(value3) * 0.001 < 100000.0) ? (Math.Abs(value3) * 0.001) : 100000.0);
				num2 = ((Math.Abs(value4) * 0.001 < 100000.0) ? (Math.Abs(value4) * 0.001) : 100000.0);
			}
			HTKKFile.RemoveValidateBackGround(sheetView, CellId_Ct30[1], CellId_Ct30[0]);
			HTKKFile.RemoveValidateBackGround(sheetView, CellId_Ct32[1], CellId_Ct32[0]);
			if (Math.Abs(value5) > Math.Abs(value3) + num || Math.Abs(value5) < Math.Abs(value3) - num)
			{
				if (text7 != "VND")
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_Ct30[1], CellId_Ct30[0], message11);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_Ct30[1], CellId_Ct30[0], message12);
				}
				value = 2;
			}
			if (Math.Abs(value6) > Math.Abs(value4) + num2 || Math.Abs(value6) < Math.Abs(value4) - num2)
			{
				if (text7 != "VND")
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_Ct32[1], CellId_Ct32[0], message11);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_Ct32[1], CellId_Ct32[0], message12);
				}
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

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/Header/CT14/donViUyQuyen");
			if (val != null)
			{
				if (HTKKFormater.isNullString(val.InnerText) || val.InnerText == "0")
				{
					val.InnerText = "false";
				}
				else
				{
					val.InnerText = "true";
				}
			}
			val = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/Header/CT14/daiLyThue");
			if (val != null)
			{
				if (HTKKFormater.isNullString(val.InnerText) || val.InnerText == "0")
				{
					val.InnerText = "false";
				}
				else
				{
					val.InnerText = "true";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
