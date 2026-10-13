using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkTns01TT80Nckt284
{
	private string cellId_ct14_donVi = "Q_46";

	private string cellId_ct14_daiLy = "AB_46";

	private string cellId_ct7 = "P_31";

	private string cellId_ct15 = "J_48";

	private string cellId_E17 = "E_17";

	private string cellId_E16 = "E_16";

	private string cellId_E15 = "E_15";

	private int[] CellId_ct14_donVi => GeneralFunctions.ArrayCellId(cellId_ct14_donVi);

	private int[] CellId_ct14_daiLy => GeneralFunctions.ArrayCellId(cellId_ct14_daiLy);

	private int[] CellId_ct7 => GeneralFunctions.ArrayCellId(cellId_ct7);

	private int[] CellId_ct15 => GeneralFunctions.ArrayCellId(cellId_ct15);

	private int[] CellId_E17 => GeneralFunctions.ArrayCellId(cellId_E17);

	private int[] CellId_E16 => GeneralFunctions.ArrayCellId(cellId_E16);

	private int[] CellId_E15 => GeneralFunctions.ArrayCellId(cellId_E15);

	private int[] CellId_Ct28 => GeneralFunctions.ArrayCellId("AW_66");

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
			string message = HTKKFile.GetMessage(sheetView, CellId_E15);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E17);
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E16);
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_ct14_donVi[1], CellId_ct14_donVi[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct14_daiLy[1], CellId_ct14_daiLy[0]].Value);
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
			string text = "";
			HTKKFile.RemoveValidate(sheetView, CellId_ct7[1], CellId_ct7[0]);
			text = sheetView.Cells[CellId_ct7[1], CellId_ct7[0]].Text;
			if (!string.IsNullOrEmpty(text) && !ValidateBase.CheckMaSoThue(text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct7[1], CellId_ct7[0], message);
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
