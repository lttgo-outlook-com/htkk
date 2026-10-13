using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Security;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Mtcn01Nckt250
{
	private string pathCQTNop = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/tenCQTNoiNop";

	private string pathMaCQTNop = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/maCQTNoiNop";

	private string pathNodeTenCQT = "HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/tenCQTNoiNop";

	private string pathNodeMaCQT = "HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/maCQTNoiNop";

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("G_4");
			int[] array2 = GeneralFunctions.ArrayCellId("AO_4");
			sView.Cells[array[1], array[0]].Text = PeriodInfo.tenCQTQuanLy;
			sView.Cells[array2[1], array2[0]].Text = PeriodInfo.maCQTQuanLy;
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
			if (winFpSpread.ActiveSheetIndex == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				int[] array = GeneralFunctions.ArrayCellId("D_51");
				if (iRow - 1 == array[1] && iCol - 1 == array[0])
				{
					string value = GeneralFunctions.GetValue(winFpSpread.Sheets[0], "D_51");
					int[] array2 = GeneralFunctions.ArrayCellId("AD_51");
					int[] array3 = GeneralFunctions.ArrayCellId("AO_51");
					string text = MoneyHelpers.MoneyToText(value);
					string text2 = "";
					sheetView.Cells[array2[1], array2[0]].Text = text;
					text2 = sheetView.Cells[array[1], array[0]].Text;
					sheetView.Cells[array3[1], array3[0]].Text = text2;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		return true;
	}

	private void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			winFpSpread.TabStrip[0].BackColor = Color.Pink;
		}
	}

	public void ExportXML_0(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode val = null;
			string str = "";
			string str2 = "";
			List<HTKKCMCCellInfo> staticCells = ToKhaiInfo.SheetInfos[0].StaticCells;
			for (int i = 0; i < staticCells.Count; i++)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = staticCells[i];
				if (hTKKCMCCellInfo.Path == pathCQTNop)
				{
					str = hTKKCMCCellInfo.Value;
				}
				if (hTKKCMCCellInfo.Path == pathMaCQTNop)
				{
					str2 = hTKKCMCCellInfo.Value;
				}
			}
			val = xmlNode.SelectSingleNode(pathNodeTenCQT);
			if (val != null)
			{
				val.InnerXml = SecurityElement.Escape(str);
			}
			val = xmlNode.SelectSingleNode(pathNodeMaCQT);
			if (val != null)
			{
				val.InnerXml = SecurityElement.Escape(str2);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
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
}
