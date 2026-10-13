using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Text.RegularExpressions;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tndn06Nckt207
{
	private Regex regexCt32a_mulVal = new Regex("^0$|^([1-9][0-9]?)(;[1-9][0-9]?)*$");

	private string cellId_ct12 = "V_34";

	private string cellId_ct13 = "AL_36";

	private string cellId_ct14 = "I_40";

	private string cellId_ct15_so = "X_42";

	private string cellId_ct15_ngay = "AC_42";

	private string cellId_ct15_phuongxa = "AH_44";

	private string cellId_ct25 = "AW_61";

	private string cellId_ct26 = "AW_62";

	private string cellId_ct27 = "AW_64";

	private string cellId_ct28 = "AW_65";

	private string cellId_ct29 = "AW_66";

	private string cellId_ct30 = "AW_67";

	private string cellId_ct31 = "AW_68";

	private string cellId_ct32 = "AW_69";

	private string cellId_ct33 = "AW_70";

	private string pathCt33 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct33";

	private string cellId_ct34 = "AW_71";

	private string pathCt34 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct34";

	private string cellId_ct35 = "AW_72";

	private static string cellId_ct36 = "AW_73";

	private string pathCt36 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct36";

	private static string cellId_ct37 = "AW_74";

	private string pathCt37 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct37";

	private int[] CellId_ct12 => GeneralFunctions.ArrayCellId(cellId_ct12);

	private int[] CellId_ct13 => GeneralFunctions.ArrayCellId(cellId_ct13);

	private int[] CellId_ct14 => GeneralFunctions.ArrayCellId(cellId_ct14);

	private int[] CellId_ct15_so => GeneralFunctions.ArrayCellId(cellId_ct15_so);

	private int[] CellId_ct15_ngay => GeneralFunctions.ArrayCellId(cellId_ct15_ngay);

	private int[] CellId_ct15_phuongxa => GeneralFunctions.ArrayCellId(cellId_ct15_phuongxa);

	private int[] CellId_ct25 => GeneralFunctions.ArrayCellId(cellId_ct25);

	private int[] CellId_ct26 => GeneralFunctions.ArrayCellId(cellId_ct26);

	private int[] CellId_ct27 => GeneralFunctions.ArrayCellId(cellId_ct27);

	private int[] CellId_ct28 => GeneralFunctions.ArrayCellId(cellId_ct28);

	private int[] CellId_ct29 => GeneralFunctions.ArrayCellId(cellId_ct29);

	private int[] CellId_ct30 => GeneralFunctions.ArrayCellId(cellId_ct30);

	private int[] CellId_ct31 => GeneralFunctions.ArrayCellId(cellId_ct31);

	private int[] CellId_ct32 => GeneralFunctions.ArrayCellId(cellId_ct32);

	private int[] CellId_ct33 => GeneralFunctions.ArrayCellId(cellId_ct33);

	private List<int> lstIdxRowsRelate_Ct33
	{
		get
		{
			List<int> list = new List<int>();
			list.Add(CellId_ct25[1]);
			list.Add(CellId_ct26[1]);
			list.Add(CellId_ct27[1]);
			list.Add(CellId_ct28[1]);
			list.Add(CellId_ct29[1]);
			list.Add(CellId_ct30[1]);
			list.Add(CellId_ct31[1]);
			list.Add(CellId_ct32[1]);
			return list;
		}
	}

	private int[] CellId_ct34 => GeneralFunctions.ArrayCellId(cellId_ct34);

	private int[] CellId_ct35 => GeneralFunctions.ArrayCellId(cellId_ct35);

	private List<int> lstIdxRowsRelate_Ct35
	{
		get
		{
			List<int> list = lstIdxRowsRelate_Ct33;
			list.Add(CellId_ct34[1]);
			return list;
		}
	}

	private static int[] CellId_ct36 => GeneralFunctions.ArrayCellId(cellId_ct36);

	private static int[] CellId_ct37 => GeneralFunctions.ArrayCellId(cellId_ct37);

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			XuLyGiaTri_Ct33(sView, ToKhaiInfo.SheetInfos[0]);
			XuLyGiaTri_Ct35(sView, ToKhaiInfo.SheetInfos[0]);
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
			int num = HTKKFile.ColumnLettersToIndex("AW");
			List<int> list = lstIdxRowsRelate_Ct33;
			list.Add(CellId_ct34[1]);
			if (num == iCol && lstIdxRowsRelate_Ct33.Contains(Convert.ToInt32(iRow) - 1))
			{
				XuLyGiaTri_Ct33(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0]);
			}
			if (num == iCol && list.Contains(Convert.ToInt32(iRow) - 1))
			{
				XuLyGiaTri_Ct35(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0]);
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
			int num = 0;
			SheetView sheetView = winFpSpread.Sheets[0];
			if (string.IsNullOrEmpty(sheetView.Cells[CellId_ct12[1], CellId_ct12[0]].Text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct12[1], CellId_ct12[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_31")));
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct12[1], CellId_ct12[0]);
			}
			if (string.IsNullOrEmpty(sheetView.Cells[CellId_ct14[1], CellId_ct14[0]].Text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct14[1], CellId_ct14[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_31")));
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct14[1], CellId_ct14[0]);
			}
			if (string.IsNullOrEmpty(sheetView.Cells[CellId_ct15_so[1], CellId_ct15_so[0]].Text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct15_so[1], CellId_ct15_so[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_31")));
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct15_so[1], CellId_ct15_so[0]);
			}
			if (string.IsNullOrEmpty(sheetView.Cells[CellId_ct15_ngay[1], CellId_ct15_ngay[0]].Text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct15_ngay[1], CellId_ct15_ngay[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_31")));
				num = 2;
			}
			else if (!ValidateBase.ValidateFullDate(sheetView.Cells[CellId_ct15_ngay[1], CellId_ct15_ngay[0]].Text, "dd/MM/yyyy"))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct15_ngay[1], CellId_ct15_ngay[0], "Ngày không hợp lệ");
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct15_ngay[1], CellId_ct15_ngay[0]);
			}
			if (string.IsNullOrEmpty(sheetView.Cells[CellId_ct15_phuongxa[1], CellId_ct15_phuongxa[0]].Text))
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct15_phuongxa[1], CellId_ct15_phuongxa[0]);
			}
			else if (!ValidateBase.ValidateFullDate(sheetView.Cells[CellId_ct15_phuongxa[1], CellId_ct15_phuongxa[0]].Text, "dd/MM/yyyy"))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct15_phuongxa[1], CellId_ct15_phuongxa[0], "Ngày không hợp lệ");
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct15_phuongxa[1], CellId_ct15_phuongxa[0]);
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
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink)
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

	private void XuLyGiaTri_Ct33(SheetView sView, HTKKCMCSheetInfo SheetInfoTK)
	{
		try
		{
			HTKKCMCCellInfo hTKKCMCCellInfo = SheetInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == pathCt34);
			double num = Convert.ToDouble(sView.Cells[CellId_ct33[1], CellId_ct33[0]].Value);
			if (num <= 0.0)
			{
				GeneralFunctions.ChangeValueCell(sView, CellId_ct34[1], CellId_ct34[0], locked: true, "", 0);
			}
			else
			{
				GeneralFunctions.ChangeValueCell(sView, CellId_ct34[1], CellId_ct34[0], locked: false, "", Convert.ToDouble(hTKKCMCCellInfo.Value));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void XuLyGiaTri_Ct35(SheetView sView, HTKKCMCSheetInfo SheetInfoTK)
	{
		try
		{
			HTKKCMCCellInfo hTKKCMCCellInfo = SheetInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == pathCt36);
			double num = Convert.ToDouble(sView.Cells[CellId_ct35[1], CellId_ct35[0]].Value);
			if (num <= 0.0)
			{
				GeneralFunctions.ChangeValueCell(sView, CellId_ct36[1], CellId_ct36[0], locked: true, "", 0);
			}
			else
			{
				GeneralFunctions.ChangeValueCell(sView, CellId_ct36[1], CellId_ct36[0], locked: false, "", Convert.ToDouble(hTKKCMCCellInfo.Value));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
