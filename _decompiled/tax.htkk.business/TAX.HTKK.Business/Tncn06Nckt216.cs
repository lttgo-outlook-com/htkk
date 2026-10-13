using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tncn06Nckt216
{
	private string cellId_ct12 = "J_7";

	private string cellId_ct13 = "J_9";

	private string cellId_ct14 = "F_11";

	private string cellId_ct15 = "F_13";

	private string cellId_ct16 = "S_13";

	private string cellId_ct17 = "F_15";

	private string cellId_ct18 = "K_15";

	private string cellId_ct19 = "S_15";

	private string cellId_SoPhatSinh_ct29 = "P_40";

	private string cellId_SoPhatSinh_ct30 = "P_41";

	private string formula_SoPhatSinh_ct30 = "ROUND(P40*0.05,0)";

	private string path_SoPhatSinh_ct30 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/SoPhatSinh/ct30";

	private string cellId_SoPhatSinh_ct31 = "P_43";

	private string cellId_SoPhatSinh_ct32 = "P_44";

	private string formula_SoPhatSinh_ct32 = "ROUND(P43*0.001,0)";

	private string path_SoPhatSinh_ct32 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/SoPhatSinh/ct32";

	private string cellId_SoPhatSinh_ct33 = "P_46";

	private string cellId_SoPhatSinh_ct34 = "P_47";

	private string cellId_SoPhatSinh_ct35 = "P_49";

	private string cellId_SoPhatSinh_ct36 = "P_50";

	private string cellId_SoPhatSinh_ct37 = "P_52";

	private string cellId_SoPhatSinh_ct38 = "P_53";

	private string cellId_SoPhatSinh_ct39 = "P_55";

	private string cellId_SoPhatSinh_ct40 = "P_56";

	private string formula_SoPhatSinh_ct40 = "ROUND(P55*0.1/100,0)";

	private string path_SoPhatSinh_ct40 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/SoPhatSinh/ct40";

	private string cellId_SoLuyKe_ct29 = "V_40";

	private string formula_SoLuyKe_ct29 = "P40+AB40";

	private string path_SoLuyKe_ct29 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/SoLuyke/ct29";

	private string cellId_SoLuyKe_ct30 = "V_41";

	private string formula_SoLuyKe_ct30 = "ROUND(V40*0.05,0)";

	private string path_SoLuyKe_ct30 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/SoLuyke/ct30";

	private string cellId_SoLuyKe_ct31 = "V_43";

	private string formula_SoLuyKe_ct31 = "P43+AB43";

	private string path_SoLuyKe_ct31 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/SoLuyke/ct31";

	private string cellId_SoLuyKe_ct32 = "V_44";

	private string formula_SoLuyKe_ct32 = "ROUND(V43*0.001,0)";

	private string path_SoLuyKe_ct32 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/SoLuyke/ct32";

	private string cellId_SoLuyKe_ct33 = "V_46";

	private string formula_SoLuyKe_ct33 = "P46+AB46";

	private string path_SoLuyKe_ct33 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/SoLuyke/ct33";

	private string cellId_SoLuyKe_ct34 = "V_47";

	private string cellId_SoLuyKe_ct35 = "V_49";

	private string formula_SoLuyKe_ct35 = "P49+AB49";

	private string path_SoLuyKe_ct35 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/SoLuyke/ct35";

	private string cellId_SoLuyKe_ct36 = "V_50";

	private string cellId_SoLuyKe_ct37 = "V_52";

	private string formula_SoLuyKe_ct37 = "P52+AB52";

	private string path_SoLuyKe_ct37 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/SoLuyke/ct37";

	private string cellId_SoLuyKe_ct38 = "V_53";

	private string formula_SoLuyKe_ct38 = "P53+AB53";

	private string path_SoLuyKe_ct38 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/SoLuyke/ct38";

	private string cellId_SoLuyKe_ct39 = "V_55";

	private string formula_SoLuyKe_ct39 = "P55+AB55";

	private string path_SoLuyKe_ct39 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/SoLuyke/ct39";

	private string cellId_SoLuyKe_ct40 = "V_56";

	private string formula_SoLuyKe_ct40 = "ROUND(V55*0.1/100,0)";

	private string path_SoLuyKe_ct40 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/SoLuyke/ct40";

	private string cellId_SoLuyKeKT_ct29 = "AB_40";

	private string cellId_SoLuyKeKT_ct30 = "AB_41";

	private string cellId_SoLuyKeKT_ct31 = "AB_43";

	private string cellId_SoLuyKeKT_ct32 = "AB_44";

	private string cellId_SoLuyKeKT_ct33 = "AB_46";

	private string cellId_SoLuyKeKT_ct34 = "AB_47";

	private string cellId_SoLuyKeKT_ct35 = "AB_49";

	private string cellId_SoLuyKeKT_ct36 = "AB_50";

	private string cellId_SoLuyKeKT_ct37 = "AB_52";

	private string cellId_SoLuyKeKT_ct38 = "AB_53";

	private string cellId_SoLuyKeKT_ct39 = "AB_55";

	private string cellId_SoLuyKeKT_ct40 = "AB_56";

	private int[] CellId_ct12 => GeneralFunctions.ArrayCellId(cellId_ct12);

	private int[] CellId_ct13 => GeneralFunctions.ArrayCellId(cellId_ct13);

	private int[] CellId_ct14 => GeneralFunctions.ArrayCellId(cellId_ct14);

	private int[] CellId_ct15 => GeneralFunctions.ArrayCellId(cellId_ct15);

	private int[] CellId_ct16 => GeneralFunctions.ArrayCellId(cellId_ct16);

	private int[] CellId_ct17 => GeneralFunctions.ArrayCellId(cellId_ct17);

	private int[] CellId_ct18 => GeneralFunctions.ArrayCellId(cellId_ct18);

	private int[] CellId_ct19 => GeneralFunctions.ArrayCellId(cellId_ct19);

	private int[] CellId_SoPhatSinh_ct29 => GeneralFunctions.ArrayCellId(cellId_SoPhatSinh_ct29);

	private int[] CellId_SoPhatSinh_ct30 => GeneralFunctions.ArrayCellId(cellId_SoPhatSinh_ct30);

	private int[] CellId_SoPhatSinh_ct31 => GeneralFunctions.ArrayCellId(cellId_SoPhatSinh_ct31);

	private int[] CellId_SoPhatSinh_ct32 => GeneralFunctions.ArrayCellId(cellId_SoPhatSinh_ct32);

	private int[] CellId_SoPhatSinh_ct33 => GeneralFunctions.ArrayCellId(cellId_SoPhatSinh_ct33);

	private int[] CellId_SoPhatSinh_ct34 => GeneralFunctions.ArrayCellId(cellId_SoPhatSinh_ct34);

	private int[] CellId_SoPhatSinh_ct35 => GeneralFunctions.ArrayCellId(cellId_SoPhatSinh_ct35);

	private int[] CellId_SoPhatSinh_ct36 => GeneralFunctions.ArrayCellId(cellId_SoPhatSinh_ct36);

	private int[] CellId_SoPhatSinh_ct37 => GeneralFunctions.ArrayCellId(cellId_SoPhatSinh_ct37);

	private int[] CellId_SoPhatSinh_ct38 => GeneralFunctions.ArrayCellId(cellId_SoPhatSinh_ct38);

	private int[] CellId_SoPhatSinh_ct39 => GeneralFunctions.ArrayCellId(cellId_SoPhatSinh_ct39);

	private int[] CellId_SoPhatSinh_ct40 => GeneralFunctions.ArrayCellId(cellId_SoPhatSinh_ct40);

	private int[] CellId_SoLuyKe_ct29 => GeneralFunctions.ArrayCellId(cellId_SoLuyKe_ct29);

	private int[] CellId_SoLuyKe_ct30 => GeneralFunctions.ArrayCellId(cellId_SoLuyKe_ct30);

	private List<int[]> lstIdxRowsRelate_SoLuyKe_Ct30
	{
		get
		{
			List<int[]> list = new List<int[]>();
			list.Add(CellId_SoLuyKe_ct29);
			list.Add(CellId_SoPhatSinh_ct29);
			return list;
		}
	}

	private int[] CellId_SoLuyKe_ct31 => GeneralFunctions.ArrayCellId(cellId_SoLuyKe_ct31);

	private int[] CellId_SoLuyKe_ct32 => GeneralFunctions.ArrayCellId(cellId_SoLuyKe_ct32);

	private List<int[]> lstIdxRowsRelate_SoLuyKe_Ct32
	{
		get
		{
			List<int[]> list = new List<int[]>();
			list.Add(CellId_SoLuyKe_ct31);
			list.Add(CellId_SoPhatSinh_ct31);
			return list;
		}
	}

	private int[] CellId_SoLuyKe_ct33 => GeneralFunctions.ArrayCellId(cellId_SoLuyKe_ct33);

	private int[] CellId_SoLuyKe_ct34 => GeneralFunctions.ArrayCellId(cellId_SoLuyKe_ct34);

	private int[] CellId_SoLuyKe_ct35 => GeneralFunctions.ArrayCellId(cellId_SoLuyKe_ct35);

	private int[] CellId_SoLuyKe_ct36 => GeneralFunctions.ArrayCellId(cellId_SoLuyKe_ct36);

	private int[] CellId_SoLuyKe_ct37 => GeneralFunctions.ArrayCellId(cellId_SoLuyKe_ct37);

	private int[] CellId_SoLuyKe_ct38 => GeneralFunctions.ArrayCellId(cellId_SoLuyKe_ct38);

	private int[] CellId_SoLuyKe_ct39 => GeneralFunctions.ArrayCellId(cellId_SoLuyKe_ct39);

	private int[] CellId_SoLuyKe_ct40 => GeneralFunctions.ArrayCellId(cellId_SoLuyKe_ct40);

	private List<int[]> lstIdxRowsRelate_SoLuyKe_Ct40
	{
		get
		{
			List<int[]> list = new List<int[]>();
			list.Add(CellId_SoLuyKe_ct39);
			list.Add(CellId_SoPhatSinh_ct39);
			return list;
		}
	}

	private int[] CellId_SoLuyKeKT_ct29 => GeneralFunctions.ArrayCellId(cellId_SoLuyKeKT_ct29);

	private int[] CellId_SoLuyKeKT_ct30 => GeneralFunctions.ArrayCellId(cellId_SoLuyKeKT_ct30);

	private int[] CellId_SoLuyKeKT_ct31 => GeneralFunctions.ArrayCellId(cellId_SoLuyKeKT_ct31);

	private int[] CellId_SoLuyKeKT_ct32 => GeneralFunctions.ArrayCellId(cellId_SoLuyKeKT_ct32);

	private int[] CellId_SoLuyKeKT_ct33 => GeneralFunctions.ArrayCellId(cellId_SoLuyKeKT_ct33);

	private int[] CellId_SoLuyKeKT_ct34 => GeneralFunctions.ArrayCellId(cellId_SoLuyKeKT_ct34);

	private int[] CellId_SoLuyKeKT_ct35 => GeneralFunctions.ArrayCellId(cellId_SoLuyKeKT_ct35);

	private int[] CellId_SoLuyKeKT_ct36 => GeneralFunctions.ArrayCellId(cellId_SoLuyKeKT_ct36);

	private int[] CellId_SoLuyKeKT_ct37 => GeneralFunctions.ArrayCellId(cellId_SoLuyKeKT_ct37);

	private int[] CellId_SoLuyKeKT_ct38 => GeneralFunctions.ArrayCellId(cellId_SoLuyKeKT_ct38);

	private int[] CellId_SoLuyKeKT_ct39 => GeneralFunctions.ArrayCellId(cellId_SoLuyKeKT_ct39);

	private int[] CellId_SoLuyKeKT_ct40 => GeneralFunctions.ArrayCellId(cellId_SoLuyKeKT_ct40);

	public Tncn06Nckt216()
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			XuLyThongTin_ToChucKhaiThay(sView);
			FillSoLuyKeKT_ToSheet(sView, PeriodInfo, ToKhaiInfo, HeaderInfo);
			if (ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				DuaGiaTriTuDataFile(sView, ToKhaiInfo.SheetInfos[0]);
			}
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
			string text = HTKKFile.ColumnIndexToLetters((int)iCol) + "_" + iRow;
			if (text == cellId_ct12 || text == cellId_ct13)
			{
				XuLyThongTin_ToChucKhaiThay(winFpSpread.Sheets[0]);
				return;
			}
			List<string> list = new List<string>();
			list.Add(cellId_SoPhatSinh_ct29);
			list.Add(cellId_SoLuyKe_ct29);
			list.Add(cellId_SoPhatSinh_ct31);
			list.Add(cellId_SoLuyKe_ct31);
			list.Add(cellId_SoPhatSinh_ct33);
			list.Add(cellId_SoPhatSinh_ct35);
			list.Add(cellId_SoPhatSinh_ct37);
			list.Add(cellId_SoPhatSinh_ct38);
			list.Add(cellId_SoPhatSinh_ct39);
			list.Add(cellId_SoLuyKe_ct39);
			List<string> list2 = list;
			if (list2.Contains(text))
			{
				XuLyCellChange(winFpSpread.Sheets[0], text);
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
			bool result = true;
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink)
			{
				result = false;
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
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			int num = 0;
			string text = winFpSpread.Sheets[0].Cells[CellId_ct13[1], CellId_ct13[0]].Text;
			if (!HTKKFormater.isNullString(text))
			{
				if (!ValidateBase.CheckMaSoThue(text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, CellId_ct13[1], CellId_ct13[0], HTKKFile.GetMessage("0047"));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, CellId_ct13[1], CellId_ct13[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sView, CellId_ct13[1], CellId_ct13[0]);
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

	private void XuLyThongTin_ToChucKhaiThay(SheetView sViewTK)
	{
		try
		{
			string text = sViewTK.Cells[CellId_ct12[1], CellId_ct12[0]].Text;
			string text2 = sViewTK.Cells[CellId_ct13[1], CellId_ct13[0]].Text;
			if (string.IsNullOrEmpty(text) && string.IsNullOrEmpty(text2))
			{
				GeneralFunctions.ChangeValueCell(sViewTK, CellId_ct14[1], CellId_ct14[0], locked: true, "", "");
				GeneralFunctions.ChangeValueCell(sViewTK, CellId_ct15[1], CellId_ct15[0], locked: true, "", "");
				GeneralFunctions.ChangeValueCell(sViewTK, CellId_ct16[1], CellId_ct16[0], locked: true, "", "");
				GeneralFunctions.ChangeValueCell(sViewTK, CellId_ct17[1], CellId_ct17[0], locked: true, "", "");
				GeneralFunctions.ChangeValueCell(sViewTK, CellId_ct18[1], CellId_ct18[0], locked: true, "", "");
				GeneralFunctions.ChangeValueCell(sViewTK, CellId_ct19[1], CellId_ct19[0], locked: true, "", "");
			}
			else
			{
				sViewTK.Cells[CellId_ct14[1], CellId_ct14[0]].Locked = false;
				sViewTK.Cells[CellId_ct15[1], CellId_ct15[0]].Locked = false;
				sViewTK.Cells[CellId_ct16[1], CellId_ct16[0]].Locked = false;
				sViewTK.Cells[CellId_ct17[1], CellId_ct17[0]].Locked = false;
				sViewTK.Cells[CellId_ct18[1], CellId_ct18[0]].Locked = false;
				sViewTK.Cells[CellId_ct19[1], CellId_ct19[0]].Locked = false;
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void XuLyCellChange(SheetView sView, string cellChangeID)
	{
		try
		{
			if (cellId_SoPhatSinh_ct29 == cellChangeID)
			{
				sView.Cells[CellId_SoLuyKe_ct29[1], CellId_SoLuyKe_ct29[0]].Formula = formula_SoLuyKe_ct29;
				sView.Cells[CellId_SoPhatSinh_ct30[1], CellId_SoPhatSinh_ct30[0]].Formula = formula_SoPhatSinh_ct30;
				sView.Cells[CellId_SoLuyKe_ct30[1], CellId_SoLuyKe_ct30[0]].Formula = formula_SoLuyKe_ct30;
			}
			else if (cellId_SoLuyKe_ct29 == cellChangeID)
			{
				sView.Cells[CellId_SoLuyKe_ct29[1], CellId_SoLuyKe_ct29[0]].ResetFormula();
				sView.Cells[CellId_SoLuyKe_ct30[1], CellId_SoLuyKe_ct30[0]].Formula = formula_SoLuyKe_ct30;
			}
			else if (cellId_SoPhatSinh_ct31 == cellChangeID)
			{
				sView.Cells[CellId_SoLuyKe_ct31[1], CellId_SoLuyKe_ct31[0]].Formula = formula_SoLuyKe_ct31;
				sView.Cells[CellId_SoPhatSinh_ct32[1], CellId_SoPhatSinh_ct32[0]].Formula = formula_SoPhatSinh_ct32;
				sView.Cells[CellId_SoLuyKe_ct32[1], CellId_SoLuyKe_ct32[0]].Formula = formula_SoLuyKe_ct32;
			}
			else if (cellId_SoLuyKe_ct31 == cellChangeID)
			{
				sView.Cells[CellId_SoLuyKe_ct31[1], CellId_SoLuyKe_ct31[0]].ResetFormula();
				sView.Cells[CellId_SoLuyKe_ct32[1], CellId_SoLuyKe_ct32[0]].Formula = formula_SoLuyKe_ct32;
			}
			else if (cellId_SoPhatSinh_ct33 == cellChangeID)
			{
				sView.Cells[CellId_SoLuyKe_ct33[1], CellId_SoLuyKe_ct33[0]].Formula = formula_SoLuyKe_ct33;
			}
			else if (cellId_SoPhatSinh_ct35 == cellChangeID)
			{
				sView.Cells[CellId_SoLuyKe_ct35[1], CellId_SoLuyKe_ct35[0]].Formula = formula_SoLuyKe_ct35;
			}
			else if (cellId_SoPhatSinh_ct37 == cellChangeID)
			{
				sView.Cells[CellId_SoLuyKe_ct37[1], CellId_SoLuyKe_ct37[0]].Formula = formula_SoLuyKe_ct37;
			}
			else if (cellId_SoPhatSinh_ct38 == cellChangeID)
			{
				sView.Cells[CellId_SoLuyKe_ct38[1], CellId_SoLuyKe_ct38[0]].Formula = formula_SoLuyKe_ct38;
			}
			else if (cellId_SoPhatSinh_ct39 == cellChangeID)
			{
				sView.Cells[CellId_SoLuyKe_ct39[1], CellId_SoLuyKe_ct39[0]].Formula = formula_SoLuyKe_ct39;
				sView.Cells[CellId_SoPhatSinh_ct40[1], CellId_SoPhatSinh_ct40[0]].Formula = formula_SoPhatSinh_ct40;
				sView.Cells[CellId_SoLuyKe_ct40[1], CellId_SoLuyKe_ct40[0]].Formula = formula_SoLuyKe_ct40;
			}
			else if (cellId_SoLuyKe_ct39 == cellChangeID)
			{
				sView.Cells[CellId_SoLuyKe_ct40[1], CellId_SoLuyKe_ct40[0]].Formula = formula_SoLuyKe_ct40;
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private void DuaGiaTriTuDataFile(SheetView sView, HTKKCMCSheetInfo sInfoTK)
	{
		try
		{
			double num = Convert.ToDouble(sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == path_SoPhatSinh_ct30).Value);
			double num2 = Convert.ToDouble(sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == path_SoPhatSinh_ct32).Value);
			double num3 = Convert.ToDouble(sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == path_SoPhatSinh_ct40).Value);
			double num4 = Convert.ToDouble(sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == path_SoLuyKe_ct29).Value);
			double num5 = Convert.ToDouble(sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == path_SoLuyKe_ct30).Value);
			double num6 = Convert.ToDouble(sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == path_SoLuyKe_ct31).Value);
			double num7 = Convert.ToDouble(sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == path_SoLuyKe_ct32).Value);
			double num8 = Convert.ToDouble(sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == path_SoLuyKe_ct33).Value);
			double num9 = Convert.ToDouble(sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == path_SoLuyKe_ct35).Value);
			double num10 = Convert.ToDouble(sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == path_SoLuyKe_ct37).Value);
			double num11 = Convert.ToDouble(sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == path_SoLuyKe_ct38).Value);
			double num12 = Convert.ToDouble(sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == path_SoLuyKe_ct39).Value);
			double num13 = Convert.ToDouble(sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == path_SoLuyKe_ct40).Value);
			sView.Cells[CellId_SoPhatSinh_ct30[1], CellId_SoPhatSinh_ct30[0]].Value = num;
			sView.Cells[CellId_SoPhatSinh_ct32[1], CellId_SoPhatSinh_ct32[0]].Value = num2;
			sView.Cells[CellId_SoPhatSinh_ct40[1], CellId_SoPhatSinh_ct40[0]].Value = num3;
			sView.Cells[CellId_SoLuyKe_ct29[1], CellId_SoLuyKe_ct29[0]].Value = num4;
			sView.Cells[CellId_SoLuyKe_ct30[1], CellId_SoLuyKe_ct30[0]].Value = num5;
			sView.Cells[CellId_SoLuyKe_ct31[1], CellId_SoLuyKe_ct31[0]].Value = num6;
			sView.Cells[CellId_SoLuyKe_ct32[1], CellId_SoLuyKe_ct32[0]].Value = num7;
			sView.Cells[CellId_SoLuyKe_ct33[1], CellId_SoLuyKe_ct33[0]].Value = num8;
			sView.Cells[CellId_SoLuyKe_ct35[1], CellId_SoLuyKe_ct35[0]].Value = num9;
			sView.Cells[CellId_SoLuyKe_ct37[1], CellId_SoLuyKe_ct37[0]].Value = num10;
			sView.Cells[CellId_SoLuyKe_ct38[1], CellId_SoLuyKe_ct38[0]].Value = num11;
			sView.Cells[CellId_SoLuyKe_ct39[1], CellId_SoLuyKe_ct39[0]].Value = num12;
			sView.Cells[CellId_SoLuyKe_ct40[1], CellId_SoLuyKe_ct40[0]].Value = num13;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private Dictionary<string, string> GetPreperiodValue(HTKKPeriodInfo periodInfo, HTKKCMCToKhaiInfo toKhaiInfo, HTKKHeaderInfo headerInfo, List<string> lstCellId)
	{
		try
		{
			Dictionary<string, string> result = new Dictionary<string, string>();
			List<string> preperiodFileName_ThangAndQuy = GeneralFunctions.GetPreperiodFileName_ThangAndQuy(periodInfo, toKhaiInfo, 0);
			List<string> preperiodFileName_ThangAndQuy2 = GeneralFunctions.GetPreperiodFileName_ThangAndQuy(periodInfo, toKhaiInfo, 1);
			string text = HTKKConfig.DATAFILES + "\\" + headerInfo.mst;
			DirectoryInfo directoryInfo = new DirectoryInfo(text);
			int num = toKhaiInfo.DataFileNameSaveExtend.Split(new char[1] { '_' }).Length - 1;
			string empty = string.Empty;
			if (preperiodFileName_ThangAndQuy2 != null && preperiodFileName_ThangAndQuy2.Count > 0)
			{
				empty = text + "\\" + preperiodFileName_ThangAndQuy2[0] + ".xml";
				if (preperiodFileName_ThangAndQuy != null && preperiodFileName_ThangAndQuy.Count > 0)
				{
					empty = text + "\\" + preperiodFileName_ThangAndQuy[0] + ".xml";
				}
				if (!string.IsNullOrEmpty(empty) && File.Exists(empty))
				{
					result = GeneralFunctions.GetValueInXml(empty, lstCellId);
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void FillSoLuyKeKT_ToSheet(SheetView sViewTK, HTKKPeriodInfo periodInfo, HTKKCMCToKhaiInfo toKhaiInfo, HTKKHeaderInfo headerInfo)
	{
		try
		{
			List<string> list = new List<string>();
			list.Add(cellId_SoLuyKe_ct29);
			list.Add(cellId_SoLuyKe_ct30);
			list.Add(cellId_SoLuyKe_ct31);
			list.Add(cellId_SoLuyKe_ct32);
			list.Add(cellId_SoLuyKe_ct33);
			list.Add(cellId_SoLuyKe_ct34);
			list.Add(cellId_SoLuyKe_ct35);
			list.Add(cellId_SoLuyKe_ct36);
			list.Add(cellId_SoLuyKe_ct37);
			list.Add(cellId_SoLuyKe_ct38);
			list.Add(cellId_SoLuyKe_ct39);
			list.Add(cellId_SoLuyKe_ct40);
			List<string> lstCellId = list;
			Dictionary<string, string> preperiodValue = GetPreperiodValue(periodInfo, toKhaiInfo, headerInfo, lstCellId);
			if (preperiodValue != null && preperiodValue.Count > 0)
			{
				sViewTK.Cells[CellId_SoLuyKeKT_ct29[1], CellId_SoLuyKeKT_ct29[0]].Value = preperiodValue[cellId_SoLuyKe_ct29];
				sViewTK.Cells[CellId_SoLuyKeKT_ct30[1], CellId_SoLuyKeKT_ct30[0]].Value = preperiodValue[cellId_SoLuyKe_ct30];
				sViewTK.Cells[CellId_SoLuyKeKT_ct31[1], CellId_SoLuyKeKT_ct31[0]].Value = preperiodValue[cellId_SoLuyKe_ct31];
				sViewTK.Cells[CellId_SoLuyKeKT_ct32[1], CellId_SoLuyKeKT_ct32[0]].Value = preperiodValue[cellId_SoLuyKe_ct32];
				sViewTK.Cells[CellId_SoLuyKeKT_ct33[1], CellId_SoLuyKeKT_ct33[0]].Value = preperiodValue[cellId_SoLuyKe_ct33];
				sViewTK.Cells[CellId_SoLuyKeKT_ct34[1], CellId_SoLuyKeKT_ct34[0]].Value = preperiodValue[cellId_SoLuyKe_ct34];
				sViewTK.Cells[CellId_SoLuyKeKT_ct35[1], CellId_SoLuyKeKT_ct35[0]].Value = preperiodValue[cellId_SoLuyKe_ct35];
				sViewTK.Cells[CellId_SoLuyKeKT_ct36[1], CellId_SoLuyKeKT_ct36[0]].Value = preperiodValue[cellId_SoLuyKe_ct36];
				sViewTK.Cells[CellId_SoLuyKeKT_ct37[1], CellId_SoLuyKeKT_ct37[0]].Value = preperiodValue[cellId_SoLuyKe_ct37];
				sViewTK.Cells[CellId_SoLuyKeKT_ct38[1], CellId_SoLuyKeKT_ct38[0]].Value = preperiodValue[cellId_SoLuyKe_ct38];
				sViewTK.Cells[CellId_SoLuyKeKT_ct39[1], CellId_SoLuyKeKT_ct39[0]].Value = preperiodValue[cellId_SoLuyKe_ct39];
				sViewTK.Cells[CellId_SoLuyKeKT_ct40[1], CellId_SoLuyKeKT_ct40[0]].Value = preperiodValue[cellId_SoLuyKe_ct40];
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
