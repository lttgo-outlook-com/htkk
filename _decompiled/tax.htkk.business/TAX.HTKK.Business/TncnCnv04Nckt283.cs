using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnCnv04Nckt283
{
	private int[] cellDanhThue2Lan = GeneralFunctions.ArrayCellId("C_22");

	private int[] cellChuyenNhuongVonGop = GeneralFunctions.ArrayCellId("C_24");

	private int[] cellChuyenNhuongChungKhoan = GeneralFunctions.ArrayCellId("C_26");

	private int[] cellCt12 = GeneralFunctions.ArrayCellId("L_66");

	private int[] cellCt13 = GeneralFunctions.ArrayCellId("H_68");

	private int[] cellCt14 = GeneralFunctions.ArrayCellId("H_76");

	private int[] cellCt15 = GeneralFunctions.ArrayCellId("H_72");

	private int[] cellCt15_1 = GeneralFunctions.ArrayCellId("H_74");

	private int[] cellCt16 = GeneralFunctions.ArrayCellId("H_70");

	private int[] cellCt17 = GeneralFunctions.ArrayCellId("F_78");

	private int[] cellCt18 = GeneralFunctions.ArrayCellId("L_78");

	private int[] cellCt19 = GeneralFunctions.ArrayCellId("S_78");

	private int[] cellCt42 = GeneralFunctions.ArrayCellId("U_135");

	private int[] cellCt44 = GeneralFunctions.ArrayCellId("U_138");

	private int[] cellCt45 = GeneralFunctions.ArrayCellId("U_139");

	private int[] cellCt45Formular = GeneralFunctions.ArrayCellId("AI_139");

	private int[] cellCt37 = GeneralFunctions.ArrayCellId("U_130");

	private int[] cellCt39 = GeneralFunctions.ArrayCellId("U_132");

	private int[] cellCt41 = GeneralFunctions.ArrayCellId("U_134");

	private int[] cellCt41Formular = GeneralFunctions.ArrayCellId("AI_134");

	private int[] cellCt46 = GeneralFunctions.ArrayCellId("U_140");

	private int[] cellTemPhuLuc = GeneralFunctions.ArrayCellId("AI_24");

	private HTKKHeaderInfo _HeaderInfo = null;

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			_HeaderInfo = HeaderInfo;
			_ToKhaiInfo = ToKhaiInfo;
			int[] array = GeneralFunctions.ArrayCellId("C_22");
			int[] array2 = GeneralFunctions.ArrayCellId("C_24");
			int[] array3 = GeneralFunctions.ArrayCellId("C_26");
			int[] array4 = GeneralFunctions.ArrayCellId("U_105");
			int[] array5 = GeneralFunctions.ArrayCellId("U_110");
			int[] array6 = GeneralFunctions.ArrayCellId("H_49");
			int[] array7 = GeneralFunctions.ArrayCellId("H_45");
			int[] array8 = GeneralFunctions.ArrayCellId("H_43");
			int[] array9 = GeneralFunctions.ArrayCellId("F_55");
			int[] array10 = GeneralFunctions.ArrayCellId("L_55");
			int[] array11 = GeneralFunctions.ArrayCellId("S_55");
			LockUnlockDanhThue2Lan(sView, ToKhaiInfo);
			bool isDefault = ((PeriodInfo.KieuKy != "D" && string.IsNullOrEmpty(sView.Cells[cellCt13[1], cellCt13[0]].Text.Trim())) ? true : false);
			ThayDoiHinhThucKhaiThay(sView, PeriodInfo, isDefault);
			GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, sView.FpSpread);
			sView.Cells[cellTemPhuLuc[1], cellTemPhuLuc[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			bool flag = ConvertHelpers.ConvertToBool(sView.Cells[array2[1], array2[0]].Value);
			bool flag2 = ConvertHelpers.ConvertToBool(sView.Cells[array3[1], array3[0]].Value);
			if (flag || flag2)
			{
				LockTable(sView, ToKhaiInfo, isLock: true);
			}
			else
			{
				LockTable(sView, ToKhaiInfo, isLock: false);
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
			SheetView sheetView = winFpSpread.Sheets[0];
			if (iRow == cellDanhThue2Lan[1] + 1)
			{
				LockUnlockDanhThue2Lan(sheetView, ToKhaiInfo);
			}
			bool flag = Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongVonGop[1], cellChuyenNhuongVonGop[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongChungKhoan[1], cellChuyenNhuongChungKhoan[0]].Value);
			bool flag3 = Convert.ToBoolean(sheetView.Cells[cellDanhThue2Lan[1], cellDanhThue2Lan[0]].Value);
			if (iRow == cellDanhThue2Lan[1] + 1 && flag3)
			{
				sheetView.Cells[cellChuyenNhuongVonGop[1], cellChuyenNhuongVonGop[0]].Value = false;
				sheetView.Cells[cellChuyenNhuongChungKhoan[1], cellChuyenNhuongChungKhoan[0]].Value = false;
			}
			if (iRow == cellChuyenNhuongVonGop[1] + 1 && flag)
			{
				sheetView.Cells[cellChuyenNhuongChungKhoan[1], cellChuyenNhuongChungKhoan[0]].Value = false;
				sheetView.Cells[cellDanhThue2Lan[1], cellDanhThue2Lan[0]].Value = false;
			}
			if (iRow == cellChuyenNhuongChungKhoan[1] + 1 && flag2)
			{
				sheetView.Cells[cellChuyenNhuongVonGop[1], cellChuyenNhuongVonGop[0]].Value = false;
				sheetView.Cells[cellDanhThue2Lan[1], cellDanhThue2Lan[0]].Value = false;
			}
			if (iRow == cellDanhThue2Lan[1] + 1 || iRow == cellChuyenNhuongVonGop[1] + 1 || iRow == cellChuyenNhuongChungKhoan[1] + 1)
			{
				ThayDoiHinhThucKhaiThay(sheetView, PeriodInfo, isDefault: true);
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			}
			if (iRow == cellCt44[1] + 1)
			{
				sheetView.Cells[cellCt45[1], cellCt45[0]].Value = sheetView.Cells[cellCt45Formular[1], cellCt45Formular[0]].Value;
			}
			if (iRow >= cellCt37[1] + 1 && iRow <= cellCt39[1] + 1)
			{
				sheetView.Cells[cellCt41[1], cellCt41[0]].Value = sheetView.Cells[cellCt41Formular[1], cellCt41Formular[0]].Value;
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
			SheetView sView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_38"));
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_39"));
			string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_60"));
			bool flag = Convert.ToBoolean(winFpSpread.Sheets[0].Cells[cellChuyenNhuongVonGop[1], cellChuyenNhuongVonGop[0]].Value);
			bool flag2 = Convert.ToBoolean(winFpSpread.Sheets[0].Cells[cellChuyenNhuongChungKhoan[1], cellChuyenNhuongChungKhoan[0]].Value);
			if ((flag || flag2) && ToKhaiInfo.SheetInfos[1] == null)
			{
				HTKKFile.SetNodeBackGroundPink(sView, cellChuyenNhuongVonGop[1], cellChuyenNhuongVonGop[0], message3);
				HTKKFile.SetNodeBackGroundPink(sView, cellChuyenNhuongChungKhoan[1], cellChuyenNhuongChungKhoan[0], message3);
				value = 2;
			}
			else if (TTinChung.kieuKy != "D" && !flag && !flag2)
			{
				HTKKFile.SetNodeBackGroundPink(sView, cellChuyenNhuongVonGop[1], cellChuyenNhuongVonGop[0], message2);
				HTKKFile.SetNodeBackGroundPink(sView, cellChuyenNhuongChungKhoan[1], cellChuyenNhuongChungKhoan[0], message2);
				value = 2;
			}
			else if (!flag && !flag2 && ToKhaiInfo.SheetInfos[1] != null)
			{
				HTKKFile.SetNodeBackGroundPink(sView, cellChuyenNhuongVonGop[1], cellChuyenNhuongVonGop[0], message);
				HTKKFile.SetNodeBackGroundPink(sView, cellChuyenNhuongChungKhoan[1], cellChuyenNhuongChungKhoan[0], message);
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sView, cellChuyenNhuongVonGop[1], cellChuyenNhuongVonGop[0]);
				HTKKFile.RemoveValidate(sView, cellChuyenNhuongChungKhoan[1], cellChuyenNhuongChungKhoan[0]);
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

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("I");
			int num4 = HTKKFile.ColumnLettersToIndex("J");
			int[] array = GeneralFunctions.ArrayCellId("Q_36");
			int[] array2 = GeneralFunctions.ArrayCellId("R_36");
			int[] array3 = GeneralFunctions.ArrayCellId("S_36");
			int[] array4 = GeneralFunctions.ArrayCellId("O_56");
			int[] array5 = GeneralFunctions.ArrayCellId("P_56");
			int[] array6 = GeneralFunctions.ArrayCellId("Q_56");
			int[] array7 = GeneralFunctions.ArrayCellId("U_104");
			int[] array8 = GeneralFunctions.ArrayCellId("U_105");
			int[] array9 = GeneralFunctions.ArrayCellId("U_106");
			int[] array10 = GeneralFunctions.ArrayCellId("U_109");
			int[] array11 = GeneralFunctions.ArrayCellId("U_110");
			int[] array12 = GeneralFunctions.ArrayCellId("U_111");
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_22"));
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13"));
			string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_21"));
			string message4 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23"));
			string message5 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_27"));
			string message6 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_32"));
			string message7 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_33"));
			string message8 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_34"));
			string message9 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_35"));
			string message10 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_36"));
			string message11 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_37"));
			string message12 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_38"));
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[1];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
			}
			switch (value)
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
		}
	}

	public void LockUnlockDanhThue2Lan(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		int num = hTKKCMCSheetInfo.Tables[0].RowsValues.Count - 1;
		bool flag = Convert.ToBoolean(sView.Cells[cellChuyenNhuongVonGop[1], cellChuyenNhuongVonGop[0]].Value);
		bool flag2 = Convert.ToBoolean(sView.Cells[cellChuyenNhuongChungKhoan[1], cellChuyenNhuongChungKhoan[0]].Value);
		if (Convert.ToBoolean(sView.Cells[cellDanhThue2Lan[1], cellDanhThue2Lan[0]].Value) || flag || flag2)
		{
			GeneralFunctions.LockUnLockCol(sView, cellCt42[1] + num + 1, cellCt42[0] + 1, locked: false, deleteData: false);
			GeneralFunctions.LockUnLockCol(sView, cellCt46[1] + num + 1, cellCt46[0] + 1, locked: false, deleteData: false);
			return;
		}
		GeneralFunctions.LockUnLockCol(sView, cellCt42[1] + num + 1, cellCt42[0] + 1, locked: true, deleteData: false);
		GeneralFunctions.LockUnLockCol(sView, cellCt46[1] + num + 1, cellCt46[0] + 1, locked: true, deleteData: false);
		sView.Cells[cellCt42[1] + num, cellCt42[0]].Value = 0;
		sView.Cells[cellCt46[1] + num, cellCt46[0]].Value = 0;
	}

	public void ThayDoiHinhThucKhaiThay(SheetView sView, HTKKPeriodInfo PeriodInfo, bool isDefault)
	{
		bool flag = Convert.ToBoolean(sView.Cells[cellChuyenNhuongVonGop[1], cellChuyenNhuongVonGop[0]].Value);
		bool flag2 = Convert.ToBoolean(sView.Cells[cellChuyenNhuongChungKhoan[1], cellChuyenNhuongChungKhoan[0]].Value);
		if (flag || flag2)
		{
			if (isDefault)
			{
				sView.Cells[cellCt12[1], cellCt12[0]].Value = _HeaderInfo.tenNNT;
				sView.Cells[cellCt13[1], cellCt13[0]].Value = _HeaderInfo.mst;
				sView.Cells[cellCt14[1], cellCt14[0]].Value = _HeaderInfo.dchiNNT;
				sView.Cells[cellCt15_1[1], cellCt15_1[0]].Value = _HeaderInfo.tenXaPhuongNNT;
				sView.Cells[cellCt16[1], cellCt16[0]].Value = _HeaderInfo.tenTinhNNT;
				sView.Cells[cellCt17[1], cellCt17[0]].Value = _HeaderInfo.dthoaiNNT;
				sView.Cells[cellCt18[1], cellCt18[0]].Value = _HeaderInfo.faxNNT;
				sView.Cells[cellCt19[1], cellCt19[0]].Value = _HeaderInfo.emailNNT;
			}
		}
		else if (isDefault)
		{
			sView.Cells[cellCt12[1], cellCt12[0]].Value = "";
			sView.Cells[cellCt13[1], cellCt13[0]].Value = "";
			sView.Cells[cellCt14[1], cellCt14[0]].Value = "";
			sView.Cells[cellCt15[1], cellCt15[0]].Value = "";
			sView.Cells[cellCt15_1[1], cellCt15_1[0]].Value = "";
			sView.Cells[cellCt16[1], cellCt16[0]].Value = "";
			sView.Cells[cellCt17[1], cellCt17[0]].Value = "";
			sView.Cells[cellCt18[1], cellCt18[0]].Value = "";
			sView.Cells[cellCt19[1], cellCt19[0]].Value = "";
		}
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellTemPhuLuc[1], cellTemPhuLuc[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			GeneralFunctions.checkLockUnLockCells(_ToKhaiInfo, winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_1(FpSpread winFpSpread, int actionType)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellTemPhuLuc[1], cellTemPhuLuc[0]].Value = ((_ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			GeneralFunctions.checkLockUnLockCells(_ToKhaiInfo, winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void LockTable(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, bool isLock)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = hTKKCMCSheetInfo.Tables[0].RowsValues.Count - 1;
			int dataLocationTemplateRowStart = hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart;
			if (num > 0 && isLock)
			{
				if (hTKKCMCSheetInfo.Tables[0].RowsValues.Count > 1)
				{
					GeneralFunctions.RemoveDataRows(sView, 0, dataLocationTemplateRowStart - 1, hTKKCMCSheetInfo.Tables[0].RowsValues.Count - 1, hTKKCMCSheetInfo);
				}
				{
					foreach (HTKKCMCTableCellInfo cell in hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells)
					{
						if (!(cell.Path == "stt"))
						{
							sView.Cells[dataLocationTemplateRowStart - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
							sView.Cells[dataLocationTemplateRowStart - 1, cell.CellTemplate.ColIndex - 1].Locked = isLock;
						}
					}
					return;
				}
			}
			foreach (HTKKCMCTableCellInfo cell2 in hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells)
			{
				if (!(cell2.Path == "stt"))
				{
					sView.Cells[dataLocationTemplateRowStart - 1, cell2.CellTemplate.ColIndex - 1].Value = cell2.DefaultValue;
					sView.Cells[dataLocationTemplateRowStart - 1, cell2.CellTemplate.ColIndex - 1].Locked = isLock;
				}
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
			SheetView sheetView = winFpSpread.Sheets[0];
			if ((idxRow == cellChuyenNhuongVonGop[1] && idxCol == cellChuyenNhuongVonGop[0]) || (idxRow == cellChuyenNhuongChungKhoan[1] && idxCol == cellChuyenNhuongChungKhoan[0]))
			{
				bool isLock = ConvertHelpers.ConvertToBool(sheetView.Cells[idxRow, idxCol].Value);
				LockTable(sheetView, ToKhaiInfo, isLock);
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
			GeneralFunctions.FillDataMstCu(xmlNode);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
