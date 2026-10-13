using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnQtt02TT80Nckt273
{
	private string MA_COLOR_LOCKED = "#e8f5ff";

	private string MA_COLOR_UNLOCK = "#FFFFFF";

	private Dictionary<string, string> dicDMNganHangTmp = new Dictionary<string, string>();

	private int[] cellChenhLechSoThang = GeneralFunctions.ArrayCellId("W_17");

	private int[] cellSoTB = GeneralFunctions.ArrayCellId("F_10");

	private int[] cellNgayTB = GeneralFunctions.ArrayCellId("N_10");

	private int[] cellTenQuyetToan = GeneralFunctions.ArrayCellId("F_11");

	private int[] cellMaQuyetToan = GeneralFunctions.ArrayCellId("AA_12");

	private int[] cellChuyenKhoan = GeneralFunctions.ArrayCellId("D_149");

	private int[] cellTenChuTK = GeneralFunctions.ArrayCellId("H_149");

	private int[] cellTienMat = GeneralFunctions.ArrayCellId("D_153");

	private int[] cellComboTinh = GeneralFunctions.ArrayCellId("G_151");

	private int[] cellMaTinh = GeneralFunctions.ArrayCellId("AA_151");

	private int[] cellComboNH = GeneralFunctions.ArrayCellId("O_151");

	private int[] cell29 = GeneralFunctions.ArrayCellId("M_99");

	private int[] cell34 = GeneralFunctions.ArrayCellId("M_109");

	private int[] cell35 = GeneralFunctions.ArrayCellId("M_110");

	private int[] cell40 = GeneralFunctions.ArrayCellId("M_115");

	private int[] cell41 = GeneralFunctions.ArrayCellId("M_116");

	private int[] cell44 = GeneralFunctions.ArrayCellId("M_124");

	private int[] cell46 = GeneralFunctions.ArrayCellId("M_126");

	private int[] cell47 = GeneralFunctions.ArrayCellId("M_127");

	private int col5 = HTKKFile.ColumnLettersToIndex("J");

	private int col7 = HTKKFile.ColumnLettersToIndex("M");

	private int col8_1 = HTKKFile.ColumnLettersToIndex("O");

	private int col8_1Ma = HTKKFile.ColumnLettersToIndex("AC");

	private int col9 = HTKKFile.ColumnLettersToIndex("Q");

	private int col9Ma = HTKKFile.ColumnLettersToIndex("AE");

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sView.Cells[cellTenQuyetToan[1], cellTenQuyetToan[0]].Value = PeriodInfo.TenQuyetToan;
				sView.Cells[cellMaQuyetToan[1], cellMaQuyetToan[0]].Value = PeriodInfo.MaQuyetToan;
			}
			LuonKhoaCot5(sView, ToKhaiInfo);
			ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
			comboBoxCellType = (ComboBoxCellType)sView.Cells[cellComboTinh[1], cellComboTinh[0]].CellType;
			comboBoxCellType.Editable = true;
			sView.Cells[cellComboTinh[1], cellComboTinh[0]].CellType = comboBoxCellType;
			if (PeriodInfo.LoaiToKhai == "C")
			{
				sView.Cells[cellSoTB[1], cellSoTB[0]].Value = "";
				sView.Cells[cellNgayTB[1], cellNgayTB[0]].Value = "";
				return;
			}
			if (!string.IsNullOrEmpty(PeriodInfo.SoTB))
			{
				sView.Cells[cellSoTB[1], cellSoTB[0]].Value = PeriodInfo.SoTB;
			}
			else
			{
				sView.Cells[cellSoTB[1], cellSoTB[0]].Value = "";
			}
			if ("  /  /" != PeriodInfo.NgayTb)
			{
				sView.Cells[cellNgayTB[1], cellNgayTB[0]].Value = PeriodInfo.NgayTb;
			}
			else
			{
				sView.Cells[cellNgayTB[1], cellNgayTB[0]].Value = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = TinhSoThang(PeriodInfo.TuThang, PeriodInfo.DenThang);
			sheetView.Cells[cellChenhLechSoThang[1], cellChenhLechSoThang[0]].Value = num;
			TinhCot16VaCot17TuPL_02_1(winFpSpread, ToKhaiInfo, PeriodInfo, -1);
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
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			string text;
			if ((int)iRow == cellComboNH[1] + 1 && (int)iCol == cellComboNH[0] + 1)
			{
				text = sheetView.Cells[cellMaTinh[1], cellMaTinh[0]].Text;
				GeneralFunctions.ResetCombobox(sheetView, cellComboNH[1], cellComboNH[0], GetDM(text));
			}
			if ((int)iRow == cellComboTinh[1] + 1 && (int)iCol == cellComboTinh[0] + 1)
			{
				Dictionary<string, string> dictionary = new Dictionary<string, string>();
				foreach (KeyValuePair<string, string> item in HTKKWinDictionaryDiaDanh.dicTinhTP)
				{
					dictionary.Add(item.Value, item.Key);
				}
				GeneralFunctions.ResetCombobox(sheetView, cellComboTinh[1], cellComboTinh[0], dictionary);
			}
			if ((int)iRow == cellChuyenKhoan[1] + 1 && (int)iCol == cellChuyenKhoan[0] + 1)
			{
				if (Convert.ToBoolean(sheetView.Cells[cellChuyenKhoan[1], cellChuyenKhoan[0]].Value))
				{
					sheetView.Cells[cellTienMat[1], cellTienMat[0]].Value = false;
				}
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			}
			if ((int)iRow == cellTienMat[1] + 1 && (int)iCol == cellTienMat[0] + 1)
			{
				if (Convert.ToBoolean(sheetView.Cells[cellTienMat[1], cellTienMat[0]].Value))
				{
					sheetView.Cells[cellChuyenKhoan[1], cellChuyenKhoan[0]].Value = false;
				}
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			}
			if (iCol == cell46[0] + 1 && iRow == cell46[1] + 1 && Convert.ToDouble(sheetView.Cells[cell46[1], cell46[0]].Value) == 0.0)
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			}
			if (iCol == cell44[0] + 1 && Convert.ToDouble(sheetView.Cells[cell44[1], cell44[0]].Value) == 0.0)
			{
				sheetView.Cells[cell46[1], cell46[0]].Value = 0;
				sheetView.Cells[cell47[1], cell47[0]].Value = 0;
			}
			if (iRow == cell46[1] + 1 || iRow == cell47[1] + 1 || (iRow >= cell34[1] + 1 && iRow <= cell41[1] + 1))
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
				LuonKhoaCot5(sheetView, ToKhaiInfo);
			}
			int num = GeneralFunctions.DetectIndexOfTable(sInfo, (int)iRow);
			if (iCol == col7 && num >= 0)
			{
				sheetView.Cells[(int)iRow - 1, col5 - 1].Value = sheetView.Cells[(int)iRow - 1, col7 - 1].Value;
			}
			if (iCol == col8_1)
			{
				string maCQTCapCuc = sheetView.Cells[(int)iRow - 1, col8_1Ma - 1].Text.Trim();
				ThayDoiTinhTheoCQT(maCQTCapCuc, sheetView, (int)iRow, col9, col9Ma);
			}
			if ((int)iRow != cellComboTinh[1] + 1 || (int)iCol != cellComboTinh[0] + 1)
			{
				return;
			}
			text = sheetView.Cells[cellMaTinh[1], cellMaTinh[0]].Text.Trim();
			dicDMNganHangTmp = GetDM(text);
			ComboBoxCellType comboBoxCellType = (ComboBoxCellType)sheetView.Cells[cellComboNH[1], cellComboNH[0]].CellType;
			string[] array = new string[dicDMNganHangTmp.Count];
			string[] array2 = new string[dicDMNganHangTmp.Count];
			int num2 = 0;
			foreach (KeyValuePair<string, string> item2 in dicDMNganHangTmp)
			{
				string[] array4;
				string[] array3 = (array4 = array);
				int num3 = num2;
				nint num4 = num3;
				array3[num3] = array4[num4] + item2.Value;
				string[] array5 = (array4 = array2);
				int num5 = num2;
				num4 = num5;
				array5[num5] = array4[num4] + item2.Key;
				num2++;
			}
			comboBoxCellType.Items = array;
			comboBoxCellType.ItemData = array2;
			sheetView.Cells[cellComboNH[1], cellComboNH[0]].CellType = comboBoxCellType;
			sheetView.Cells[cellComboNH[1], cellComboNH[0]].Text = "";
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("O");
			int num3 = HTKKFile.ColumnLettersToIndex("Q");
			if (iCol == num || iCol == num2 || iCol == num3)
			{
				TinhCot16VaCot17TuPL_02_1(winFpSpread, ToKhaiInfo, PeriodInfo, (int)iRow);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
		winFpSpread.Sheets[0].Cells[cell34[1], cell34[0]].Note = "";
		winFpSpread.Sheets[0].Cells[cell35[1], cell35[0]].Note = "";
		winFpSpread.Sheets[0].Cells[cell29[1], cell29[0]].Note = "";
		winFpSpread.Sheets[0].Cells[cell40[1], cell40[0]].Note = "";
		winFpSpread.Sheets[0].Cells[cell47[1], cell47[0]].Note = "";
		if ((int)iRow == cell34[1] + 1 && (int)iCol == cell34[0] + 1)
		{
			winFpSpread.Sheets[0].Cells[cell34[1], cell34[0]].NoteStyle = NoteStyle.StickyNote;
			winFpSpread.Sheets[0].Cells[cell34[1], cell34[0]].Note = xMLMessage.GetMessage("0640");
		}
		if ((int)iRow == cell35[1] + 1 && (int)iCol == cell35[0] + 1)
		{
			winFpSpread.Sheets[0].Cells[cell35[1], cell35[0]].NoteStyle = NoteStyle.StickyNote;
			winFpSpread.Sheets[0].Cells[cell35[1], cell35[0]].Note = xMLMessage.GetMessage("0641");
		}
		if ((int)iRow == cell29[1] + 1 && (int)iCol == cell29[0] + 1)
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			object value = sheetView.Cells[cell29[1], cell29[0]].Value;
			if (ConvertHelpers.ConvertToDouble(value) == 0.0)
			{
				winFpSpread.Sheets[0].Cells[cell29[1], cell29[0]].NoteStyle = NoteStyle.StickyNote;
				winFpSpread.Sheets[0].Cells[cell29[1], cell29[0]].Note = xMLMessage.GetMessage("0643");
			}
		}
		if ((int)iRow == cell40[1] + 1 && (int)iCol == cell40[0] + 1)
		{
			winFpSpread.Sheets[0].Cells[cell40[1], cell40[0]].NoteStyle = NoteStyle.StickyNote;
			winFpSpread.Sheets[0].Cells[cell40[1], cell40[0]].Note = xMLMessage.GetMessage("0644");
		}
		if ((int)iRow == cell47[1] + 1 && (int)iCol == cell47[0] + 1)
		{
			winFpSpread.Sheets[0].Cells[cell47[1], cell47[0]].NoteStyle = NoteStyle.StickyNote;
			winFpSpread.Sheets[0].Cells[cell47[1], cell47[0]].Note = xMLMessage.GetMessage("0645");
		}
	}

	public void EnterCell_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
		winFpSpread.Sheets[0].Cells[cell34[1], cell34[0]].Note = "";
		winFpSpread.Sheets[0].Cells[cell35[1], cell35[0]].Note = "";
		winFpSpread.Sheets[0].Cells[cell29[1], cell29[0]].Note = "";
		winFpSpread.Sheets[0].Cells[cell40[1], cell40[0]].Note = "";
		winFpSpread.Sheets[0].Cells[cell47[1], cell47[0]].Note = "";
		if ((int)iRow == cell34[1] + 1 && (int)iCol == cell34[0] + 1)
		{
			winFpSpread.Sheets[0].Cells[cell34[1], cell34[0]].NoteStyle = NoteStyle.StickyNote;
			winFpSpread.Sheets[0].Cells[cell34[1], cell34[0]].Note = xMLMessage.GetMessage("0640");
		}
		if ((int)iRow == cell35[1] + 1 && (int)iCol == cell35[0] + 1)
		{
			winFpSpread.Sheets[0].Cells[cell35[1], cell35[0]].NoteStyle = NoteStyle.StickyNote;
			winFpSpread.Sheets[0].Cells[cell35[1], cell35[0]].Note = xMLMessage.GetMessage("0641");
		}
		if ((int)iRow == cell29[1] + 1 && (int)iCol == cell29[0] + 1)
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			object value = sheetView.Cells[cell29[1], cell29[0]].Value;
			if (ConvertHelpers.ConvertToDouble(value) == 0.0)
			{
				winFpSpread.Sheets[0].Cells[cell29[1], cell29[0]].NoteStyle = NoteStyle.StickyNote;
				winFpSpread.Sheets[0].Cells[cell29[1], cell29[0]].Note = xMLMessage.GetMessage("0643");
			}
		}
		if ((int)iRow == cell40[1] + 1 && (int)iCol == cell40[0] + 1)
		{
			winFpSpread.Sheets[0].Cells[cell40[1], cell40[0]].NoteStyle = NoteStyle.StickyNote;
			winFpSpread.Sheets[0].Cells[cell40[1], cell40[0]].Note = xMLMessage.GetMessage("0644");
		}
		if ((int)iRow == cell47[1] + 1 && (int)iCol == cell47[0] + 1)
		{
			winFpSpread.Sheets[0].Cells[cell47[1], cell47[0]].NoteStyle = NoteStyle.StickyNote;
			winFpSpread.Sheets[0].Cells[cell47[1], cell47[0]].Note = xMLMessage.GetMessage("0645");
		}
		if ((int)iRow == cellComboTinh[1] + 1 && (int)iCol == cellComboTinh[0] + 1)
		{
			winFpSpread.Sheets[0].Cells[cellComboTinh[1], cellComboTinh[0]].NoteStyle = NoteStyle.StickyNote;
			winFpSpread.Sheets[0].Cells[cellComboTinh[1], cellComboTinh[0]].Note = xMLMessage.GetMessage("0647");
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			string text = "O_149~G_151~O_151";
			string text2 = "H_153~O_153~F_155~O_155~G_157";
			if (idxRow == cellChuyenKhoan[1] && idxCol == cellChuyenKhoan[0])
			{
				string[] array = text.Split(new char[1] { '~' });
				bool flag = Convert.ToBoolean(sheetView.Cells[cellChuyenKhoan[1], cellChuyenKhoan[0]].Value);
				if (flag)
				{
					sheetView.Cells[cellTienMat[1], cellTienMat[0]].Value = false;
					sheetView.Cells[cellTenChuTK[1], cellTenChuTK[0]].Value = GeneralFunctions.RemoveDiacritics(TTinChung.tenNNT).ToUpper();
				}
				else
				{
					sheetView.Cells[cellTenChuTK[1], cellTenChuTK[0]].Value = "";
				}
				string[] array2 = array;
				foreach (string cellId in array2)
				{
					int[] array3 = GeneralFunctions.ArrayCellId(cellId);
					sheetView.Cells[array3[1], array3[0]].Locked = !flag;
					sheetView.Cells[array3[1], array3[0]].Text = "";
					sheetView.Cells[array3[1], array3[0]].BackColor = ((!flag) ? ColorTranslator.FromHtml(MA_COLOR_LOCKED) : ColorTranslator.FromHtml(MA_COLOR_UNLOCK));
				}
			}
			if (idxRow == cellTienMat[1] && idxCol == cellTienMat[0])
			{
				string[] array4 = text2.Split(new char[1] { '~' });
				bool flag = Convert.ToBoolean(sheetView.Cells[cellTienMat[1], cellTienMat[0]].Value);
				if (flag)
				{
					sheetView.Cells[cellChuyenKhoan[1], cellChuyenKhoan[0]].Value = false;
					sheetView.Cells[cellTenChuTK[1], cellTenChuTK[0]].Value = "";
				}
				string[] array2 = array4;
				foreach (string cellId in array2)
				{
					int[] array3 = GeneralFunctions.ArrayCellId(cellId);
					sheetView.Cells[array3[1], array3[0]].Locked = !flag;
					sheetView.Cells[array3[1], array3[0]].Text = "";
					sheetView.Cells[array3[1], array3[0]].BackColor = ((!flag) ? ColorTranslator.FromHtml(MA_COLOR_LOCKED) : ColorTranslator.FromHtml(MA_COLOR_UNLOCK));
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
		checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
		checkPL2_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
		return true;
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			switch (value)
			{
			case 2:
				sView.FpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL2_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("E");
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_49"));
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num - 1);
				string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text))
				{
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCTableInfo, num, text);
					if (rowIndexInSViewByCompare.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, message3);
						value = 2;
					}
				}
				if (!string.IsNullOrEmpty(text) && !ValidateBase.CheckMaSoThue(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, message);
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[1].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			GeneralFunctions.fpSpReindexingSheetTemplate(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0]);
			GeneralFunctions.fpSpReindexingSheetTemplate(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ThayDoiTinhTheoCQT(string maCQTCapCuc, SheetView sView, int row, int colTenTinh, int colMaTinh)
	{
		try
		{
			if (maCQTCapCuc == "82500")
			{
				sView.Cells[row - 1, colTenTinh - 1].Locked = false;
			}
			else
			{
				if (string.IsNullOrEmpty(maCQTCapCuc))
				{
					return;
				}
				sView.Cells[row - 1, colTenTinh - 1].Value = "";
				sView.Cells[row - 1, colTenTinh - 1].Locked = false;
				sView.Cells[row - 1, colMaTinh - 1].Value = "";
				string maTinh = (HTKKWinDictionaryDiaDanh.dicTinhTP.ContainsValue(maCQTCapCuc.Substring(0, 3)) ? maCQTCapCuc.Substring(0, 3) : "");
				if (!string.IsNullOrEmpty(maCQTCapCuc))
				{
					KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
					sView.Cells[row - 1, colTenTinh - 1].Value = keyValuePair.Key;
					sView.Cells[row - 1, colMaTinh - 1].Value = maTinh;
					sView.Cells[row - 1, colTenTinh - 1].Locked = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LuonKhoaCot5(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[ToKhaiInfo.SheetInfos[0].Tables.Count - 1];
			int num = HTKKFile.ColumnLettersToIndex("J");
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Locked = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public int TinhSoThang(string tuThang, string denThang)
	{
		try
		{
			if (string.IsNullOrEmpty(tuThang) || string.IsNullOrEmpty(denThang))
			{
				return 0;
			}
			return (ConvertHelpers.ConvertToDateTime2(denThang, "MM/yyyy").Year - ConvertHelpers.ConvertToDateTime2(tuThang, "MM/yyyy").Year) * 12 + ConvertHelpers.ConvertToDateTime2(denThang, "MM/yyyy").Month - ConvertHelpers.ConvertToDateTime2(tuThang, "MM/yyyy").Month + 1;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return 0;
		}
	}

	public void TinhCot16VaCot17TuPL_02_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo, int row)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = HTKKFile.ColumnLettersToIndex("O");
			int num2 = HTKKFile.ColumnLettersToIndex("Q");
			int num3 = HTKKFile.ColumnLettersToIndex("AB");
			int num4 = HTKKFile.ColumnLettersToIndex("X");
			int num5 = HTKKFile.ColumnLettersToIndex("Y");
			DateTime t = ConvertHelpers.ConvertToDateTime2("01/" + PeriodInfo.TuThang, "dd/MM/yyyy");
			DateTime t2 = ConvertHelpers.ConvertToDateTime2("01/" + PeriodInfo.DenThang, "dd/MM/yyyy");
			if (row < 0)
			{
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
					string text2 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text.Trim();
					DateTime dateTime = ConvertHelpers.ConvertToDateTime2("01/" + text, "dd/MM/yyyy");
					DateTime dateTime2 = ConvertHelpers.ConvertToDateTime2("01/" + text2, "dd/MM/yyyy");
					if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2))
					{
						int num6 = TinhSoThang(text, text2);
						sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Value = num6;
						int num7 = ((DateTime.Compare(dateTime, t) < 0 || DateTime.Compare(t2, dateTime) < 0) ? 1 : 0);
						sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Value = num7;
						int num8 = ((DateTime.Compare(dateTime2, t) < 0 || DateTime.Compare(t2, dateTime2) < 0) ? 1 : 0);
						sheetView.Cells[dataLocationTemplateRowStart + i - 1, num5 - 1].Value = num8;
					}
				}
			}
			else
			{
				string text3 = sheetView.Cells[row - 1, num - 1].Text.Trim();
				string text = sheetView.Cells[row - 1, num2 - 1].Text.Trim();
				DateTime dateTime = ConvertHelpers.ConvertToDateTime2("01/" + text3, "dd/MM/yyyy");
				DateTime dateTime2 = ConvertHelpers.ConvertToDateTime2("01/" + text, "dd/MM/yyyy");
				if (!string.IsNullOrEmpty(text3) && !string.IsNullOrEmpty(text))
				{
					int num6 = TinhSoThang(text3, text);
					sheetView.Cells[row - 1, num3 - 1].Value = num6;
					int num7 = ((DateTime.Compare(dateTime, t) < 0 || DateTime.Compare(t2, dateTime) < 0) ? 1 : 0);
					sheetView.Cells[row - 1, num4 - 1].Value = num7;
					int num8 = ((DateTime.Compare(dateTime2, t) < 0 || DateTime.Compare(t2, dateTime2) < 0) ? 1 : 0);
					sheetView.Cells[row - 1, num5 - 1].Value = num8;
				}
				else
				{
					sheetView.Cells[row - 1, num3 - 1].Value = 0;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo.Tables[0].RowsValues.Count == 1)
			{
				TinhCot16VaCot17TuPL_02_1(winFpSpread, ToKhaiInfo, PeriodInfo, 38);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AB");
			SheetView sheetView = winFpSpread.Sheets[1];
			sheetView.Cells[37, num - 1].Value = 0;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private Dictionary<string, string> GetDM(string maTinh)
	{
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0024: Expected O, but got Unknown
		//IL_0063: Unknown result type (might be due to invalid IL or missing references)
		//IL_006a: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\DM_Tinh_NganHang.xml";
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["MaNH"]).Value;
				string value2 = ((XmlNode)val2.Attributes["TenNH"]).Value;
				string value3 = ((XmlNode)val2.Attributes["MaTinh"]).Value;
				if (maTinh == "" || maTinh == value3)
				{
					dictionary2.Add(value, value2);
				}
			}
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void KeyPress(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e, Keys keyData)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Invalid comparison between Unknown and I4
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			if ((int)keyData != 114)
			{
				return;
			}
			if (idxRow == cellComboNH[1] + 1 && idxCol == cellComboNH[0] + 1)
			{
				string text = sheetView.Cells[cellMaTinh[1], cellMaTinh[0]].Text;
				GeneralFunctions.ComboAutocomplete(sheetView, idxRow - 1, idxCol - 1, GetDM(text));
			}
			if (idxRow != cellComboTinh[1] + 1 || idxCol != cellComboTinh[0] + 1)
			{
				return;
			}
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			foreach (KeyValuePair<string, string> item in HTKKWinDictionaryDiaDanh.dicTinhTP)
			{
				dictionary.Add(item.Value, item.Key);
			}
			GeneralFunctions.ComboAutocomplete(sheetView, idxRow - 1, idxCol - 1, dictionary);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
