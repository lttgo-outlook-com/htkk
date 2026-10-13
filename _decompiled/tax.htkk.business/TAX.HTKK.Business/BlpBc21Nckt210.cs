using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BlpBc21Nckt210
{
	private Dictionary<string, string> dicBLPTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicBLP = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicGhiChu = new Dictionary<string, string>();

	private ComboBoxCellType cbxGhiChu;

	private string pathLyDoMat = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/lyDoMatChayBLai";

	private string pathDonViNhanBC = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/donViNhanBCao";

	private string pathCellNguoiDaiDien = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/nguoiDaiDien";

	private string pathCellNgayBaoCao = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayBCao";

	private int[] CellId_NgayPhatHien => GeneralFunctions.ArrayCellId("O_8");

	private int[] CellId_GioPhatHien => GeneralFunctions.ArrayCellId("K_8");

	private int[] CellId_PhutPhatHien => GeneralFunctions.ArrayCellId("M_8");

	private int idxCol_MaBienLai => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_MaBienLai_hide => HTKKFile.ColumnLettersToIndex("AI") - 1;

	private int idxCol_TenBienLai => HTKKFile.ColumnLettersToIndex("E") - 1;

	private int idxCol_KyHieuMau => HTKKFile.ColumnLettersToIndex("O") - 1;

	private int idxCol_KyHieuBL => HTKKFile.ColumnLettersToIndex("S") - 1;

	private int idxCol_TuSo => HTKKFile.ColumnLettersToIndex("V") - 1;

	private int idxCol_DenSo => HTKKFile.ColumnLettersToIndex("Y") - 1;

	private int idxCol_SoLuong => HTKKFile.ColumnLettersToIndex("AA") - 1;

	private int idxCol_LienBienLai => HTKKFile.ColumnLettersToIndex("AB") - 1;

	private int idxCol_GhiChu => HTKKFile.ColumnLettersToIndex("AC") - 1;

	private int idxCol_MaGhiChu => HTKKFile.ColumnLettersToIndex("AL") - 1;

	public BlpBc21Nckt210()
	{
		if (dicGhiChu == null || dicGhiChu.Count == 0)
		{
			LoadGhiChuToDic();
		}
		if (cbxGhiChu == null)
		{
			BindCbxGhiChu();
		}
	}

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("K");
			int num2 = HTKKFile.ColumnLettersToIndex("M");
			int num3 = HTKKFile.ColumnLettersToIndex("O");
			int num4 = 8;
			if (HTKKFormater.isNullString(sView.Cells[num4 - 1, num - 1].Text) && HTKKFormater.isNullString(sView.Cells[num4 - 1, num2 - 1].Text) && HTKKFormater.isNullString(sView.Cells[num4 - 1, num3 - 1].Text))
			{
				sView.Cells[9, HTKKFile.ColumnLettersToIndex("S") - 1].Text = DateTime.Now.Second.ToString();
				sView.Cells[num4 - 1, num - 1].Text = DateTime.Now.Hour.ToString();
				sView.Cells[num4 - 1, num2 - 1].Text = DateTime.Now.Minute.ToString();
				sView.Cells[num4 - 1, num3 - 1].Text = DateTime.Now.Date.ToString("dd/MM/yyyy");
				int num5 = HTKKFile.ColumnLettersToIndex("K");
				int num6 = HTKKFile.ColumnLettersToIndex("M");
				int num7 = HTKKFile.ColumnLettersToIndex("O");
				int num8 = HTKKFile.ColumnLettersToIndex("S");
				string text = sView.Cells[num4 - 1, num7 - 1].Text;
				string text2 = sView.Cells[num4 - 1, num5 - 1].Text;
				string text3 = sView.Cells[num4 - 1, num6 - 1].Text;
				string text4 = sView.Cells[9, num8 - 1].Text;
				if (text.IndexOf("/") > 0 && !HTKKFormater.isNullString(text2) && !HTKKFormater.isNullString(text3))
				{
					string[] array = text.Split(new char[1] { '/' });
					sView.Cells[9, num5 - 1].Text = array[2] + "-" + array[1] + "-" + array[0] + "T" + GeneralFunctions.GetStrByLengh(text2, 2) + ":" + GeneralFunctions.GetStrByLengh(text3, 2) + ":" + GeneralFunctions.GetStrByLengh(text4, 2) + "Z";
				}
				else
				{
					sView.Cells[9, num5 - 1].Text = "";
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num9 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num10 = num9 + hTKKCMCTableInfo.RowsValues.Count - 1;
			for (int i = num9; i <= num10; i++)
			{
				sView.Cells[i, idxCol_GhiChu].CellType = cbxGhiChu;
				if (dicGhiChu.ContainsKey(sView.Cells[i, idxCol_MaGhiChu].Text))
				{
					sView.Cells[i, idxCol_GhiChu].Text = dicGhiChu[sView.Cells[i, idxCol_MaGhiChu].Text];
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void KeyPress(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e)
	{
		try
		{
			if (idxSheet != 0)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (idxRow >= num && idxRow <= num2)
			{
				if (idxCol == idxCol_TuSo || idxCol == idxCol_DenSo)
				{
					GeneralFunctions.CompareKeycode(e, 1);
				}
				if (idxCol == idxCol_LienBienLai)
				{
					GeneralFunctions.CompareKeycode(e, 2);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (iRow < dataLocationTemplateRowStart || iRow > num)
			{
				return;
			}
			if (iCol - 1 == idxCol_MaBienLai)
			{
				string text = sheetView.Cells[(int)iRow - 1, idxCol_MaBienLai].Text;
				if (HTKKFormater.isNullString(text))
				{
					sheetView.Cells[(int)iRow - 1, idxCol_MaBienLai_hide].ResetValue();
					sheetView.Cells[(int)iRow - 1, idxCol_TenBienLai].ResetValue();
					sheetView.Cells[(int)iRow - 1, idxCol_KyHieuMau].ResetValue();
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, idxCol_MaBienLai_hide].Text = text;
					sheetView.Cells[(int)iRow - 1, idxCol_TenBienLai].Text = dicBLP[text][2];
					sheetView.Cells[(int)iRow - 1, idxCol_KyHieuMau].Text = text;
				}
			}
			else
			{
				if (iCol - 1 != idxCol_GhiChu)
				{
					return;
				}
				string text2 = sheetView.Cells[(int)iRow - 1, idxCol_GhiChu].Text;
				if (HTKKFormater.isNullString(text2))
				{
					sheetView.Cells[(int)iRow - 1, idxCol_MaGhiChu].ResetValue();
					return;
				}
				string text3 = "";
				foreach (KeyValuePair<string, string> item in dicGhiChu)
				{
					if (item.Value.Equals(text2))
					{
						text3 = item.Key;
						break;
					}
				}
				sheetView.Cells[(int)iRow - 1, idxCol_MaGhiChu].Text = text3;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void LeaveCell_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, long iRow, long iCol, long iNewRow, long iNewCol)
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

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("K");
			int num4 = HTKKFile.ColumnLettersToIndex("M");
			int num5 = HTKKFile.ColumnLettersToIndex("O");
			int num6 = HTKKFile.ColumnLettersToIndex("S");
			int num7 = 8;
			string text = sheetView.Cells[num7 - 1, num5 - 1].Text;
			string text2 = sheetView.Cells[num7 - 1, num3 - 1].Text;
			string text3 = sheetView.Cells[num7 - 1, num4 - 1].Text;
			string text4 = sheetView.Cells[9, num6 - 1].Text;
			if (text.IndexOf("/") > 0 && !HTKKFormater.isNullString(text2) && !HTKKFormater.isNullString(text3))
			{
				string[] array = text.Split(new char[1] { '/' });
				sheetView.Cells[9, num3 - 1].Text = array[2] + "-" + array[1] + "-" + array[0] + "T" + GeneralFunctions.GetStrByLengh(text2, 2) + ":" + GeneralFunctions.GetStrByLengh(text3, 2) + ":" + GeneralFunctions.GetStrByLengh(text4, 2) + "Z";
			}
			else
			{
				sheetView.Cells[9, num3 - 1].Text = "";
			}
			if (iRow - 1 < num || iRow - 1 > num2)
			{
				return;
			}
			if (iCol - 1 == idxCol_KyHieuBL)
			{
				string text5 = sheetView.Cells[(int)iRow - 1, idxCol_KyHieuBL].Text;
				sheetView.Cells[(int)iRow - 1, idxCol_KyHieuBL].Text = GeneralFunctions.FormatKyHieuBLP(text5);
			}
			else if (iCol - 1 == idxCol_KyHieuMau)
			{
				string text6 = sheetView.Cells[(int)iRow - 1, idxCol_KyHieuMau].Text;
				sheetView.Cells[(int)iRow - 1, idxCol_KyHieuMau].Text = text6.ToUpper();
			}
			else
			{
				if (iCol - 1 != idxCol_TuSo && iCol - 1 != idxCol_DenSo)
				{
					return;
				}
				string text7 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				if (!HTKKFormater.isNullString(text7))
				{
					sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text7);
				}
				string text8 = sheetView.Cells[(int)iRow - 1, idxCol_TuSo].Text;
				string text9 = sheetView.Cells[(int)iRow - 1, idxCol_DenSo].Text;
				if (!HTKKFormater.isNullString(text8) && !HTKKFormater.isNullString(text9))
				{
					int num8 = ConvertHelpers.ConvertToInt(text8);
					int num9 = ConvertHelpers.ConvertToInt(text9);
					sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].ResetFormula();
					if (num8 == 0 && num9 == 0)
					{
						sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = 0;
					}
					else if (num8 == 0 && num9 != 0)
					{
						sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = num9;
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = num9 - num8 + 1;
					}
				}
				else if (HTKKFormater.isNullString(text8) && HTKKFormater.isNullString(text9))
				{
					sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].ResetValue();
				}
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			List<HTKKCMCCellInfo> staticCells = hTKKCMCSheetInfo.StaticCells;
			string text = string.Empty;
			string value = string.Empty;
			string value2 = string.Empty;
			string value3 = string.Empty;
			foreach (HTKKCMCCellInfo item in staticCells)
			{
				if (item.Path.Contains("ngayMatChayHongHD"))
				{
					text = item.Value.Substring(0, 10);
					value = item.Value.Substring(11, 2);
					value2 = item.Value.Substring(14, 2);
					value3 = item.Value;
					break;
				}
			}
			if (!text.IsNulOrEmpty())
			{
				DateTime dateTime = DateTime.ParseExact(text, "yyyy-MM-dd", CultureInfo.InvariantCulture);
				winFpSpread.Sheets[0].Cells[CellId_NgayPhatHien[1], CellId_NgayPhatHien[0]].Value = dateTime.ToString("dd/MM/yyyy");
				winFpSpread.Sheets[0].Cells[CellId_GioPhatHien[1], CellId_GioPhatHien[0]].Value = value;
				winFpSpread.Sheets[0].Cells[CellId_GioPhatHien[1] + 2, CellId_GioPhatHien[0]].Value = value3;
				winFpSpread.Sheets[0].Cells[CellId_PhutPhatHien[1], CellId_PhutPhatHien[0]].Value = value2;
			}
			string text2 = "";
			string empty = string.Empty;
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			int num3 = HTKKFile.ColumnLettersToIndex("O");
			int num4 = HTKKFile.ColumnLettersToIndex("AC");
			int num5 = HTKKFile.ColumnLettersToIndex("AL");
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text2 = sheetView.Cells[num + i, num3 - 1].Text;
				if (text2.Length > 5)
				{
					sheetView.Cells[num + i, num2 - 1].Text = text2.Substring(0, 5);
				}
				empty = winFpSpread.Sheets[0].Cells[num + i, num5 - 1].Text;
				if (!string.IsNullOrEmpty(empty) && dicGhiChu.ContainsKey(empty))
				{
					winFpSpread.Sheets[0].Cells[num + i, num4 - 1].Text = dicGhiChu[empty];
				}
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
			bool result = true;
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
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
			}
			return result;
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
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			string message9 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23"));
			string message10 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
			string message11 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25"));
			string message12 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			string message13 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27"));
			string message14 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28"));
			string message15 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29"));
			string message16 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30"));
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string[] array = GeneralFunctions.checkTrung(winFpSpread, num2, "V", "Y", "B", "C", "O", "S", "", isMST: false, 1);
			string[] array2 = new string[rowsInfo.Cells.Count];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array2[0] = sheetView.Cells[num2 + i, idxCol_MaBienLai].Text.Trim();
				array2[1] = sheetView.Cells[num2 + i, idxCol_TenBienLai].Text.Trim();
				array2[2] = sheetView.Cells[num2 + i, idxCol_KyHieuMau].Text.Trim();
				array2[3] = sheetView.Cells[num2 + i, idxCol_KyHieuBL].Text.Trim();
				array2[4] = sheetView.Cells[num2 + i, idxCol_TuSo].Text.Trim();
				array2[5] = sheetView.Cells[num2 + i, idxCol_DenSo].Text.Trim();
				array2[7] = sheetView.Cells[num2 + i, idxCol_LienBienLai].Text.Trim();
				array2[9] = sheetView.Cells[num2 + i, idxCol_GhiChu].Text.Trim();
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array2))
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_MaBienLai);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_KyHieuMau);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_GhiChu);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TuSo);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_DenSo);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TenBienLai);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_KyHieuBL);
					continue;
				}
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_MaBienLai);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_KyHieuMau);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_GhiChu);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TuSo);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_DenSo);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TenBienLai);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_KyHieuBL);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_LienBienLai);
				int num3 = 0;
				if (array2[2].Length > array2[0].Length + 1)
				{
					num3 = ConvertHelpers.ConvertToInt(array2[2].Substring(array2[0].Length, 1));
				}
				if (!string.IsNullOrEmpty(array2[7]))
				{
					string[] array3 = array2[7].Split(new char[1] { ';' });
					if (array3 != null)
					{
						for (int j = 0; j < array3.Length; j++)
						{
							if (array3[j].Equals("") || ConvertHelpers.ConvertToInt(array3[j]) > 9 || ConvertHelpers.ConvertToInt(array3[j]) < 0)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_LienBienLai, HTKKFile.GetMessage("0481"));
								num = 2;
							}
							else if (ConvertHelpers.ConvertToInt(array3[j]) > num3)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_LienBienLai, message16);
								num = 2;
							}
						}
					}
				}
				if (HTKKFormater.isNullString(array2[9]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_GhiChu, message10);
					num = 2;
				}
				if (array2[4].ToIntOrDefault() == 0 && array2[5].ToIntOrDefault() == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, message7);
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, message7);
					num = 2;
				}
				else if (array2[4].ToIntOrDefault() == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, message7);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_DenSo);
					num = 2;
				}
				else if (array2[5].ToIntOrDefault() == 0)
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TuSo);
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, message7);
					num = 2;
				}
				else
				{
					double num4 = ConvertHelpers.ConvertToDouble(array2[4]);
					double num5 = ConvertHelpers.ConvertToDouble(array2[5]);
					if (num4 > num5)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, message6);
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, message6);
						num = 2;
					}
				}
				if (array2[4].ToIntOrDefault(-1) == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32")));
					num = 2;
				}
				if (array2[5].ToIntOrDefault(-1) == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32")));
					num = 2;
				}
				if (HTKKFormater.isNullString(array2[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_MaBienLai, message7);
					num = 2;
				}
				if (HTKKFormater.isNullString(array2[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TenBienLai, message7);
					num = 2;
				}
				if (HTKKFormater.isNullString(array2[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, message7);
					num = 2;
				}
				else
				{
					string strLoai = dicBLP[array2[0]][0];
					string text = GeneralFunctions.CheckMauSoBLP(array2[2], strLoai, array2[0]);
					if (text.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, message);
						num = 2;
					}
					else if (text.Equals("2"))
					{
						if (array2[3].Length > 0)
						{
							if (array2[3].Substring(array2[3].Length - 1).Equals("E"))
							{
								if (GeneralFunctions.CheckSoLienBLPDT(array2[2], array2[0], array2[3]).Equals("3"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, message15);
									num = 2;
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, message3);
								num = 2;
							}
						}
					}
					else if (array2[3].Length > 0 && array2[3].Substring(array2[3].Length - 1).Equals("E") && GeneralFunctions.CheckSoLienBLPDT(array2[2], array2[0], array2[3]).Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, message15);
						num = 2;
					}
				}
				if (HTKKFormater.isNullString(array2[3]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuBL, message7);
					num = 2;
				}
				else
				{
					string text2 = GeneralFunctions.CheckSoBLP(array2[3], array2[0], DateTime.ParseExact(PeriodInfo.kyKKhai, "dd/MM/yyyy", CultureInfo.InvariantCulture) >= new DateTime(2017, 1, 1));
					if (HTKKFormater.isNullString(text2) || text2.Equals("1") || text2.Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuBL, message2);
						num = 2;
					}
				}
				if (array == null || array.Length <= 1)
				{
					continue;
				}
				for (int k = 0; k < array.Length; k++)
				{
					if (num2 + i == ConvertHelpers.ConvertToInt(array[k]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, message4);
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, message4);
						num = 2;
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, message4);
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, message4);
					}
				}
			}
			int num6 = sheetView.Cells[GeneralFunctions.ArrayCellId("K_8")[1], GeneralFunctions.ArrayCellId("K_8")[0]].Text.ToIntOrDefault(-1);
			int num7 = sheetView.Cells[GeneralFunctions.ArrayCellId("M_8")[1], GeneralFunctions.ArrayCellId("M_8")[0]].Value.ToIntOrDefault(-1);
			if (num6 < 0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, GeneralFunctions.ArrayCellId("K_8")[1], GeneralFunctions.ArrayCellId("K_8")[0], message7);
				num = 2;
			}
			else if (num6 > 23)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, GeneralFunctions.ArrayCellId("K_8")[1], GeneralFunctions.ArrayCellId("K_8")[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_31")));
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, GeneralFunctions.ArrayCellId("K_8")[1], GeneralFunctions.ArrayCellId("K_8")[0]);
			}
			if (num7 < 0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, GeneralFunctions.ArrayCellId("M_8")[1], GeneralFunctions.ArrayCellId("M_8")[0], message7);
				num = 2;
			}
			else if (num7 > 59)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, GeneralFunctions.ArrayCellId("M_8")[1], GeneralFunctions.ArrayCellId("M_8")[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_31")));
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, GeneralFunctions.ArrayCellId("M_8")[1], GeneralFunctions.ArrayCellId("M_8")[0]);
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellNgayBaoCao));
			string text3 = sheetView.Cells[CellId_NgayPhatHien[1], CellId_NgayPhatHien[0]].Text.Trim();
			string text4 = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
			DateTime t = new DateTime(2011, 1, 1);
			bool flag = true;
			DateTime dtDate;
			if (HTKKFormater.isNullString(text3))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_NgayPhatHien[1], CellId_NgayPhatHien[0], message7);
				num = 2;
			}
			else if (!HTKKFormater.IsDateTime(text3, "dd/MM/yyyy", out dtDate))
			{
				XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_NgayPhatHien[1], CellId_NgayPhatHien[0], xMLMessage.GetMessage("0081"));
				num = 2;
			}
			else if (DateTime.Compare(dtDate, t) < 0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_NgayPhatHien[1], CellId_NgayPhatHien[0], message11);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_NgayPhatHien[1], CellId_NgayPhatHien[0]);
				flag = false;
			}
			DateTime dtDate2;
			if (HTKKFormater.isNullString(text4))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message7);
				num = 2;
			}
			else if (!HTKKFormater.IsDateTime(text4, "dd/MM/yyyy", out dtDate2))
			{
				XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, xMLMessage.GetMessage("0081"));
				num = 2;
			}
			else if (DateTime.Compare(dtDate2, t) < 0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message12);
				num = 2;
			}
			else if (DateTime.Compare(dtDate2, DateTime.Now) > 0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message8);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
				if (!flag)
				{
					HTKKFormater.IsDateTime(text3, "dd/MM/yyyy", out dtDate);
					if (DateTime.Compare(dtDate, dtDate2) > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, CellId_NgayPhatHien[1], CellId_NgayPhatHien[0], message9);
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, CellId_NgayPhatHien[1], CellId_NgayPhatHien[0]);
					}
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathLyDoMat));
			if (hTKKCMCCellInfo2 != null)
			{
				string str = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
				if (HTKKFormater.isNullString(str))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, message7);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellNguoiDaiDien));
			if (hTKKCMCCellInfo3 != null)
			{
				string str2 = sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
				if (HTKKFormater.isNullString(str2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, message7);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
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

	private void LoadGhiChuToDic()
	{
		try
		{
			string message = HTKKFile.GetMessage("0472");
			string message2 = HTKKFile.GetMessage("0473");
			string message3 = HTKKFile.GetMessage("0474");
			dicGhiChu.Add("02", message);
			dicGhiChu.Add("03", message2);
			dicGhiChu.Add("04", message3);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void BindCbxGhiChu()
	{
		try
		{
			cbxGhiChu = new ComboBoxCellType();
			string[] array = new string[dicGhiChu.Count];
			int num = 0;
			foreach (KeyValuePair<string, string> item in dicGhiChu)
			{
				array[num] = item.Value;
				num++;
			}
			cbxGhiChu.Items = array;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public Dictionary<string, string> GetDMBienLaiPhi(string pathFile, HTKKPeriodInfo PeriodInfo)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		try
		{
			if (dicBLPTmp.Count > 0)
			{
				return dicBLPTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[1];
				string value2 = array[1];
				if (!dictionary2.ContainsKey(key))
				{
					dictionary2.Add(key, value2);
				}
				if (!dictionary.ContainsKey(key))
				{
					dictionary.Add(key, array);
				}
			}
			dicBLP = dictionary;
			dicBLPTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}
}
