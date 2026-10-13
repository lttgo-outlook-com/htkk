using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class AcBcSl26Nckt201
{
	private Dictionary<string, string[]> dicDMThueBVMT = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicHTHoaDon = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMThueBVMTTmp = new Dictionary<string, string>();

	private bool isExistsDlKyTruoc = false;

	private string pathCellNguoiDaiDien = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/nguoiDaiDien";

	private string pathCellNgayBaoCao = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayBCao";

	private string pathCellNguoiLapBieu = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/nguoiLapBieu";

	private List<Dictionary<string, string>> rowValues;

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 4 && idxCol == 1)
			{
				ShowBangKe("Bao_Cao_BC26_AC.xls");
			}
			if (winFpSpread.Sheets[2].Visible)
			{
				if (idxCol == HTKKFile.ColumnLettersToIndex("B") - 1 && idxRow == 13 && winFpSpread.Sheets[0].Cells[13, HTKKFile.ColumnLettersToIndex("G") - 1].Value.ToBoolOrDefault())
				{
					winFpSpread.Sheets[0].Cells[13, HTKKFile.ColumnLettersToIndex("B") - 1].Value = false;
					winFpSpread.Sheets[0].Cells[13, HTKKFile.ColumnLettersToIndex("G") - 1].Value = true;
				}
				else if (idxCol == HTKKFile.ColumnLettersToIndex("B") - 1 && idxRow == 13)
				{
					winFpSpread.Sheets[0].Cells[13, HTKKFile.ColumnLettersToIndex("G") - 1].Value = false;
				}
				else if (idxCol == HTKKFile.ColumnLettersToIndex("G") - 1 && idxRow == 13)
				{
					if (!winFpSpread.Sheets[0].Cells[13, HTKKFile.ColumnLettersToIndex("G") - 1].Value.ToBoolOrDefault())
					{
						winFpSpread.Sheets[0].Cells[13, HTKKFile.ColumnLettersToIndex("G") - 1].Value = true;
					}
					winFpSpread.Sheets[0].Cells[13, HTKKFile.ColumnLettersToIndex("B") - 1].Value = false;
				}
			}
			else if (idxCol == HTKKFile.ColumnLettersToIndex("B") - 1 && idxRow == 13)
			{
				winFpSpread.Sheets[0].Cells[13, HTKKFile.ColumnLettersToIndex("B") - 1].Value = true;
				winFpSpread.Sheets[0].Cells[13, HTKKFile.ColumnLettersToIndex("G") - 1].Value = false;
			}
			else
			{
				winFpSpread.Sheets[0].Cells[13, HTKKFile.ColumnLettersToIndex("B") - 1].Value = false;
				winFpSpread.Sheets[0].Cells[13, HTKKFile.ColumnLettersToIndex("G") - 1].Value = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 1 && idxCol == 1)
			{
				ShowBangKe("Bangke_BC26_Mau_3_10.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 1 && idxCol == 1)
			{
				ShowBangKe("Bangke_BC26_Mau_3_12.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void ShowBangKe(string sFileName)
	{
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string text = Path.GetTempPath() + sFileName;
			File.Copy(HTKKConfig.INTERFACE_TEMPLATES + "\\excel\\" + sFileName, text, overwrite: true);
			Process.Start(text);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage = new frmMessage("0433", "3", "", "Thông báo");
			((Form)frmMessage).ShowDialog();
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			sView.Cells[13, HTKKFile.ColumnLettersToIndex("G") - 1].Locked = false;
			sView.Cells[13, HTKKFile.ColumnLettersToIndex("B") - 1].Locked = false;
			if (hTKKCMCSheetInfo != null && !ToKhaiInfo.SheetInfos[1].IsReadDataFile)
			{
				sView.Cells[13, HTKKFile.ColumnLettersToIndex("B") - 1].Value = true;
			}
			else if (hTKKCMCSheetInfo != null && !sView.Cells[13, HTKKFile.ColumnLettersToIndex("B") - 1].Value.ToBoolOrDefault() && !sView.Cells[13, HTKKFile.ColumnLettersToIndex("G") - 1].Value.ToBoolOrDefault())
			{
				sView.Cells[13, HTKKFile.ColumnLettersToIndex("B") - 1].Value = true;
				sView.Cells[13, HTKKFile.ColumnLettersToIndex("G") - 1].Value = false;
			}
			hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo != null && !ToKhaiInfo.SheetInfos[2].IsReadDataFile)
			{
				sView.Cells[13, HTKKFile.ColumnLettersToIndex("G") - 1].Value = true;
			}
			else if (hTKKCMCSheetInfo != null && !sView.Cells[13, HTKKFile.ColumnLettersToIndex("B") - 1].Value.ToBoolOrDefault() && !sView.Cells[13, HTKKFile.ColumnLettersToIndex("G") - 1].Value.ToBoolOrDefault())
			{
				sView.Cells[13, HTKKFile.ColumnLettersToIndex("B") - 1].Value = false;
				sView.Cells[13, HTKKFile.ColumnLettersToIndex("G") - 1].Value = true;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables[0];
			if (!hTKKCMCSheetInfo2.IsReadDataFile)
			{
				LayGiaTriTonKyTruoc(sView, PeriodInfo, HeaderInfo, ToKhaiInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
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

	public void AutoBindData_2(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
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

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("E");
			if (iCol == num)
			{
				string text = sheetView.Cells[(int)iRow - 1, num - 1].Text;
				string[] array = dicDMThueBVMT[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length >= 3)
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = array[1];
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[2];
				}
				else if (array.Length == 2)
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = array[1];
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = "";
				}
				string text2 = "";
				if (!string.IsNullOrEmpty(text))
				{
					text2 = dicDMThueBVMT[text][0];
				}
				if (text2.Equals("0") || text2.Equals("4"))
				{
					TextCellType textCellType = new TextCellType();
					textCellType.WordWrap = true;
					TextCellType textCellType2 = textCellType;
					textCellType2.MaxLength = 13;
					sheetView.Cells[(int)iRow - 1, num3 - 1].CellType = textCellType2;
				}
				else
				{
					TextCellType textCellType3 = new TextCellType();
					textCellType3.WordWrap = true;
					TextCellType textCellType2 = textCellType3;
					textCellType2.MaxLength = 20;
					sheetView.Cells[(int)iRow - 1, num3 - 1].CellType = textCellType2;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellEditchange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("O");
			if (iCol == num)
			{
				string text = sheetView.Cells[(int)iRow - 1, num - 1].Text;
				string[] array = dicDMThueBVMT[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length >= 3)
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = array[1];
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[2];
				}
				else if (array.Length == 2)
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = array[1];
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = "";
				}
				string text2 = "";
				if (!string.IsNullOrEmpty(text))
				{
					text2 = dicDMThueBVMT[text][0];
				}
				if (text2.Equals("0") || text2.Equals("4"))
				{
					TextCellType textCellType = new TextCellType();
					textCellType.WordWrap = true;
					TextCellType textCellType2 = textCellType;
					textCellType2.MaxLength = 13;
					sheetView.Cells[(int)iRow - 1, num3 - 1].CellType = textCellType2;
				}
				else
				{
					TextCellType textCellType3 = new TextCellType();
					textCellType3.WordWrap = true;
					TextCellType textCellType2 = textCellType3;
					textCellType2.MaxLength = 20;
					sheetView.Cells[(int)iRow - 1, num3 - 1].CellType = textCellType2;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellEditchange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("O");
			if (iCol == num)
			{
				string text = sheetView.Cells[(int)iRow - 1, num - 1].Text;
				string[] array = dicDMThueBVMT[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length >= 3)
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = array[1];
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[2];
				}
				else if (array.Length == 2)
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = array[1];
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = "";
				}
				string text2 = "";
				if (!string.IsNullOrEmpty(text))
				{
					text2 = dicDMThueBVMT[text][0];
				}
				if (text2.Equals("0") || text2.Equals("4"))
				{
					TextCellType textCellType = new TextCellType();
					textCellType.WordWrap = true;
					TextCellType textCellType2 = textCellType;
					textCellType2.MaxLength = 13;
					sheetView.Cells[(int)iRow - 1, num3 - 1].CellType = textCellType2;
				}
				else
				{
					TextCellType textCellType3 = new TextCellType();
					textCellType3.WordWrap = true;
					TextCellType textCellType2 = textCellType3;
					textCellType2.MaxLength = 20;
					sheetView.Cells[(int)iRow - 1, num3 - 1].CellType = textCellType2;
				}
			}
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
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			int num3 = HTKKFile.ColumnLettersToIndex("AA");
			string text = "";
			if (iCol == num)
			{
				string text2 = sheetView.Cells[(int)iRow - 1, num2 - 1].Text;
				if (!string.IsNullOrEmpty(text2))
				{
					text = dicDMThueBVMT[text2][0];
				}
				if (!text.Equals("3"))
				{
					string text3 = sheetView.Cells[(int)iRow - 1, num - 1].Text;
					sheetView.Cells[(int)iRow - 1, num - 1].Text = text3.ToUpper();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("O");
			int num3 = HTKKFile.ColumnLettersToIndex("S");
			int num4 = HTKKFile.ColumnLettersToIndex("V");
			int num5 = HTKKFile.ColumnLettersToIndex("Y");
			int num6 = HTKKFile.ColumnLettersToIndex("AA");
			string text = "";
			string text2 = "";
			if (iCol == num4 || iCol == num5)
			{
				text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text);
				string text3 = sheetView.Cells[(int)iRow - 1, num4 - 1].Text;
				string text4 = sheetView.Cells[(int)iRow - 1, num5 - 1].Text;
				if (!string.IsNullOrEmpty(text3) && !string.IsNullOrEmpty(text4))
				{
					int num7 = ConvertHelpers.ConvertToInt(text4) - ConvertHelpers.ConvertToInt(text3) + 1;
					sheetView.Cells[(int)iRow - 1, num6 - 1].Text = num7.ToString();
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num6 - 1].Text = "0";
				}
			}
			if (iCol == num3)
			{
				text = sheetView.Cells[(int)iRow - 1, num3 - 1].Text;
				sheetView.Cells[(int)iRow - 1, num3 - 1].Text = GeneralFunctions.FormatKyHieu(text);
			}
			if (iCol == num2)
			{
				string text5 = sheetView.Cells[(int)iRow - 1, num - 1].Text;
				if (!string.IsNullOrEmpty(text5))
				{
					text2 = dicDMThueBVMT[text5][0];
				}
				if (!text2.Equals("3"))
				{
					text = sheetView.Cells[(int)iRow - 1, num2 - 1].Text;
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = text.ToUpper();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("O");
			int num3 = HTKKFile.ColumnLettersToIndex("S");
			int num4 = HTKKFile.ColumnLettersToIndex("V");
			int num5 = HTKKFile.ColumnLettersToIndex("Y");
			int num6 = HTKKFile.ColumnLettersToIndex("AA");
			string text = "";
			string text2 = "";
			if (iCol == num4 || iCol == num5)
			{
				text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text);
				string text3 = sheetView.Cells[(int)iRow - 1, num4 - 1].Text;
				string text4 = sheetView.Cells[(int)iRow - 1, num5 - 1].Text;
				if (!string.IsNullOrEmpty(text3) && !string.IsNullOrEmpty(text4))
				{
					int num7 = ConvertHelpers.ConvertToInt(text4) - ConvertHelpers.ConvertToInt(text3) + 1;
					sheetView.Cells[(int)iRow - 1, num6 - 1].Text = num7.ToString();
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num6 - 1].Text = "0";
				}
			}
			if (iCol == num3)
			{
				text = sheetView.Cells[(int)iRow - 1, num3 - 1].Text;
				sheetView.Cells[(int)iRow - 1, num3 - 1].Text = GeneralFunctions.FormatKyHieu(text);
			}
			if (iCol == num2)
			{
				string text5 = sheetView.Cells[(int)iRow - 1, num - 1].Text;
				if (!string.IsNullOrEmpty(text5))
				{
					text2 = dicDMThueBVMT[text5][0];
				}
				if (!text2.Equals("3"))
				{
					text = sheetView.Cells[(int)iRow - 1, num2 - 1].Text;
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = text.ToUpper();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
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

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		int value = 0;
		levelWarning.TryGetValue(0, out value);
		string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[12, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[31, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[52, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[55, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			List<int> list = new List<int>();
			for (int i = 0; i < rowsValues.Count; i++)
			{
				int num = HTKKFile.ColumnLettersToIndex("C");
				int num2 = HTKKFile.ColumnLettersToIndex("D");
				int num3 = HTKKFile.ColumnLettersToIndex("E");
				int num4 = HTKKFile.ColumnLettersToIndex("F");
				int num5 = HTKKFile.ColumnLettersToIndex("G");
				int num6 = HTKKFile.ColumnLettersToIndex("H");
				int num7 = HTKKFile.ColumnLettersToIndex("I");
				int num8 = HTKKFile.ColumnLettersToIndex("J");
				int num9 = HTKKFile.ColumnLettersToIndex("K");
				int num10 = HTKKFile.ColumnLettersToIndex("L");
				int num11 = HTKKFile.ColumnLettersToIndex("M");
				int num12 = HTKKFile.ColumnLettersToIndex("N");
				int num13 = HTKKFile.ColumnLettersToIndex("P");
				string text8 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text;
				string text9 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text;
				string text10 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Text;
				string text11 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Text;
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, num5, "16");
				string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, num6, "16");
				string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, num7, "16");
				string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, num8, "16");
				string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, num9, "16");
				string valueCustom6 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, num10, "16");
				string valueCustom7 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, num11, "16");
				string valueCustom8 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, num12, "16");
				string text12 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num13 - 1].Text;
				string text13 = "";
				if (HTKKFormater.isNullString(text8) && HTKKFormater.isNullString(text9) && HTKKFormater.isNullString(text10) && HTKKFormater.isNullString(text11) && ConvertHelpers.ConvertToDouble(valueCustom2) == 0.0 && ConvertHelpers.ConvertToDouble(valueCustom3) == 0.0 && ConvertHelpers.ConvertToDouble(valueCustom4) == 0.0 && ConvertHelpers.ConvertToDouble(valueCustom5) == 0.0 && ConvertHelpers.ConvertToDouble(valueCustom6) == 0.0)
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num5 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num10 - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart + i - 1, num11 - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart + i - 1, num12 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num13 - 1);
					continue;
				}
				GeneralFunctions.strMauSoHD_01GTKT = text8;
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num5 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num10 - 1);
				HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart + i - 1, num11 - 1);
				HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart + i - 1, num12 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num13 - 1);
				if (string.IsNullOrEmpty(text8))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, text);
					list.Add(1);
				}
				if (string.IsNullOrEmpty(text9))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1, text);
					list.Add(1);
				}
				string strLoai = "";
				string[] value2 = null;
				dicDMThueBVMT.TryGetValue(text8, out value2);
				if (value2 != null && value2.Length > 0)
				{
					strLoai = value2[0];
				}
				if (string.IsNullOrEmpty(text10))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text);
					list.Add(1);
				}
				else
				{
					text13 = GeneralFunctions.CheckMauSoHD(text10, strLoai, text8);
					if (text13.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text2);
						list.Add(1);
					}
					else if (text13.Equals("2"))
					{
						if (text11.Trim().Equals("E"))
						{
							if (GeneralFunctions.CheckSoLienHDDT(text10, strLoai, text8, text11).Equals("3"))
							{
								HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text6);
								list.Add(0);
							}
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text3);
							list.Add(1);
						}
					}
					else if (text11.Length > 0 && text11.Trim().Equals("E") && GeneralFunctions.CheckSoLienHDDT(text10, strLoai, text8, text11).Equals("3"))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text6);
						list.Add(0);
					}
				}
				if (string.IsNullOrEmpty(text11))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1, text);
					list.Add(1);
				}
				if (ConvertHelpers.ConvertToInt(valueCustom8) < 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num5 - 1, text7);
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1, text7);
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1, text7);
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, text7);
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1, text7);
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num10 - 1, text7);
					list.Add(1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellNgayBaoCao));
			string text14 = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
			if (string.IsNullOrEmpty(text14))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, text);
				list.Add(1);
			}
			else
			{
				HTKKFormater.IsDateTime(text14, "dd/MM/yyyy", out var dtDate);
				if (DateTime.Compare(dtDate, DateTime.Now) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					list.Add(1);
				}
				else if (DateTime.Compare(dtDate, DateTime.ParseExact(PeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture)) < 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_52")));
					list.Add(1);
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellNguoiDaiDien));
			if (hTKKCMCCellInfo2 != null)
			{
				string str = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
				if (HTKKFormater.isNullString(str))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, text);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellNguoiLapBieu));
			if (hTKKCMCCellInfo3 != null)
			{
				string str2 = sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
				if (HTKKFormater.isNullString(str2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, text);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
				}
			}
			for (int i = 0; i < list.Count; i++)
			{
				if (list[i] == 1)
				{
					value = 2;
					break;
				}
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		int value = 0;
		levelWarning.TryGetValue(1, out value);
		string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[31, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[33, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[35, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[36, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[53, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[54, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[55, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			List<int> list = new List<int>();
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int startRow = location.DataLocationTemplateRowStart - 1;
			string[] array = GeneralFunctions.checkTrung_01(winFpSpread, startRow, "V", "Y", "B", "D", "O", "S", "", isMST: false, 2, "bb");
			string strLoai = "";
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location2 = hTKKCMCTableInfo2.Location;
				int dataLocationTemplateRowStart = location2.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo2.RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					int num = HTKKFile.ColumnLettersToIndex("C");
					int num2 = HTKKFile.ColumnLettersToIndex("D");
					int num3 = HTKKFile.ColumnLettersToIndex("O");
					int num4 = HTKKFile.ColumnLettersToIndex("S");
					int num5 = HTKKFile.ColumnLettersToIndex("V");
					int num6 = HTKKFile.ColumnLettersToIndex("Y");
					string text13 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text;
					string text14 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text;
					string text15 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Text;
					string text16 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text;
					string text17 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Text;
					string text18 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num6 - 1].Text;
					string text19 = "";
					if (HTKKFormater.isNullString(text13) && HTKKFormater.isNullString(text14) && HTKKFormater.isNullString(text15) && HTKKFormater.isNullString(text16) && ConvertHelpers.ConvertToDouble(text17) == 0.0 && ConvertHelpers.ConvertToDouble(text18) == 0.0)
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1);
						continue;
					}
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1);
					GeneralFunctions.strMauSoHD_01GTKT = text13;
					if (string.IsNullOrEmpty(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num - 1, text8);
						list.Add(1);
					}
					if (string.IsNullOrEmpty(text14))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1, text8);
						list.Add(1);
					}
					if (!string.IsNullOrEmpty(text13))
					{
						strLoai = dicDMThueBVMT[text13][0];
					}
					if (string.IsNullOrEmpty(text15))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text8);
						list.Add(1);
					}
					else
					{
						text19 = GeneralFunctions.CheckMauSoHD(text15, strLoai, text13);
						if (text19.Equals("1"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text3);
							list.Add(1);
						}
						else if (text19.Equals("2"))
						{
							if (text16.Length > 0)
							{
								if (text16.Substring(text16.Length - 1).Equals("E"))
								{
									if (GeneralFunctions.CheckSoLienHDDT(text15, strLoai, text13, text16).Equals("3"))
									{
										HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text10);
										list.Add(0);
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text);
									list.Add(1);
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text);
								list.Add(1);
							}
						}
						else if (text16.Length > 0 && text16.Substring(text16.Length - 1).Equals("E") && GeneralFunctions.CheckSoLienHDDT(text15, strLoai, text13, text16).Equals("3"))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text10);
							list.Add(0);
						}
					}
					if (string.IsNullOrEmpty(text16))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, text8);
						list.Add(1);
					}
					else
					{
						text19 = GeneralFunctions.CheckSoHD(text16, strLoai);
						if (text19.Equals("1"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, text4);
							list.Add(1);
						}
						else if (text19.Equals("2"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, text6);
							list.Add(1);
						}
					}
					if (string.IsNullOrEmpty(text17) && string.IsNullOrEmpty(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text8);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text8);
						list.Add(1);
					}
					else if (!string.IsNullOrEmpty(text17) && string.IsNullOrEmpty(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text8);
						list.Add(1);
					}
					else if (string.IsNullOrEmpty(text17) && !string.IsNullOrEmpty(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text8);
						list.Add(1);
					}
					else if ((ConvertHelpers.ConvertToInt(text18) < ConvertHelpers.ConvertToInt(text17) || ConvertHelpers.ConvertToInt(text18) <= 0) && (ConvertHelpers.ConvertToInt(text18) != 0 || ConvertHelpers.ConvertToInt(text17) != 0))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text7);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text7);
						list.Add(1);
					}
					if (HTKKFormater.isNullString(text17) || HTKKFormater.isNullString(text18) || array.Length <= 1)
					{
						continue;
					}
					for (int k = 0; k < array.Length; k++)
					{
						if (dataLocationTemplateRowStart + j - 1 == Convert.ToInt32(array[k]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text9);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text9);
							list.Add(1);
							break;
						}
					}
				}
			}
			for (int j = 0; j < list.Count; j++)
			{
				if (list[j] == 1)
				{
					value = 2;
					break;
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
			throw ex;
		}
	}

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		int value = 0;
		levelWarning.TryGetValue(2, out value);
		string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[31, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[33, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[35, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[36, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[53, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[54, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[55, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			List<int> list = new List<int>();
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				string[] array = GeneralFunctions.checkTrung(winFpSpread, dataLocationTemplateRowStart - 1, "V", "Y", "B", "D", "O", "S", "", isMST: false, 3);
				for (int j = 0; j < rowsValues.Count; j++)
				{
					int num = HTKKFile.ColumnLettersToIndex("C");
					int num2 = HTKKFile.ColumnLettersToIndex("D");
					int num3 = HTKKFile.ColumnLettersToIndex("O");
					int num4 = HTKKFile.ColumnLettersToIndex("S");
					int num5 = HTKKFile.ColumnLettersToIndex("V");
					int num6 = HTKKFile.ColumnLettersToIndex("Y");
					string text13 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text;
					string text14 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text;
					string text15 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Text;
					string text16 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text;
					string text17 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Text;
					string text18 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num6 - 1].Text;
					string text19 = "";
					string strLoai = "";
					if (HTKKFormater.isNullString(text13) && HTKKFormater.isNullString(text14) && HTKKFormater.isNullString(text15) && HTKKFormater.isNullString(text16) && ConvertHelpers.ConvertToDouble(text17) == 0.0 && ConvertHelpers.ConvertToDouble(text18) == 0.0)
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1);
						continue;
					}
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1);
					GeneralFunctions.strMauSoHD_01GTKT = text13;
					if (string.IsNullOrEmpty(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num - 1, text8);
						list.Add(1);
					}
					if (string.IsNullOrEmpty(text14))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1, text8);
						list.Add(1);
					}
					if (string.IsNullOrEmpty(text15))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text8);
						list.Add(1);
					}
					else
					{
						strLoai = dicDMThueBVMT[text13][0];
						text19 = GeneralFunctions.CheckMauSoHD(text15, strLoai, text13);
						if (text19.Equals("1"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text3);
							list.Add(1);
						}
						else if (text19.Equals("2"))
						{
							if (text16.Length > 0)
							{
								if (text16.Substring(text16.Length - 1).Equals("E"))
								{
									if (GeneralFunctions.CheckSoLienHDDT(text15, strLoai, text13, text16).Equals("3"))
									{
										HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text11);
										list.Add(0);
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text);
									list.Add(1);
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text);
								list.Add(1);
							}
						}
						else if (text16.Length > 0 && text16.Substring(text16.Length - 1).Equals("E") && GeneralFunctions.CheckSoLienHDDT(text15, strLoai, text13, text16).Equals("3"))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text11);
							list.Add(0);
						}
					}
					if (string.IsNullOrEmpty(text16))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, text8);
						list.Add(1);
					}
					else
					{
						text19 = GeneralFunctions.CheckSoHD(text16, strLoai);
						if (text19.Equals("1"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, text4);
							list.Add(1);
						}
						else if (text19.Equals("2"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, text6);
							list.Add(1);
						}
					}
					if (string.IsNullOrEmpty(text17) && string.IsNullOrEmpty(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text8);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text8);
						list.Add(1);
					}
					else if (!string.IsNullOrEmpty(text17) && string.IsNullOrEmpty(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text8);
						list.Add(1);
					}
					else if (string.IsNullOrEmpty(text17) && !string.IsNullOrEmpty(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text8);
						list.Add(1);
					}
					else if ((ConvertHelpers.ConvertToInt(text18) < ConvertHelpers.ConvertToInt(text17) || ConvertHelpers.ConvertToInt(text18) <= 0) && (ConvertHelpers.ConvertToInt(text18) != 0 || ConvertHelpers.ConvertToInt(text17) != 0))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text7);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text7);
						list.Add(1);
					}
					if (HTKKFormater.isNullString(text17) || HTKKFormater.isNullString(text18) || array.Length <= 1)
					{
						continue;
					}
					for (int k = 0; k < array.Length; k++)
					{
						if (dataLocationTemplateRowStart + j - 1 == Convert.ToInt32(array[k]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text9);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text9);
							list.Add(1);
							break;
						}
					}
				}
			}
			for (int j = 0; j < list.Count; j++)
			{
				if (list[j] == 1)
				{
					value = 2;
					break;
				}
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[2].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[2].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[2].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public Dictionary<string, string> GetDMHoaDon(string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		try
		{
			if (dicDMThueBVMTTmp.Count > 0)
			{
				return dicDMThueBVMTTmp;
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
				dictionary2.Add(key, value2);
				dictionary.Add(key, array);
			}
			dicDMThueBVMT = dictionary;
			dicDMThueBVMTTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetHTHoaDon()
	{
		try
		{
			if (dicHTHoaDon.Count > 0)
			{
				return dicHTHoaDon;
			}
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			dictionary.Add("1", "T");
			dictionary.Add("2", "P");
			dictionary.Add("3", "E");
			dicHTHoaDon = dictionary;
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void CellFormatImport(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int sheetIndex, long iRow, long iCol)
	{
		try
		{
			switch (sheetIndex)
			{
			case 1:
				CellChange_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, iRow, iCol);
				break;
			case 2:
				CellChange_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, iRow, iCol);
				break;
			}
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
			int num = HTKKFile.ColumnLettersToIndex("B");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = 14;
			switch (actionType)
			{
			case 2:
				winFpSpread.Sheets[0].Cells[num3 - 1, num - 1].Value = false;
				winFpSpread.Sheets[0].Cells[num3 - 1, num2 - 1].Value = false;
				winFpSpread.Sheets[0].Cells[num3 - 1, num - 1].Locked = true;
				winFpSpread.Sheets[0].Cells[num3 - 1, num2 - 1].Locked = true;
				break;
			case 1:
				winFpSpread.Sheets[0].Cells[num3 - 1, num2 - 1].Value = false;
				winFpSpread.Sheets[0].Cells[num3 - 1, num - 1].Value = true;
				winFpSpread.Sheets[0].Cells[num3 - 1, num - 1].Locked = false;
				winFpSpread.Sheets[0].Cells[num3 - 1, num2 - 1].Locked = false;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_2(FpSpread winFpSpread, int actionType)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("B");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = 14;
			switch (actionType)
			{
			case 2:
				winFpSpread.Sheets[0].Cells[num3 - 1, num2 - 1].Value = false;
				winFpSpread.Sheets[0].Cells[num3 - 1, num - 1].Value = false;
				break;
			case 1:
				winFpSpread.Sheets[0].Cells[num3 - 1, num2 - 1].Value = true;
				winFpSpread.Sheets[0].Cells[num3 - 1, num - 1].Value = false;
				break;
			}
			winFpSpread.Sheets[0].Cells[num3 - 1, num - 1].Locked = true;
			winFpSpread.Sheets[0].Cells[num3 - 1, num2 - 1].Locked = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void LayGiaTriTonKyTruoc(SheetView sViewTK, HTKKPeriodInfo periodInfo, HTKKHeaderInfo headerInfo, HTKKCMCToKhaiInfo tokhaiInfo)
	{
		try
		{
			string preperiodFileName_Quy = GeneralFunctions.GetPreperiodFileName_Quy(periodInfo, tokhaiInfo);
			if (string.IsNullOrEmpty(preperiodFileName_Quy))
			{
				return;
			}
			string text = HTKKConfig.DATAFILES + "\\" + headerInfo.mst;
			DirectoryInfo directoryInfo = new DirectoryInfo(text);
			bool flag = false;
			FileInfo[] files = directoryInfo.GetFiles();
			if (files.Length > 0)
			{
				FileInfo[] array = files;
				foreach (FileInfo fileInfo in array)
				{
					if (fileInfo.Name.Contains(preperiodFileName_Quy))
					{
						flag = true;
						break;
					}
				}
			}
			if (!flag)
			{
				return;
			}
			isExistsDlKyTruoc = true;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = tokhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			List<string> list = new List<string>();
			list.Add("maHoaDon");
			list.Add("kHieuMauHDon");
			list.Add("tenHDon");
			list.Add("hinhThucHDon");
			list.Add("tonCuoiKy");
			List<string> lstCellPath = list;
			rowValues = GetRowDataValueInXml($"{text}\\{preperiodFileName_Quy}.xml", hTKKCMCTableInfo, lstCellPath);
			int count = rowValues.Count;
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int j = 0; j < count; j++)
			{
				sViewTK.Rows.Add(dataLocationTemplateRowStart, 1);
				sViewTK.CopyRange(dataLocationTemplateRowStart - 1, -1, dataLocationTemplateRowStart, -1, 1, -1, dataOnly: false);
				hTKKCMCTableInfo.addNewDefaultRow(j + 1);
			}
			hTKKCMCSheetInfo.IncrementStaticCell(dataLocationTemplateRowStart, count);
			int num = dataLocationTemplateRowStart - 1;
			for (int j = 0; j < rowValues.Count; j++)
			{
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("B") - 1].Text = (j + 1).ToString();
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("C") - 1].Text = rowValues[j]["maHoaDon"];
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("D") - 1].Text = rowValues[j]["tenHDon"];
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("E") - 1].Text = rowValues[j]["kHieuMauHDon"];
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("F") - 1].Text = rowValues[j]["hinhThucHDon"];
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("G") - 1].Value = ConvertHelpers.ConvertToDouble(rowValues[j]["tonCuoiKy"]);
				Row row = sViewTK.Rows[num + j];
				float preferredHeight = row.GetPreferredHeight();
				if (row.Height < preferredHeight - 5f || row.Height > preferredHeight + 5f)
				{
					row.Height = preferredHeight;
				}
			}
			sViewTK.Cells[num + rowValues.Count, HTKKFile.ColumnLettersToIndex("B") - 1].Text = (rowValues.Count + 1).ToString();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private List<Dictionary<string, string>> GetRowDataValueInXml(string filename, HTKKCMCTableInfo tableInfo, List<string> lstCellPath)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		List<Dictionary<string, string>> list = new List<Dictionary<string, string>>();
		try
		{
			XmlDocument val = new XmlDocument();
			val.Load(filename);
			XmlNode val2 = ((XmlNode)val).SelectSingleNode("Sections/Section[@TableName='" + tableInfo.TableName + "']");
			XmlNodeList val3 = val2.SelectNodes("RowInfo/Cell");
			Dictionary<int, string> dictionary = new Dictionary<int, string>();
			for (int i = 0; i < val3.Count; i++)
			{
				string value = ((XmlNode)val3[i].Attributes["Path"]).Value;
				if (lstCellPath.Contains(value))
				{
					dictionary.Add(i, value);
				}
				lstCellPath.Remove(value);
				if (lstCellPath.Count == 0)
				{
					break;
				}
			}
			if (dictionary.Count == 0)
			{
				return list;
			}
			XmlNodeList val4 = val2.SelectNodes("RowData");
			for (int j = 0; j < val4.Count; j++)
			{
				Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
				string text = "###";
				string[] array = ConvertHelpers.ConvertToString(val4[j].InnerXml).Split(new string[1] { text }, StringSplitOptions.None);
				for (int k = 0; k < array.Length; k++)
				{
					if (dictionary.ContainsKey(k))
					{
						dictionary2.Add(dictionary[k], array[k]);
					}
				}
				if (!dictionary2["tonCuoiKy"].Equals("0"))
				{
					list.Add(dictionary2);
				}
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return list;
		}
	}
}
