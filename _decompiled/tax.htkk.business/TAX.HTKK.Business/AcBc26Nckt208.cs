using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
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

public class AcBc26Nckt208
{
	private Dictionary<string, string[]> dicDMHoaDon = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMHoaDonTmp = new Dictionary<string, string>();

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
			else if (winFpSpread.Sheets[2].Visible)
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
				ShowBangKe("Bangke_BC26_Mau_3_12.xls");
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
				ShowBangKe("Bangke_BC26_Mau_3_10.xls");
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

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				CellChange_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i, HTKKFile.ColumnLettersToIndex("I") - 1);
			}
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
			list.Add("kHieuHDon");
			list.Add("tonCuoiKy_tuSo");
			list.Add("tonCuoiKy_denSo");
			list.Add("tonCuoiKy_soLuong");
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
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("F") - 1].Text = rowValues[j]["kHieuHDon"];
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("G") - 1].Value = ConvertHelpers.ConvertToDouble(rowValues[j]["tonCuoiKy_soLuong"]);
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("H") - 1].Text = rowValues[j]["tonCuoiKy_tuSo"];
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("I") - 1].Text = rowValues[j]["tonCuoiKy_denSo"];
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("Z") - 1].Value = ConvertHelpers.ConvertToDouble(rowValues[j]["tonCuoiKy_soLuong"]);
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("X") - 1].Text = rowValues[j]["tonCuoiKy_tuSo"];
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("Y") - 1].Text = rowValues[j]["tonCuoiKy_denSo"];
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("AA") - 1].Text = rowValues[j]["maHoaDon"];
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("AC") - 1].Text = rowValues[j]["tonCuoiKy_tuSo"];
				sViewTK.Cells[num + j, HTKKFile.ColumnLettersToIndex("AD") - 1].Text = rowValues[j]["tonCuoiKy_denSo"];
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
				if (!dictionary2["tonCuoiKy_tuSo"].Equals("") && !dictionary2["tonCuoiKy_denSo"].Equals(""))
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

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = ToKhaiInfo.InterfaceTemplateExcel.LastIndexOf('\\');
			string text = ToKhaiInfo.InterfaceTemplateExcel.Substring(0, num + 1);
			text += "Bangke_BC26_Mau_3_12.xls";
			GeneralFunctions.InsertHyperLink(sView, 2, HTKKFile.ColumnLettersToIndex("B"), "Mẫu tải bảng kê", text);
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
			int num = ToKhaiInfo.InterfaceTemplateExcel.LastIndexOf('\\');
			string text = ToKhaiInfo.InterfaceTemplateExcel.Substring(0, num + 1);
			text += "Bangke_BC26_Mau_3_10.xls";
			GeneralFunctions.InsertHyperLink(sView, 2, HTKKFile.ColumnLettersToIndex("B"), "Mẫu tải bảng kê", text);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			if (hTKKCMCTableInfo.RowsValues.Count() > 0)
			{
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count(); i++)
				{
					reloadCacheTuSoDenSo(sView, ToKhaiInfo, num2 + i);
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
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("E");
			if (iCol == num)
			{
				string text = sheetView.Cells[(int)iRow - 1, num - 1].Text;
				string[] array = dicDMHoaDon[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
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
					text2 = dicDMHoaDon[text][0];
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
				string[] array = dicDMHoaDon[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
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
					text2 = dicDMHoaDon[text][0];
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
				string[] array = dicDMHoaDon[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
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
					text2 = dicDMHoaDon[text][0];
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			string text = "";
			int num3 = 0;
			string obj = "0";
			string obj2 = "0";
			string obj3 = "0";
			string obj4 = "0";
			string[] array = new string[1];
			int num4 = HTKKFile.ColumnLettersToIndex("C");
			int num5 = HTKKFile.ColumnLettersToIndex("E");
			int num6 = HTKKFile.ColumnLettersToIndex("F");
			int num7 = HTKKFile.ColumnLettersToIndex("G");
			int num8 = HTKKFile.ColumnLettersToIndex("H");
			int num9 = HTKKFile.ColumnLettersToIndex("I");
			int num10 = HTKKFile.ColumnLettersToIndex("J");
			int num11 = HTKKFile.ColumnLettersToIndex("K");
			int num12 = HTKKFile.ColumnLettersToIndex("L");
			int num13 = HTKKFile.ColumnLettersToIndex("M");
			int iCol2 = HTKKFile.ColumnLettersToIndex("N");
			int num14 = HTKKFile.ColumnLettersToIndex("P");
			int num15 = HTKKFile.ColumnLettersToIndex("Q");
			int num16 = HTKKFile.ColumnLettersToIndex("R");
			int num17 = HTKKFile.ColumnLettersToIndex("S");
			int num18 = HTKKFile.ColumnLettersToIndex("T");
			int num19 = HTKKFile.ColumnLettersToIndex("U");
			int num20 = HTKKFile.ColumnLettersToIndex("V");
			int num21 = HTKKFile.ColumnLettersToIndex("X");
			int num22 = HTKKFile.ColumnLettersToIndex("Y");
			int num23 = HTKKFile.ColumnLettersToIndex("Z");
			string text2 = sheetView.Cells[(int)iRow, num4 - 1].Text;
			string text3 = "";
			if (iRow - 1 < num || iRow - 1 > num2)
			{
				return;
			}
			if (iCol == num6)
			{
				text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatKyHieu(text);
			}
			if (iCol == num8 || iCol == num9 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num13 || iCol == num14 || iCol == num16 || iCol == num18 || iCol == num20)
			{
				if (iCol == num8 || iCol == num9 || iCol == num10 || iCol == num11)
				{
					text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
					sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text);
				}
				string text4 = sheetView.Cells[(int)iRow - 1, num16 - 1].Text;
				string soLuongHD = GeneralFunctions.GetSoLuongHD(text4);
				sheetView.Cells[(int)iRow - 1, num15 - 1].Value = ConvertHelpers.ConvertToDouble(soLuongHD);
				text4 = sheetView.Cells[(int)iRow - 1, num18 - 1].Text;
				soLuongHD = GeneralFunctions.GetSoLuongHD(text4);
				sheetView.Cells[(int)iRow - 1, num17 - 1].Value = ConvertHelpers.ConvertToDouble(soLuongHD);
				text4 = sheetView.Cells[(int)iRow - 1, num20 - 1].Text;
				soLuongHD = GeneralFunctions.GetSoLuongHD(text4);
				sheetView.Cells[(int)iRow - 1, num19 - 1].Value = ConvertHelpers.ConvertToDouble(soLuongHD);
				obj = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num14, "16");
				obj2 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num15, "16");
				obj3 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num17, "16");
				obj4 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num19, "16");
				string text5 = sheetView.Cells[(int)iRow - 1, num8 - 1].Text;
				string text6 = sheetView.Cells[(int)iRow - 1, num9 - 1].Text;
				string text7 = sheetView.Cells[(int)iRow - 1, num10 - 1].Text;
				string text8 = sheetView.Cells[(int)iRow - 1, num11 - 1].Text;
				array = GeneralFunctions.getTusoDenso(text5, text6, text7, text8);
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, iCol2, "16");
				string str;
				if (!array[0].Equals("-1"))
				{
					if (ConvertHelpers.ConvertToInt(obj4) + ConvertHelpers.ConvertToInt(obj2) + ConvertHelpers.ConvertToInt(obj3) + ConvertHelpers.ConvertToInt(obj) > 0)
					{
						sheetView.Cells[(int)iRow - 1, num12 - 1].Text = GeneralFunctions.FormatSoHD(array[0]);
						str = (ConvertHelpers.ConvertToInt(array[0]) + ConvertHelpers.ConvertToInt(valueCustom) - 1).ToString();
						sheetView.Cells[(int)iRow - 1, num13 - 1].Text = GeneralFunctions.FormatSoHD(str);
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, num12 - 1].Text = "";
					}
					sheetView.Cells[(int)iRow - 1, num7 - 1].Value = ConvertHelpers.ConvertToDouble(array[2]);
				}
				str = sheetView.Cells[(int)iRow - 1, num13 - 1].Text;
				string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num7, "16");
				if (ConvertHelpers.ConvertToInt(valueCustom2) == ConvertHelpers.ConvertToInt(valueCustom) && ConvertHelpers.ConvertToInt(obj4) + ConvertHelpers.ConvertToInt(obj2) + ConvertHelpers.ConvertToInt(obj3) + ConvertHelpers.ConvertToInt(obj) > 0)
				{
					num3 = ConvertHelpers.ConvertToInt(valueCustom2) - ConvertHelpers.ConvertToInt(valueCustom);
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num22 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num23 - 1].Value = ConvertHelpers.ConvertToDouble(num3.ToString());
				}
				else if (ConvertHelpers.ConvertToInt(valueCustom2) > ConvertHelpers.ConvertToInt(valueCustom) && ConvertHelpers.ConvertToInt(obj4) + ConvertHelpers.ConvertToInt(obj2) + ConvertHelpers.ConvertToInt(obj3) + ConvertHelpers.ConvertToInt(obj) > 0)
				{
					num3 = ConvertHelpers.ConvertToInt(valueCustom2) - ConvertHelpers.ConvertToInt(valueCustom);
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = GeneralFunctions.FormatSoHD((ConvertHelpers.ConvertToInt(str) + 1).ToString());
					sheetView.Cells[(int)iRow - 1, num22 - 1].Text = GeneralFunctions.FormatSoHD((ConvertHelpers.ConvertToInt(str) + num3).ToString());
					sheetView.Cells[(int)iRow - 1, num23 - 1].Value = ConvertHelpers.ConvertToDouble(num3.ToString());
				}
				else if (ConvertHelpers.ConvertToInt(valueCustom2) > ConvertHelpers.ConvertToInt(valueCustom) && ConvertHelpers.ConvertToInt(obj4) + ConvertHelpers.ConvertToInt(obj2) + ConvertHelpers.ConvertToInt(obj3) + ConvertHelpers.ConvertToInt(obj) == 0 && string.IsNullOrEmpty(text5) && string.IsNullOrEmpty(text6))
				{
					num3 = ConvertHelpers.ConvertToInt(valueCustom2);
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = GeneralFunctions.FormatSoHD(ConvertHelpers.ConvertToInt(text7).ToString());
					sheetView.Cells[(int)iRow - 1, num22 - 1].Text = GeneralFunctions.FormatSoHD(ConvertHelpers.ConvertToInt(text8).ToString());
					sheetView.Cells[(int)iRow - 1, num23 - 1].Value = ConvertHelpers.ConvertToDouble(num3.ToString());
				}
				else if (ConvertHelpers.ConvertToInt(valueCustom2) > ConvertHelpers.ConvertToInt(valueCustom) && ConvertHelpers.ConvertToInt(obj4) + ConvertHelpers.ConvertToInt(obj2) + ConvertHelpers.ConvertToInt(obj3) + ConvertHelpers.ConvertToInt(obj) == 0 && string.IsNullOrEmpty(text7) && string.IsNullOrEmpty(text8))
				{
					num3 = int.Parse(valueCustom2);
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = GeneralFunctions.FormatSoHD(ConvertHelpers.ConvertToInt(text5).ToString());
					sheetView.Cells[(int)iRow - 1, num22 - 1].Text = GeneralFunctions.FormatSoHD(ConvertHelpers.ConvertToInt(text6).ToString());
					sheetView.Cells[(int)iRow - 1, num23 - 1].Value = ConvertHelpers.ConvertToDouble(num3.ToString());
				}
				else
				{
					num3 = int.Parse(valueCustom2);
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = GeneralFunctions.FormatSoHD(ConvertHelpers.ConvertToInt(text5).ToString());
					sheetView.Cells[(int)iRow - 1, num22 - 1].Text = GeneralFunctions.FormatSoHD(ConvertHelpers.ConvertToInt(text8).ToString());
					sheetView.Cells[(int)iRow - 1, num23 - 1].Value = ConvertHelpers.ConvertToDouble(num3.ToString());
				}
			}
			if (iCol == num14 || iCol == num16 || iCol == num18 || iCol == num20)
			{
				text = sheetView.Cells[(int)iRow - 1, num16 - 1].Text;
				string soLuongHD2 = GeneralFunctions.GetSoLuongHD(text);
				sheetView.Cells[(int)iRow - 1, num15 - 1].Value = ConvertHelpers.ConvertToDouble(soLuongHD2);
				text = sheetView.Cells[(int)iRow - 1, num18 - 1].Text;
				soLuongHD2 = GeneralFunctions.GetSoLuongHD(text);
				sheetView.Cells[(int)iRow - 1, num17 - 1].Value = ConvertHelpers.ConvertToDouble(soLuongHD2);
				text = sheetView.Cells[(int)iRow - 1, num20 - 1].Text;
				soLuongHD2 = GeneralFunctions.GetSoLuongHD(text);
				sheetView.Cells[(int)iRow - 1, num19 - 1].Value = ConvertHelpers.ConvertToDouble(soLuongHD2);
				obj = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num14, "16");
				obj2 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num15, "16");
				obj3 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num17, "16");
				obj4 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num19, "16");
				string text9 = sheetView.Cells[(int)iRow - 1, num12 - 1].Text;
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, iCol2, "16");
				string str = (ConvertHelpers.ConvertToInt(text9) + ConvertHelpers.ConvertToInt(valueCustom) - 1).ToString();
				if (ConvertHelpers.ConvertToInt(obj4) + ConvertHelpers.ConvertToInt(obj3) + ConvertHelpers.ConvertToInt(obj2) + ConvertHelpers.ConvertToInt(obj) > 0)
				{
					sheetView.Cells[(int)iRow - 1, num13 - 1].Text = GeneralFunctions.FormatSoHD(str);
					if (!array[0].Equals("-1"))
					{
						sheetView.Cells[(int)iRow - 1, num12 - 1].Text = GeneralFunctions.FormatSoHD(array[0]);
					}
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num12 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num13 - 1].Text = "";
				}
				string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num7, "16");
				if (ConvertHelpers.ConvertToInt(valueCustom2) > ConvertHelpers.ConvertToInt(valueCustom) && ConvertHelpers.ConvertToInt(obj4) + ConvertHelpers.ConvertToInt(obj3) + ConvertHelpers.ConvertToInt(obj2) + ConvertHelpers.ConvertToInt(obj) > 0)
				{
					num3 = ConvertHelpers.ConvertToInt(valueCustom2) - ConvertHelpers.ConvertToInt(valueCustom);
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = GeneralFunctions.FormatSoHD((ConvertHelpers.ConvertToInt(str) + 1).ToString());
					sheetView.Cells[(int)iRow - 1, num22 - 1].Text = GeneralFunctions.FormatSoHD((ConvertHelpers.ConvertToInt(str) + num3).ToString());
					sheetView.Cells[(int)iRow - 1, num23 - 1].Value = ConvertHelpers.ConvertToDouble(num3.ToString());
				}
			}
			if (iCol == num14)
			{
				string text9 = sheetView.Cells[(int)iRow - 1, num12 - 1].Text;
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, iCol2, "16");
				string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num7, "16");
				string str = (ConvertHelpers.ConvertToInt(text9) + ConvertHelpers.ConvertToInt(valueCustom) - 1).ToString();
				if (ConvertHelpers.ConvertToInt(obj4) + ConvertHelpers.ConvertToInt(obj3) + ConvertHelpers.ConvertToInt(obj2) + ConvertHelpers.ConvertToInt(obj) > 0)
				{
					sheetView.Cells[(int)iRow - 1, num13 - 1].Text = GeneralFunctions.FormatSoHD(str);
					if (!array[0].Equals("-1"))
					{
						sheetView.Cells[(int)iRow - 1, num12 - 1].Text = GeneralFunctions.FormatSoHD(array[0]);
					}
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num12 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num13 - 1].Text = "";
				}
				if (ConvertHelpers.ConvertToInt(valueCustom2) > ConvertHelpers.ConvertToInt(valueCustom) && ConvertHelpers.ConvertToInt(obj4) + ConvertHelpers.ConvertToInt(obj3) + ConvertHelpers.ConvertToInt(obj2) + ConvertHelpers.ConvertToInt(obj) > 0)
				{
					num3 = ConvertHelpers.ConvertToInt(valueCustom2) - ConvertHelpers.ConvertToInt(valueCustom);
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = GeneralFunctions.FormatSoHD((ConvertHelpers.ConvertToInt(str) + 1).ToString());
					sheetView.Cells[(int)iRow - 1, num22 - 1].Text = GeneralFunctions.FormatSoHD((ConvertHelpers.ConvertToInt(str) + num3).ToString());
					sheetView.Cells[(int)iRow - 1, num23 - 1].Value = ConvertHelpers.ConvertToDouble(num3.ToString());
				}
			}
			if (iCol == num8 || iCol == num9 || iCol == num10 || iCol == num11 || iCol == num14 || iCol == num16 || iCol == num18 || iCol == num20)
			{
				string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num7, "16");
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, iCol2, "16");
				obj = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num14, "16");
				obj2 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num15, "16");
				obj3 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num17, "16");
				obj4 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num19, "16");
				if (ConvertHelpers.ConvertToInt(valueCustom2) + ConvertHelpers.ConvertToInt(valueCustom) + ConvertHelpers.ConvertToInt(obj) + ConvertHelpers.ConvertToInt(obj2) + ConvertHelpers.ConvertToInt(obj4) + ConvertHelpers.ConvertToInt(obj3) == 0)
				{
					num3 = 0;
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num22 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num23 - 1].Value = ConvertHelpers.ConvertToDouble(num3.ToString());
				}
				else if (ConvertHelpers.ConvertToInt(valueCustom2) == ConvertHelpers.ConvertToInt(valueCustom) && ConvertHelpers.ConvertToInt(obj4) + ConvertHelpers.ConvertToInt(obj3) + ConvertHelpers.ConvertToInt(obj2) + ConvertHelpers.ConvertToInt(obj) > 0)
				{
					num3 = 0;
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num22 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num23 - 1].Value = ConvertHelpers.ConvertToDouble(num3.ToString());
				}
				else if (ConvertHelpers.ConvertToInt(valueCustom2) < ConvertHelpers.ConvertToInt(valueCustom))
				{
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num22 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num23 - 1].Value = "0";
				}
			}
			if (iCol == num5)
			{
				if (!string.IsNullOrEmpty(text2))
				{
					text3 = dicDMHoaDon[text2][0];
				}
				if (!text3.Equals("3"))
				{
					text = sheetView.Cells[(int)iRow - 1, num5 - 1].Text;
					sheetView.Cells[(int)iRow - 1, num5 - 1].Text = text.ToUpper();
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
			iRow--;
			int num = HTKKFile.ColumnLettersToIndex("V");
			int num2 = HTKKFile.ColumnLettersToIndex("Y");
			int num3 = HTKKFile.ColumnLettersToIndex("AA");
			int num4 = HTKKFile.ColumnLettersToIndex("S");
			int num5 = HTKKFile.ColumnLettersToIndex("O");
			int num6 = HTKKFile.ColumnLettersToIndex("C");
			int num7 = 0;
			if (iCol == num || iCol == num2)
			{
				string text = sheetView.Cells[(int)iRow, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text);
				string text2 = sheetView.Cells[(int)iRow, num - 1].Text;
				string text3 = sheetView.Cells[(int)iRow, num2 - 1].Text;
				if (!string.IsNullOrEmpty(text2) || !string.IsNullOrEmpty(text3))
				{
					num7 = ConvertHelpers.ConvertToInt(text3) - ConvertHelpers.ConvertToInt(text2) + 1;
					if (num7 >= 1)
					{
						sheetView.Cells[(int)iRow, num3 - 1].Value = num7;
					}
					else
					{
						sheetView.Cells[(int)iRow, num3 - 1].Value = 0;
					}
				}
				else
				{
					sheetView.Cells[(int)iRow, num3 - 1].Value = 0;
				}
			}
			if (iCol == num4)
			{
				string text = sheetView.Cells[(int)iRow, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow, (int)iCol - 1].Text = GeneralFunctions.FormatKyHieu(text);
			}
			if (iCol == num5)
			{
				string text4 = sheetView.Cells[(int)iRow, num6 - 1].Text;
				string text5 = "";
				if (!string.IsNullOrEmpty(text4))
				{
					text5 = dicDMHoaDon[text4][0];
				}
				if (!text5.Equals("3"))
				{
					string text = sheetView.Cells[(int)iRow, num5 - 1].Text;
					sheetView.Cells[(int)iRow, num5 - 1].Text = text.ToUpper();
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
			iRow--;
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("V");
			int num3 = HTKKFile.ColumnLettersToIndex("Y");
			int num4 = HTKKFile.ColumnLettersToIndex("AA");
			int num5 = HTKKFile.ColumnLettersToIndex("S");
			int num6 = HTKKFile.ColumnLettersToIndex("O");
			int num7 = HTKKFile.ColumnLettersToIndex("C");
			int num8 = 0;
			if (iCol == num2 || iCol == num3)
			{
				string text = sheetView.Cells[(int)iRow, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text);
				string text2 = sheetView.Cells[(int)iRow, num2 - 1].Text;
				string text3 = sheetView.Cells[(int)iRow, num3 - 1].Text;
				if (!string.IsNullOrEmpty(text2) || !string.IsNullOrEmpty(text3))
				{
					num8 = ConvertHelpers.ConvertToInt(text3) - ConvertHelpers.ConvertToInt(text2) + 1;
					if (num8 >= 1)
					{
						sheetView.Cells[(int)iRow, num4 - 1].Value = num8;
					}
					else
					{
						sheetView.Cells[(int)iRow, num4 - 1].Value = 0;
					}
				}
				else
				{
					sheetView.Cells[(int)iRow, num4 - 1].Value = 0;
				}
				reloadCacheTuSoDenSo(sheetView, ToKhaiInfo, iRow);
			}
			if (iCol == num5)
			{
				string text = sheetView.Cells[(int)iRow, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow, (int)iCol - 1].Text = GeneralFunctions.FormatKyHieu(text);
			}
			if (iCol == num6)
			{
				string text4 = sheetView.Cells[(int)iRow, num7 - 1].Text;
				string text5 = "";
				if (!string.IsNullOrEmpty(text4))
				{
					text5 = dicDMHoaDon[text4][0];
				}
				if (!text5.Equals("3"))
				{
					string text = sheetView.Cells[(int)iRow, num6 - 1].Text;
					sheetView.Cells[(int)iRow, num6 - 1].Text = text.ToUpper();
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
			int num = HTKKFile.ColumnLettersToIndex("H") - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("I") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("J") - 1;
			int num4 = HTKKFile.ColumnLettersToIndex("K") - 1;
			int num5 = HTKKFile.ColumnLettersToIndex("L") - 1;
			int num6 = HTKKFile.ColumnLettersToIndex("M") - 1;
			int num7 = HTKKFile.ColumnLettersToIndex("R") - 1;
			int num8 = HTKKFile.ColumnLettersToIndex("T") - 1;
			int num9 = HTKKFile.ColumnLettersToIndex("V") - 1;
			int num10 = HTKKFile.ColumnLettersToIndex("X") - 1;
			int num11 = HTKKFile.ColumnLettersToIndex("Y") - 1;
			if (idxSheet == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num12 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int num13 = num12 + hTKKCMCTableInfo.RowsValues.Count - 1;
				if (idxRow >= num12 && idxRow <= num13)
				{
					if (idxCol == num || idxCol == num2 || idxCol == num3 || idxCol == num4 || idxCol == num5 || idxCol == num6 || idxCol == num10 || idxCol == num11)
					{
						GeneralFunctions.CompareKeycode(e, 1);
					}
					if (idxCol == num7 || idxCol == num8 || idxCol == num9)
					{
						GeneralFunctions.CompareKeycode(e, 0);
					}
				}
			}
			if (idxSheet == 1)
			{
				int num14 = HTKKFile.ColumnLettersToIndex("V") - 1;
				int num15 = HTKKFile.ColumnLettersToIndex("Y") - 1;
				SheetView sheetView2 = winFpSpread.Sheets[1];
				HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
				HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo2.Tables[0];
				int num16 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
				int num17 = num16 + hTKKCMCTableInfo2.RowsValues.Count - 1;
				HTKKCMCTableInfo hTKKCMCTableInfo3 = hTKKCMCSheetInfo2.Tables[1];
				int num18 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart - 1;
				int num19 = num18 + hTKKCMCTableInfo3.RowsValues.Count - 1;
				if (((idxRow >= num16 && idxRow <= num17) || (idxRow >= num18 && idxRow <= num19)) && (idxCol == num14 || idxCol == num15))
				{
					GeneralFunctions.CompareKeycode(e, 1);
				}
			}
			if (idxSheet == 2)
			{
				int num20 = HTKKFile.ColumnLettersToIndex("V") - 1;
				int num21 = HTKKFile.ColumnLettersToIndex("Y") - 1;
				SheetView sheetView3 = winFpSpread.Sheets[2];
				HTKKCMCSheetInfo hTKKCMCSheetInfo3 = ToKhaiInfo.SheetInfos[2];
				HTKKCMCTableInfo hTKKCMCTableInfo4 = hTKKCMCSheetInfo3.Tables[0];
				int num22 = hTKKCMCTableInfo4.Location.DataLocationTemplateRowStart - 1;
				int num23 = num22 + hTKKCMCTableInfo4.RowsValues.Count - 1;
				if (idxRow >= num22 && idxRow <= num23 && (idxCol == num20 || idxCol == num21))
				{
					GeneralFunctions.CompareKeycode(e, 1);
				}
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

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
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
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[12, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[13, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[15, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[24, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[26, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text13 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[27, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text14 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[31, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text15 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text16 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text17 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[52, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text18 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[53, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text19 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[54, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text20 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[55, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int num = location.DataLocationTemplateRowStart - 1;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			List<int> list = new List<int>();
			bool flag = false;
			bool flag2 = false;
			bool flag3 = false;
			string[] array = GeneralFunctions.checkTrung(winFpSpread, num, "H", "I", "B", "D", "E", "F", "", isMST: false, 1);
			string[] array2 = GeneralFunctions.checkTrung(winFpSpread, num, "J", "K", "B", "D", "E", "F", "", isMST: false, 1);
			string[] array3 = GeneralFunctions.checkTrung(winFpSpread, num, "X", "Y", "B", "D", "E", "F", "", isMST: false, 1);
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			int num3 = HTKKFile.ColumnLettersToIndex("D");
			int num4 = HTKKFile.ColumnLettersToIndex("E");
			int num5 = HTKKFile.ColumnLettersToIndex("F");
			int num6 = HTKKFile.ColumnLettersToIndex("G");
			int num7 = HTKKFile.ColumnLettersToIndex("H");
			int num8 = HTKKFile.ColumnLettersToIndex("I");
			int num9 = HTKKFile.ColumnLettersToIndex("J");
			int num10 = HTKKFile.ColumnLettersToIndex("K");
			int num11 = HTKKFile.ColumnLettersToIndex("L");
			int num12 = HTKKFile.ColumnLettersToIndex("M");
			int num13 = HTKKFile.ColumnLettersToIndex("N");
			int num14 = HTKKFile.ColumnLettersToIndex("P");
			int num15 = HTKKFile.ColumnLettersToIndex("Q");
			int num16 = HTKKFile.ColumnLettersToIndex("R");
			int num17 = HTKKFile.ColumnLettersToIndex("S");
			int num18 = HTKKFile.ColumnLettersToIndex("T");
			int num19 = HTKKFile.ColumnLettersToIndex("U");
			int num20 = HTKKFile.ColumnLettersToIndex("V");
			int num21 = HTKKFile.ColumnLettersToIndex("X");
			int num22 = HTKKFile.ColumnLettersToIndex("Y");
			for (int i = 0; i < rowsValues.Count; i++)
			{
				string text21 = sheetView.Cells[num + i, num2 - 1].Text;
				string text22 = sheetView.Cells[num + i, num3 - 1].Text;
				string text23 = sheetView.Cells[num + i, num4 - 1].Text;
				string text24 = sheetView.Cells[num + i, num5 - 1].Text;
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, num + i + 1, num6, "16");
				string text25 = sheetView.Cells[num + i, num7 - 1].Text;
				string text26 = sheetView.Cells[num + i, num8 - 1].Text;
				string text27 = sheetView.Cells[num + i, num9 - 1].Text;
				string text28 = sheetView.Cells[num + i, num10 - 1].Text;
				string text29 = sheetView.Cells[num + i, num11 - 1].Text;
				string text30 = sheetView.Cells[num + i, num12 - 1].Text;
				string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, num + i + 1, num13, "16");
				string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, num + i + 1, num14, "16");
				string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, num + i + 1, num15, "16");
				string text31 = sheetView.Cells[num + i, num16 - 1].Text;
				string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, num + i + 1, num17, "16");
				string text32 = sheetView.Cells[num + i, num18 - 1].Text;
				string valueCustom6 = GeneralFunctions.GetValueCustom(sheetView, num + i + 1, num19, "16");
				string text33 = sheetView.Cells[num + i, num20 - 1].Text;
				string text34 = sheetView.Cells[num + i, num21 - 1].Text;
				string text35 = sheetView.Cells[num + i, num22 - 1].Text;
				string text36 = "";
				if (HTKKFormater.isNullString(text21) && HTKKFormater.isNullString(text23) && HTKKFormater.isNullString(text24) && ConvertHelpers.ConvertToDouble(text25) == 0.0 && ConvertHelpers.ConvertToDouble(text26) == 0.0 && ConvertHelpers.ConvertToDouble(text27) == 0.0 && ConvertHelpers.ConvertToDouble(text28) == 0.0 && ConvertHelpers.ConvertToDouble(valueCustom3) == 0.0 && ConvertHelpers.ConvertToDouble(text31) == 0.0 && ConvertHelpers.ConvertToDouble(text32) == 0.0 && ConvertHelpers.ConvertToDouble(text33) == 0.0)
				{
					HTKKFile.RemoveValidate(sheetView, num + i, num2 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num4 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num5 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num6 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num7 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num8 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num9 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num10 - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, num + i, num11 - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, num + i, num12 - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, num + i, num13 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num14 - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, num + i, num15 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num16 - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, num + i, num17 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num18 - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, num + i, num19 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num20 - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, num + i, num21 - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, num + i, num22 - 1);
				}
				else
				{
					GeneralFunctions.strMauSoHD_01GTKT = text21;
					HTKKFile.RemoveValidate(sheetView, num + i, num2 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num3 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num4 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num5 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num7 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num8 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num9 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num10 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num14 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num16 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num18 - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, num20 - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, num + i, num21 - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, num + i, num22 - 1);
					if (string.IsNullOrEmpty(text21))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2 - 1, text);
						list.Add(1);
					}
					if (string.IsNullOrEmpty(text22))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3 - 1, text);
						list.Add(1);
					}
					string strLoai = "";
					if (!string.IsNullOrEmpty(text21))
					{
						strLoai = dicDMHoaDon[text21][0];
					}
					if (string.IsNullOrEmpty(text23))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4 - 1, text);
						list.Add(1);
					}
					else
					{
						text36 = GeneralFunctions.CheckMauSoHD(text23, strLoai, text21);
						if (text36.Equals("1"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4 - 1, text14);
							list.Add(1);
						}
						else if (text36.Equals("2"))
						{
							if (text24.Length > 0)
							{
								if (text24.Substring(text24.Length - 1).Equals("E"))
								{
									if (GeneralFunctions.CheckSoLienHDDT(text23, strLoai, text21, text24).Equals("3"))
									{
										HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, num4 - 1, text17);
										list.Add(0);
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4 - 1, text4);
									list.Add(1);
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4 - 1, text4);
								list.Add(1);
							}
						}
						else if (text24.Length > 0 && text24.Substring(text24.Length - 1).Equals("E") && GeneralFunctions.CheckSoLienHDDT(text23, strLoai, text21, text24).Equals("3"))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, num4 - 1, text17);
							list.Add(0);
						}
					}
					if (string.IsNullOrEmpty(text24))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, text);
						list.Add(1);
					}
					else
					{
						text36 = GeneralFunctions.CheckSoHD(text24, strLoai);
						if (text36.Equals("1"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, text15);
							list.Add(1);
						}
						else if (text36.Equals("2"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, text5);
							list.Add(1);
						}
					}
					string[] array4 = GeneralFunctions.checkSoTonSoPH(text25, text26, text27, text28);
					if (array4[0].Equals("1"))
					{
						if (array4[1].Equals("1"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num7 - 1, text);
							list.Add(1);
						}
						else if (array4[1].Equals("2"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8 - 1, text);
							list.Add(1);
						}
						else if (array4[1].Equals("3"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num7 - 1, text3);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8 - 1, text3);
							list.Add(1);
						}
						else if (array4[1].Equals("4"))
						{
							bool flag4 = GeneralFunctions.isDuLieuHL(text31);
							if (!HTKKFormater.isNullString(text31))
							{
								if (flag4)
								{
									if (!HTKKFormater.isNullString(text29) && !HTKKFormater.isNullString(text30))
									{
										string text37 = GeneralFunctions.checkHDXoa(text29, text30, text31);
										if (text37.Equals("-1"))
										{
											HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text7);
											list.Add(1);
											flag = true;
										}
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text8);
									list.Add(1);
								}
							}
							flag4 = GeneralFunctions.isDuLieuHL(text32);
							if (!HTKKFormater.isNullString(text32))
							{
								if (flag4)
								{
									if (!HTKKFormater.isNullString(text29) && !HTKKFormater.isNullString(text30))
									{
										string text37 = GeneralFunctions.checkHDXoa(text29, text30, text32);
										if (text37.Equals("-1"))
										{
											HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text7);
											list.Add(1);
											flag3 = true;
										}
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text8);
									list.Add(1);
								}
							}
							flag4 = GeneralFunctions.isDuLieuHL(text33);
							if (!HTKKFormater.isNullString(text33))
							{
								if (flag4)
								{
									if (!HTKKFormater.isNullString(text29) && !HTKKFormater.isNullString(text30))
									{
										string text37 = GeneralFunctions.checkHDXoa(text29, text30, text33);
										if (text37.Equals("-1"))
										{
											HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text7);
											list.Add(1);
											flag2 = true;
										}
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text8);
									list.Add(1);
								}
							}
							if ((!string.IsNullOrEmpty(text31) || !string.IsNullOrEmpty(text32) || !string.IsNullOrEmpty(text33)) && !flag && !flag3 && !flag2)
							{
								string text37 = GeneralFunctions.checkTrungXoa(text31, text32, text33);
								if (text37.Equals("1"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text7);
									list.Add(1);
								}
								else if (text37.Equals("2"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text7);
									list.Add(1);
								}
								else if (text37.Equals("3"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text7);
									list.Add(1);
								}
								else if (text37.Equals("4"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text7);
									list.Add(1);
								}
							}
						}
					}
					else if (array4[0].Equals("2"))
					{
						if (array4[2].Equals("1"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9 - 1, text);
							list.Add(1);
						}
						else if (array4[2].Equals("2"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num10 - 1, text);
							list.Add(1);
						}
						else if (array4[2].Equals("3"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9 - 1, text3);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num10 - 1, text3);
							list.Add(1);
						}
						else if (array4[2].Equals("4"))
						{
							bool flag4 = GeneralFunctions.isDuLieuHL(text31);
							if (!HTKKFormater.isNullString(text31))
							{
								if (flag4)
								{
									if (!HTKKFormater.isNullString(text29) && !HTKKFormater.isNullString(text30))
									{
										string text37 = GeneralFunctions.checkHDXoa(text29, text30, text31);
										if (text37.Equals("-1"))
										{
											HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text7);
											list.Add(1);
											flag = true;
										}
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text8);
									list.Add(1);
								}
							}
							flag4 = GeneralFunctions.isDuLieuHL(text32);
							if (!HTKKFormater.isNullString(text32))
							{
								if (flag4)
								{
									if (!HTKKFormater.isNullString(text29) && !HTKKFormater.isNullString(text30))
									{
										string text37 = GeneralFunctions.checkHDXoa(text29, text30, text32);
										if (text37.Equals("-1"))
										{
											HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text7);
											list.Add(1);
											flag3 = true;
										}
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text8);
									list.Add(1);
								}
							}
							flag4 = GeneralFunctions.isDuLieuHL(text33);
							if (!HTKKFormater.isNullString(text33))
							{
								if (flag4)
								{
									if (!HTKKFormater.isNullString(text29) && !HTKKFormater.isNullString(text30))
									{
										string text37 = GeneralFunctions.checkHDXoa(text29, text30, text33);
										if (text37.Equals("-1"))
										{
											HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text7);
											list.Add(1);
											flag2 = true;
										}
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text8);
									list.Add(1);
								}
							}
							if ((!string.IsNullOrEmpty(text31) || !string.IsNullOrEmpty(text32) || !string.IsNullOrEmpty(text33)) && !flag && !flag3 && !flag2)
							{
								string text37 = GeneralFunctions.checkTrungXoa(text31, text32, text33);
								if (text37.Equals("1"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text7);
									list.Add(1);
								}
								else if (text37.Equals("2"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text7);
									list.Add(1);
								}
								else if (text37.Equals("3"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text7);
									list.Add(1);
								}
								else if (text37.Equals("4"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text7);
									list.Add(1);
								}
							}
						}
					}
					else if (array4[0].Equals("3"))
					{
						if (array4[1].Equals("1"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num7 - 1, text);
							list.Add(1);
						}
						else if (array4[1].Equals("2"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8 - 1, text);
							list.Add(1);
						}
						else if (array4[1].Equals("3"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num7 - 1, text3);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8 - 1, text3);
							list.Add(1);
						}
						if (array4[2].Equals("1"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9 - 1, text);
							list.Add(1);
						}
						else if (array4[2].Equals("2"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num10 - 1, text);
							list.Add(1);
						}
						else if (array4[2].Equals("3"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9 - 1, text3);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num10 - 1, text3);
							list.Add(1);
						}
						if (array4[3].Equals("1") && !text26.Equals("0000000"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8 - 1, text6);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9 - 1, text6);
							list.Add(1);
						}
						if (array4[2].Equals("4") && array4[1].Equals("4"))
						{
							bool flag4 = GeneralFunctions.isDuLieuHL(text31);
							if (!HTKKFormater.isNullString(text31))
							{
								if (flag4)
								{
									if (!HTKKFormater.isNullString(text29) && !HTKKFormater.isNullString(text30))
									{
										string text37 = GeneralFunctions.checkHDXoa(text29, text30, text31);
										if (text37.Equals("-1"))
										{
											HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text7);
											list.Add(1);
											flag = true;
										}
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text8);
									list.Add(1);
								}
							}
							flag4 = GeneralFunctions.isDuLieuHL(text32);
							if (!HTKKFormater.isNullString(text32))
							{
								if (flag4)
								{
									if (!HTKKFormater.isNullString(text29) && !HTKKFormater.isNullString(text30))
									{
										string text37 = GeneralFunctions.checkHDXoa(text29, text30, text32);
										if (text37.Equals("-1"))
										{
											HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text7);
											list.Add(1);
											flag3 = true;
										}
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text8);
									list.Add(1);
								}
							}
							flag4 = GeneralFunctions.isDuLieuHL(text33);
							if (!HTKKFormater.isNullString(text33))
							{
								if (flag4)
								{
									if (!HTKKFormater.isNullString(text29) && !HTKKFormater.isNullString(text30))
									{
										string text37 = GeneralFunctions.checkHDXoa(text29, text30, text33);
										if (text37.Equals("-1"))
										{
											HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text7);
											list.Add(1);
											flag2 = true;
										}
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text8);
									list.Add(1);
								}
							}
							if ((!string.IsNullOrEmpty(text31) || !string.IsNullOrEmpty(text32) || !string.IsNullOrEmpty(text33)) && !flag && !flag3 && !flag2)
							{
								string text37 = GeneralFunctions.checkTrungXoa(text31, text32, text33);
								if (text37.Equals("1"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text7);
									list.Add(1);
								}
								else if (text37.Equals("2"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text7);
									list.Add(1);
								}
								else if (text37.Equals("3"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text7);
									list.Add(1);
								}
								else if (text37.Equals("4"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text7);
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text7);
									list.Add(1);
								}
							}
						}
					}
					string text38 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AA") - 1].Text;
					string text39 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AC") - 1].Text;
					string text40 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AD") - 1].Text;
					if (!HTKKFormater.isNullString(text25) && !HTKKFormater.isNullString(text26) && !HTKKFormater.isNullString(text39) && !HTKKFormater.isNullString(text40) && !HTKKFormater.isNullString(text21) && !HTKKFormater.isNullString(text38) && isExistsDlKyTruoc && text21 == text38 && (!text25.Equals(text39) || !text26.Equals(text40)))
					{
						if (sheetView.Cells[num + i, num7 - 1].BackColor != Color.Pink)
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, num7 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_49")));
						}
						if (sheetView.Cells[num + i, num8 - 1].BackColor != Color.Pink)
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, num8 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_49")));
						}
						list.Add(2);
					}
					if (ConvertHelpers.ConvertToInt(valueCustom3) > ConvertHelpers.ConvertToInt(valueCustom) && !valueCustom3.Equals(""))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num14 - 1, text9);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, num14 - 1);
						list.Add(0);
					}
					int num23 = ConvertHelpers.ConvertToInt(valueCustom);
					if (!text31.Trim().Equals(""))
					{
						int num24 = ConvertHelpers.ConvertToInt(GeneralFunctions.GetSoLuongHD(text31));
						if (num24 > num23)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text11);
							list.Add(1);
						}
					}
					if (!text32.Trim().Equals(""))
					{
						int num25 = ConvertHelpers.ConvertToInt(GeneralFunctions.GetSoLuongHD(text32));
						if (num25 > num23)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text12);
							list.Add(1);
						}
					}
					if (!text33.Trim().Equals(""))
					{
						int num26 = ConvertHelpers.ConvertToInt(GeneralFunctions.GetSoLuongHD(text33));
						if (num26 > num23)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text13);
							list.Add(1);
						}
					}
					int num27 = ConvertHelpers.ConvertToInt(valueCustom2);
					if (num27 > num23)
					{
						if (ConvertHelpers.ConvertToInt(valueCustom3) != 0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num14 - 1, text10);
							list.Add(1);
						}
						if (!text31.Equals(""))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num16 - 1, text10);
							list.Add(1);
						}
						if (!text32.Equals(""))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num18 - 1, text10);
							list.Add(1);
						}
						if (!text33.Equals(""))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num20 - 1, text10);
							list.Add(1);
						}
					}
					if (!HTKKFormater.isNullString(text25) && !HTKKFormater.isNullString(text26) && array.Length > 1)
					{
						for (int j = 0; j < array.Length; j++)
						{
							if (num + i == Convert.ToInt32(array[j]))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num7 - 1, text2);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8 - 1, text2);
								list.Add(1);
								break;
							}
						}
					}
					if (!HTKKFormater.isNullString(text27) && !HTKKFormater.isNullString(text28) && array2.Length > 1)
					{
						for (int j = 0; j < array2.Length; j++)
						{
							if (num + i == Convert.ToInt32(array2[j]))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9 - 1, text2);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num10 - 1, text2);
								list.Add(1);
								break;
							}
						}
					}
					if (!HTKKFormater.isNullString(text34) && !HTKKFormater.isNullString(text35) && array3.Length > 1)
					{
						for (int j = 0; j < array3.Length; j++)
						{
							if (num + i == Convert.ToInt32(array3[j]))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num21 - 1, text2);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num22 - 1, text2);
								list.Add(1);
								break;
							}
						}
					}
				}
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellNgayBaoCao));
				string text41 = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
				if (string.IsNullOrEmpty(text41))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					list.Add(1);
				}
				else
				{
					HTKKFormater.IsDateTime(text41, "dd/MM/yyyy", out var dtDate);
					if (DateTime.Compare(dtDate, DateTime.Now) > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
						list.Add(1);
					}
					else if (DateTime.Compare(dtDate, DateTime.ParseExact(PeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture)) < 0)
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
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
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
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
					}
				}
			}
			for (int i = 0; i < list.Count; i++)
			{
				if (list[i] == 1)
				{
					value = 2;
					break;
				}
				if (list[i] == 2 && value != 2)
				{
					value = 1;
				}
			}
			if (value == 2 && winFpSpread.TabStrip[0].BackColor != Color.Pink)
			{
				winFpSpread.TabStrip[0].BackColor = Color.Pink;
			}
			else if (value == 1 && winFpSpread.TabStrip[0].BackColor == Color.White)
			{
				winFpSpread.TabStrip[0].BackColor = Color.Yellow;
			}
			else if (value == 0)
			{
				winFpSpread.TabStrip[0].BackColor = Color.White;
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
		string text13 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[49, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text14 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[47, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
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
			string[] array = GeneralFunctions.checkTrung_01(winFpSpread, startRow, "V", "Y", "B", "C", "O", "S", "", isMST: false, 2, "bb");
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
					string text15 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text;
					string text16 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text;
					string valueMauSoHD = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Text;
					string valueKyHieuHD = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text;
					string text17 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Text;
					string text18 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num6 - 1].Text;
					string text19 = "";
					string strLoai = "";
					if (HTKKFormater.isNullString(text15) && HTKKFormater.isNullString(valueMauSoHD) && HTKKFormater.isNullString(text16) && HTKKFormater.isNullString(valueKyHieuHD) && ConvertHelpers.ConvertToDouble(text17) == 0.0 && ConvertHelpers.ConvertToDouble(text18) == 0.0)
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num - 1);
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
					GeneralFunctions.strMauSoHD_01GTKT = text15;
					if (string.IsNullOrEmpty(text15))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num - 1, text8);
						list.Add(1);
					}
					if (string.IsNullOrEmpty(text16))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1, text8);
						list.Add(1);
					}
					if (string.IsNullOrEmpty(valueMauSoHD))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text8);
						list.Add(1);
					}
					else
					{
						if (!string.IsNullOrEmpty(text15))
						{
							strLoai = dicDMHoaDon[text15][0];
						}
						text19 = GeneralFunctions.CheckMauSoHD(valueMauSoHD, strLoai, text15);
						if (text19.Equals("1"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text3);
							list.Add(1);
						}
						else if (text19.Equals("2"))
						{
							if (valueKyHieuHD.Length > 0)
							{
								if (valueKyHieuHD.Substring(valueKyHieuHD.Length - 1).Equals("E"))
								{
									if (GeneralFunctions.CheckSoLienHDDT(valueMauSoHD, strLoai, text15, valueKyHieuHD).Equals("3"))
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
						else if (valueKyHieuHD.Length > 0 && valueKyHieuHD.Substring(valueKyHieuHD.Length - 1).Equals("E") && GeneralFunctions.CheckSoLienHDDT(valueMauSoHD, strLoai, text15, valueKyHieuHD).Equals("3"))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text10);
							list.Add(0);
						}
					}
					if (string.IsNullOrEmpty(valueKyHieuHD))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, text8);
						list.Add(1);
					}
					else
					{
						text19 = GeneralFunctions.CheckSoHD(valueKyHieuHD, strLoai);
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
					if (!HTKKFormater.isNullString(text17) && !HTKKFormater.isNullString(text18) && array.Length > 1)
					{
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
					List<HTKKCMCRowValue> list2 = ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Where((HTKKCMCRowValue p) => p.RowValue[2].Equals(valueMauSoHD) && p.RowValue[3].Equals(valueKyHieuHD)).ToList();
					if (list2.Count == 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text13);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text13);
						list.Add(1);
						continue;
					}
					bool flag = false;
					for (int num7 = 0; num7 < list2.Count; num7++)
					{
						if (ConvertHelpers.ConvertToDouble(list2[num7].RowValue[21]) <= ConvertHelpers.ConvertToDouble(text17) && ConvertHelpers.ConvertToDouble(list2[num7].RowValue[22]) >= ConvertHelpers.ConvertToDouble(text18))
						{
							flag = true;
							break;
						}
					}
					if (!flag && i == 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text14);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text14);
						list.Add(1);
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
		string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[31, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[35, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[36, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[55, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[53, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[53, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[54, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text13 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[49, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text14 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[47, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			List<int> list = new List<int>();
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int startRow = location.DataLocationTemplateRowStart - 1;
			string[] array = GeneralFunctions.checkTrung_01(winFpSpread, startRow, "V", "Y", "B", "C", "O", "S", "", isMST: false, 3, "aa");
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
					string text15 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text;
					string text16 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text;
					string valueCot3 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Text;
					string valueCot4 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text;
					string text17 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Text;
					string text18 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num6 - 1].Text;
					string text19 = "";
					if (HTKKFormater.isNullString(text15) && HTKKFormater.isNullString(valueCot3) && HTKKFormater.isNullString(valueCot4) && ConvertHelpers.ConvertToDouble(text17) == 0.0 && ConvertHelpers.ConvertToDouble(text18) == 0.0)
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
					if (string.IsNullOrEmpty(text15))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num - 1, text);
						list.Add(1);
					}
					if (string.IsNullOrEmpty(text16))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1, text);
						list.Add(1);
					}
					string strLoai = "";
					if (string.IsNullOrEmpty(valueCot3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text);
						list.Add(1);
					}
					else
					{
						if (!string.IsNullOrEmpty(text15))
						{
							strLoai = dicDMHoaDon[text15][0];
						}
						text19 = GeneralFunctions.CheckMauSoHD(valueCot3, strLoai, text15);
						if (text19.Equals("1"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text2);
							list.Add(1);
						}
						else if (text19.Equals("2"))
						{
							if (valueCot4.Length > 0)
							{
								if (valueCot4.Substring(valueCot4.Length - 1).Equals("E"))
								{
									if (GeneralFunctions.CheckSoLienHDDT(valueCot3, strLoai, text15, valueCot4).Equals("3"))
									{
										HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text11);
										list.Add(0);
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text3);
									list.Add(1);
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text3);
								list.Add(1);
							}
						}
						else if (valueCot4.Length > 0 && valueCot4.Substring(valueCot4.Length - 1).Equals("E") && GeneralFunctions.CheckSoLienHDDT(valueCot3, strLoai, text15, valueCot4).Equals("3"))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text11);
							list.Add(0);
						}
					}
					if (string.IsNullOrEmpty(valueCot4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, text);
						list.Add(1);
					}
					else
					{
						text19 = GeneralFunctions.CheckSoHD(valueCot4, strLoai);
						if (text19.Equals("1"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, text4);
							list.Add(1);
						}
						else if (text19.Equals("2"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, text5);
							list.Add(1);
						}
					}
					if (string.IsNullOrEmpty(text17) && string.IsNullOrEmpty(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text);
						list.Add(1);
					}
					else if (!string.IsNullOrEmpty(text17) && string.IsNullOrEmpty(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text);
						list.Add(1);
					}
					else if (string.IsNullOrEmpty(text17) && !string.IsNullOrEmpty(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text);
						list.Add(1);
					}
					else if (ConvertHelpers.ConvertToInt(text17) > ConvertHelpers.ConvertToInt(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text6);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text6);
						list.Add(1);
					}
					if (!HTKKFormater.isNullString(text17) && !HTKKFormater.isNullString(text18) && array.Length > 1)
					{
						for (int k = 0; k < array.Length; k++)
						{
							if (dataLocationTemplateRowStart + j - 1 == Convert.ToInt32(array[k]))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text10);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text10);
								list.Add(1);
								break;
							}
						}
					}
					List<HTKKCMCRowValue> list2 = ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Where((HTKKCMCRowValue p) => p.RowValue[2].Equals(valueCot3) && p.RowValue[3].Equals(valueCot4)).ToList();
					if (list2.Count == 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text13);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text13);
						list.Add(1);
						continue;
					}
					bool flag = false;
					for (int num7 = 0; num7 < list2.Count; num7++)
					{
						if (ConvertHelpers.ConvertToDouble(list2[num7].RowValue[21]) <= ConvertHelpers.ConvertToDouble(text17) && ConvertHelpers.ConvertToDouble(list2[num7].RowValue[22]) >= ConvertHelpers.ConvertToDouble(text18))
						{
							if (ConvertHelpers.ConvertToDouble(list2[num7].RowValue[21]) != ConvertHelpers.ConvertToDouble(text17))
							{
								HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_59")));
							}
							if (ConvertHelpers.ConvertToDouble(list2[num7].RowValue[22]) != ConvertHelpers.ConvertToDouble(text18))
							{
								HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_60")));
							}
							flag = true;
							break;
						}
					}
					if (!flag)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text14);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text14);
						list.Add(1);
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
			throw ex;
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
			if (dicDMHoaDonTmp.Count > 0)
			{
				return dicDMHoaDonTmp;
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
			dicDMHoaDon = dictionary;
			dicDMHoaDonTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void Refresh_1(FpSpread winFpSpread, int actionType, HTKKCMCToKhaiInfo ToKhaiInfo)
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
			{
				winFpSpread.Sheets[0].Cells[num3 - 1, num2 - 1].Value = false;
				winFpSpread.Sheets[0].Cells[num3 - 1, num - 1].Value = true;
				winFpSpread.Sheets[0].Cells[num3 - 1, num - 1].Locked = false;
				winFpSpread.Sheets[0].Cells[num3 - 1, num2 - 1].Locked = false;
				int num4 = ToKhaiInfo.InterfaceTemplateExcel.LastIndexOf('\\');
				string text = ToKhaiInfo.InterfaceTemplateExcel.Substring(0, num4 + 1);
				text += "Bangke_BC26_Mau_3_12.xls";
				GeneralFunctions.InsertHyperLink(winFpSpread.Sheets[1], 2, HTKKFile.ColumnLettersToIndex("B"), "Mẫu tải bảng kê", text);
				break;
			}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SheetTabClick(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_2(FpSpread winFpSpread, int actionType, HTKKCMCToKhaiInfo ToKhaiInfo)
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
			{
				winFpSpread.Sheets[0].Cells[num3 - 1, num2 - 1].Value = true;
				winFpSpread.Sheets[0].Cells[num3 - 1, num - 1].Value = false;
				int num4 = ToKhaiInfo.InterfaceTemplateExcel.LastIndexOf('\\');
				string text = ToKhaiInfo.InterfaceTemplateExcel.Substring(0, num4 + 1);
				text += "Bangke_BC26_Mau_3_10.xls";
				GeneralFunctions.InsertHyperLink(winFpSpread.Sheets[2], 2, HTKKFile.ColumnLettersToIndex("B"), "Mẫu tải bảng kê", text);
				break;
			}
			}
			winFpSpread.Sheets[0].Cells[num3 - 1, num - 1].Locked = true;
			winFpSpread.Sheets[0].Cells[num3 - 1, num2 - 1].Locked = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int num = location.DataLocationTemplateRowStart - 1;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				string text = "";
				int num2 = 0;
				string text2 = "0";
				string text3 = "0";
				string text4 = "0";
				string text5 = "0";
				string[] array = new string[1];
				int num3 = HTKKFile.ColumnLettersToIndex("C");
				int num4 = HTKKFile.ColumnLettersToIndex("E");
				int num5 = HTKKFile.ColumnLettersToIndex("F");
				int num6 = HTKKFile.ColumnLettersToIndex("G");
				int num7 = HTKKFile.ColumnLettersToIndex("H");
				int num8 = HTKKFile.ColumnLettersToIndex("I");
				int num9 = HTKKFile.ColumnLettersToIndex("J");
				int num10 = HTKKFile.ColumnLettersToIndex("K");
				int num11 = HTKKFile.ColumnLettersToIndex("L");
				int num12 = HTKKFile.ColumnLettersToIndex("M");
				int iCol = HTKKFile.ColumnLettersToIndex("N");
				int iCol2 = HTKKFile.ColumnLettersToIndex("P");
				int num13 = HTKKFile.ColumnLettersToIndex("Q");
				int num14 = HTKKFile.ColumnLettersToIndex("R");
				int num15 = HTKKFile.ColumnLettersToIndex("S");
				int num16 = HTKKFile.ColumnLettersToIndex("T");
				int num17 = HTKKFile.ColumnLettersToIndex("U");
				int num18 = HTKKFile.ColumnLettersToIndex("V");
				int num19 = HTKKFile.ColumnLettersToIndex("X");
				int num20 = HTKKFile.ColumnLettersToIndex("Y");
				int num21 = HTKKFile.ColumnLettersToIndex("Z");
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string text6 = sheetView.Cells[num + j, num3 - 1].Text;
					string text7 = "";
					text = sheetView.Cells[num + j, num5 - 1].Text;
					sheetView.Cells[num + j, num5 - 1].Text = GeneralFunctions.FormatKyHieu(text);
					text = sheetView.Cells[num + j, num7 - 1].Text;
					sheetView.Cells[num + j, num7 - 1].Text = GeneralFunctions.FormatSoHD(text);
					text = sheetView.Cells[num + j, num8 - 1].Text;
					sheetView.Cells[num + j, num8 - 1].Text = GeneralFunctions.FormatSoHD(text);
					text = sheetView.Cells[num + j, num9 - 1].Text;
					sheetView.Cells[num + j, num9 - 1].Text = GeneralFunctions.FormatSoHD(text);
					text = sheetView.Cells[num + j, num10 - 1].Text;
					sheetView.Cells[num + j, num10 - 1].Text = GeneralFunctions.FormatSoHD(text);
					text2 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, iCol2, "16");
					text3 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, num13, "16");
					text4 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, num15, "16");
					text5 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, num17, "16");
					string text8 = sheetView.Cells[num + j, num7 - 1].Text;
					string text9 = sheetView.Cells[num + j, num8 - 1].Text;
					string text10 = sheetView.Cells[num + j, num9 - 1].Text;
					string text11 = sheetView.Cells[num + j, num10 - 1].Text;
					array = GeneralFunctions.getTusoDenso(text8, text9, text10, text11);
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, iCol, "16");
					string str;
					if (!array[0].Equals("-1"))
					{
						if (ConvertHelpers.ConvertToInt(text5) + ConvertHelpers.ConvertToInt(text3) + ConvertHelpers.ConvertToInt(text4) + ConvertHelpers.ConvertToInt(text2) > 0)
						{
							sheetView.Cells[num + j, num11 - 1].Text = GeneralFunctions.FormatSoHD(array[0]);
							str = (ConvertHelpers.ConvertToInt(array[0]) + ConvertHelpers.ConvertToInt(valueCustom) - 1).ToString();
							sheetView.Cells[num + j, num12 - 1].Text = GeneralFunctions.FormatSoHD(str);
						}
						else
						{
							sheetView.Cells[num + j, num11 - 1].Text = "";
						}
						sheetView.Cells[num + j, num6 - 1].Value = ConvertHelpers.ConvertToDouble(array[2]);
					}
					str = sheetView.Cells[num + j, num12 - 1].Text;
					string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, num6, "16");
					if (ConvertHelpers.ConvertToInt(valueCustom2) == ConvertHelpers.ConvertToInt(valueCustom) && ConvertHelpers.ConvertToInt(text5) + ConvertHelpers.ConvertToInt(text3) + ConvertHelpers.ConvertToInt(text4) + ConvertHelpers.ConvertToInt(text2) > 0)
					{
						num2 = ConvertHelpers.ConvertToInt(valueCustom2) - ConvertHelpers.ConvertToInt(valueCustom);
						sheetView.Cells[num + j, num19 - 1].Text = "";
						sheetView.Cells[num + j, num20 - 1].Text = "";
						sheetView.Cells[num + j, num21 - 1].Value = ConvertHelpers.ConvertToDouble(num2.ToString());
					}
					else if (ConvertHelpers.ConvertToInt(valueCustom2) > ConvertHelpers.ConvertToInt(valueCustom) && ConvertHelpers.ConvertToInt(text5) + ConvertHelpers.ConvertToInt(text3) + ConvertHelpers.ConvertToInt(text4) + ConvertHelpers.ConvertToInt(text2) > 0)
					{
						num2 = ConvertHelpers.ConvertToInt(valueCustom2) - ConvertHelpers.ConvertToInt(valueCustom);
						sheetView.Cells[num + j, num19 - 1].Text = GeneralFunctions.FormatSoHD((ConvertHelpers.ConvertToInt(str) + 1).ToString());
						sheetView.Cells[num + j, num20 - 1].Text = GeneralFunctions.FormatSoHD((ConvertHelpers.ConvertToInt(str) + num2).ToString());
						sheetView.Cells[num + j, num21 - 1].Value = ConvertHelpers.ConvertToDouble(num2.ToString());
					}
					else if (ConvertHelpers.ConvertToInt(valueCustom2) > ConvertHelpers.ConvertToInt(valueCustom) && ConvertHelpers.ConvertToInt(text5) + ConvertHelpers.ConvertToInt(text3) + ConvertHelpers.ConvertToInt(text4) + ConvertHelpers.ConvertToInt(text2) == 0 && string.IsNullOrEmpty(text8) && string.IsNullOrEmpty(text9))
					{
						num2 = ConvertHelpers.ConvertToInt(valueCustom2);
						sheetView.Cells[num + j, num19 - 1].Text = GeneralFunctions.FormatSoHD(ConvertHelpers.ConvertToInt(text10).ToString());
						sheetView.Cells[num + j, num20 - 1].Text = GeneralFunctions.FormatSoHD(ConvertHelpers.ConvertToInt(text11).ToString());
						sheetView.Cells[num + j, num21 - 1].Value = ConvertHelpers.ConvertToDouble(num2.ToString());
					}
					else if (ConvertHelpers.ConvertToInt(valueCustom2) > ConvertHelpers.ConvertToInt(valueCustom) && ConvertHelpers.ConvertToInt(text5) + ConvertHelpers.ConvertToInt(text3) + ConvertHelpers.ConvertToInt(text4) + ConvertHelpers.ConvertToInt(text2) == 0 && string.IsNullOrEmpty(text10) && string.IsNullOrEmpty(text11))
					{
						num2 = int.Parse(valueCustom2);
						sheetView.Cells[num + j, num19 - 1].Text = GeneralFunctions.FormatSoHD(ConvertHelpers.ConvertToInt(text8).ToString());
						sheetView.Cells[num + j, num20 - 1].Text = GeneralFunctions.FormatSoHD(ConvertHelpers.ConvertToInt(text9).ToString());
						sheetView.Cells[num + j, num21 - 1].Value = ConvertHelpers.ConvertToDouble(num2.ToString());
					}
					else
					{
						num2 = int.Parse(valueCustom2);
						sheetView.Cells[num + j, num19 - 1].Text = GeneralFunctions.FormatSoHD(ConvertHelpers.ConvertToInt(text8).ToString());
						sheetView.Cells[num + j, num20 - 1].Text = GeneralFunctions.FormatSoHD(ConvertHelpers.ConvertToInt(text11).ToString());
						sheetView.Cells[num + j, num21 - 1].Value = ConvertHelpers.ConvertToDouble(num2.ToString());
					}
					text = sheetView.Cells[num + j, num14 - 1].Text;
					string soLuongHD = GeneralFunctions.GetSoLuongHD(text);
					sheetView.Cells[num + j, num13 - 1].Value = ConvertHelpers.ConvertToDouble(soLuongHD);
					text = sheetView.Cells[num + j, num16 - 1].Text;
					soLuongHD = GeneralFunctions.GetSoLuongHD(text);
					sheetView.Cells[num + j, num15 - 1].Value = ConvertHelpers.ConvertToDouble(soLuongHD);
					text = sheetView.Cells[num + j, num18 - 1].Text;
					soLuongHD = GeneralFunctions.GetSoLuongHD(text);
					sheetView.Cells[num + j, num17 - 1].Value = ConvertHelpers.ConvertToDouble(soLuongHD);
					text2 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, iCol2, "16");
					text3 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, num13, "16");
					text4 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, num15, "16");
					text5 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, num17, "16");
					string text12 = sheetView.Cells[num + j, num11 - 1].Text;
					valueCustom = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, iCol, "16");
					str = (ConvertHelpers.ConvertToInt(text12) + ConvertHelpers.ConvertToInt(valueCustom) - 1).ToString();
					if (ConvertHelpers.ConvertToInt(text5) + ConvertHelpers.ConvertToInt(text4) + ConvertHelpers.ConvertToInt(text3) + ConvertHelpers.ConvertToInt(text2) > 0)
					{
						sheetView.Cells[num + j, num12 - 1].Text = GeneralFunctions.FormatSoHD(str);
						if (!array[0].Equals("-1"))
						{
							sheetView.Cells[num + j, num11 - 1].Text = GeneralFunctions.FormatSoHD(array[0]);
						}
					}
					else
					{
						sheetView.Cells[num + j, num11 - 1].Text = "";
						sheetView.Cells[num + j, num12 - 1].Text = "";
					}
					valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, num6, "16");
					if (ConvertHelpers.ConvertToInt(valueCustom2) > ConvertHelpers.ConvertToInt(valueCustom) && ConvertHelpers.ConvertToInt(text5) + ConvertHelpers.ConvertToInt(text4) + ConvertHelpers.ConvertToInt(text3) + ConvertHelpers.ConvertToInt(text2) > 0)
					{
						num2 = ConvertHelpers.ConvertToInt(valueCustom2) - ConvertHelpers.ConvertToInt(valueCustom);
						sheetView.Cells[num + j, num19 - 1].Text = GeneralFunctions.FormatSoHD((ConvertHelpers.ConvertToInt(str) + 1).ToString());
						sheetView.Cells[num + j, num20 - 1].Text = GeneralFunctions.FormatSoHD((ConvertHelpers.ConvertToInt(str) + num2).ToString());
						sheetView.Cells[num + j, num21 - 1].Value = ConvertHelpers.ConvertToDouble(num2.ToString());
					}
					text12 = sheetView.Cells[num + j, num11 - 1].Text;
					valueCustom = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, iCol, "16");
					valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, num6, "16");
					str = (ConvertHelpers.ConvertToInt(text12) + ConvertHelpers.ConvertToInt(valueCustom) - 1).ToString();
					if (ConvertHelpers.ConvertToInt(text5) + ConvertHelpers.ConvertToInt(text4) + ConvertHelpers.ConvertToInt(text3) + ConvertHelpers.ConvertToInt(text2) > 0)
					{
						sheetView.Cells[num + j, num12 - 1].Text = GeneralFunctions.FormatSoHD(str);
						if (!array[0].Equals("-1"))
						{
							sheetView.Cells[num + j, num11 - 1].Text = GeneralFunctions.FormatSoHD(array[0]);
						}
					}
					else
					{
						sheetView.Cells[num + j, num11 - 1].Text = "";
						sheetView.Cells[num + j, num12 - 1].Text = "";
					}
					if (ConvertHelpers.ConvertToInt(valueCustom2) > ConvertHelpers.ConvertToInt(valueCustom) && ConvertHelpers.ConvertToInt(text5) + ConvertHelpers.ConvertToInt(text4) + ConvertHelpers.ConvertToInt(text3) + ConvertHelpers.ConvertToInt(text2) > 0)
					{
						num2 = ConvertHelpers.ConvertToInt(valueCustom2) - ConvertHelpers.ConvertToInt(valueCustom);
						sheetView.Cells[num + j, num19 - 1].Text = GeneralFunctions.FormatSoHD((ConvertHelpers.ConvertToInt(str) + 1).ToString());
						sheetView.Cells[num + j, num20 - 1].Text = GeneralFunctions.FormatSoHD((ConvertHelpers.ConvertToInt(str) + num2).ToString());
						sheetView.Cells[num + j, num21 - 1].Value = ConvertHelpers.ConvertToDouble(num2.ToString());
					}
					valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, num6, "16");
					valueCustom = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, iCol, "16");
					text2 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, iCol2, "16");
					text3 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, num13, "16");
					text4 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, num15, "16");
					text5 = GeneralFunctions.GetValueCustom(sheetView, num + j + 1, num17, "16");
					if (ConvertHelpers.ConvertToInt(valueCustom2) + ConvertHelpers.ConvertToInt(valueCustom) + ConvertHelpers.ConvertToInt(text2) + ConvertHelpers.ConvertToInt(text3) + ConvertHelpers.ConvertToInt(text5) + ConvertHelpers.ConvertToInt(text4) == 0)
					{
						num2 = 0;
						sheetView.Cells[num + j, num19 - 1].Text = "";
						sheetView.Cells[num + j, num20 - 1].Text = "";
						sheetView.Cells[num + j, num21 - 1].Value = ConvertHelpers.ConvertToDouble(num2.ToString());
					}
					else if (ConvertHelpers.ConvertToInt(valueCustom2) == ConvertHelpers.ConvertToInt(valueCustom) && ConvertHelpers.ConvertToInt(text5) + ConvertHelpers.ConvertToInt(text4) + ConvertHelpers.ConvertToInt(text3) + ConvertHelpers.ConvertToInt(text2) > 0)
					{
						sheetView.Cells[num + j, num19 - 1].Text = "";
						sheetView.Cells[num + j, num20 - 1].Text = "";
						sheetView.Cells[num + j, num21 - 1].Value = ConvertHelpers.ConvertToDouble(num2.ToString());
					}
					if (!string.IsNullOrEmpty(text6))
					{
						text7 = dicDMHoaDon[text6][0];
					}
					if (!text7.Equals("3"))
					{
						text = sheetView.Cells[num + j, num4 - 1].Text;
						sheetView.Cells[num + j, num4 - 1].Text = text.ToUpper();
					}
				}
			}
			SheetView sheetView2 = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
			int num22;
			int num23;
			int num24;
			int num25;
			int num26;
			int num27;
			int num28;
			if (hTKKCMCSheetInfo2 != null)
			{
				for (int i = 0; i < hTKKCMCSheetInfo2.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables[i];
					HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
					int num = location.DataLocationTemplateRowStart - 1;
					List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
					num22 = HTKKFile.ColumnLettersToIndex("V");
					num23 = HTKKFile.ColumnLettersToIndex("Y");
					num24 = HTKKFile.ColumnLettersToIndex("AA");
					num25 = HTKKFile.ColumnLettersToIndex("S");
					num26 = HTKKFile.ColumnLettersToIndex("O");
					num27 = HTKKFile.ColumnLettersToIndex("C");
					for (int j = 0; j < rowsValues.Count; j++)
					{
						num28 = 0;
						string text = sheetView2.Cells[num + j, num22 - 1].Text;
						sheetView2.Cells[num + j, num22 - 1].Text = GeneralFunctions.FormatSoHD(text);
						text = sheetView2.Cells[num + j, num23 - 1].Text;
						sheetView2.Cells[num + j, num23 - 1].Text = GeneralFunctions.FormatSoHD(text);
						string text13 = sheetView2.Cells[num + j, num22 - 1].Text;
						string text14 = sheetView2.Cells[num + j, num23 - 1].Text;
						if (!string.IsNullOrEmpty(text13) || !string.IsNullOrEmpty(text14))
						{
							num28 = ConvertHelpers.ConvertToInt(text14) - ConvertHelpers.ConvertToInt(text13) + 1;
							if (num28 >= 1)
							{
								sheetView2.Cells[num + j, num24 - 1].Value = num28;
							}
							else
							{
								sheetView2.Cells[num + j, num24 - 1].Value = 0;
							}
						}
						else
						{
							sheetView2.Cells[num + j, num24 - 1].Value = 0;
						}
						text = sheetView2.Cells[num + j, num25 - 1].Text;
						sheetView2.Cells[num + j, num25 - 1].Text = GeneralFunctions.FormatKyHieu(text);
						string text6 = sheetView2.Cells[num + j, num27 - 1].Text;
						string text7 = "";
						if (!string.IsNullOrEmpty(text6))
						{
							text7 = dicDMHoaDon[text6][0];
						}
						if (!text7.Equals("3"))
						{
							text = sheetView2.Cells[num + j, num26 - 1].Text;
							sheetView2.Cells[num + j, num26 - 1].Text = text.ToUpper();
						}
					}
				}
			}
			SheetView sheetView3 = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo3 = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo3 == null)
			{
				return;
			}
			num22 = HTKKFile.ColumnLettersToIndex("V");
			num23 = HTKKFile.ColumnLettersToIndex("Y");
			num24 = HTKKFile.ColumnLettersToIndex("AA");
			num25 = HTKKFile.ColumnLettersToIndex("S");
			num26 = HTKKFile.ColumnLettersToIndex("O");
			num27 = HTKKFile.ColumnLettersToIndex("C");
			num28 = 0;
			for (int i = 0; i < hTKKCMCSheetInfo3.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo3.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int num = location.DataLocationTemplateRowStart - 1;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					num28 = 0;
					string text = sheetView3.Cells[num + j, num22 - 1].Text;
					sheetView3.Cells[num + j, num22 - 1].Text = GeneralFunctions.FormatSoHD(text);
					text = sheetView3.Cells[num + j, num23 - 1].Text;
					sheetView3.Cells[num + j, num23 - 1].Text = GeneralFunctions.FormatSoHD(text);
					string text13 = sheetView3.Cells[num + j, num22 - 1].Text;
					string text14 = sheetView3.Cells[num + j, num23 - 1].Text;
					if (!string.IsNullOrEmpty(text13) || !string.IsNullOrEmpty(text14))
					{
						num28 = ConvertHelpers.ConvertToInt(text14) - ConvertHelpers.ConvertToInt(text13) + 1;
						if (num28 >= 1)
						{
							sheetView3.Cells[num + j, num24 - 1].Value = num28;
						}
						else
						{
							sheetView3.Cells[num + j, num24 - 1].Value = 0;
						}
					}
					else
					{
						sheetView3.Cells[num + j, num24 - 1].Value = 0;
					}
					text = sheetView3.Cells[num + j, num25 - 1].Text;
					sheetView3.Cells[num + j, num25 - 1].Text = GeneralFunctions.FormatKyHieu(text);
					string text6 = sheetView3.Cells[num + j, num27 - 1].Text;
					string text7 = "";
					if (!string.IsNullOrEmpty(text6))
					{
						text7 = dicDMHoaDon[text6][0];
					}
					if (!text7.Equals("3"))
					{
						text = sheetView3.Cells[num + j, num26 - 1].Text;
						sheetView3.Cells[num + j, num26 - 1].Text = text.ToUpper();
					}
				}
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
			if (PeriodInfo.KieuKy.Equals("Q"))
			{
				xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/maTKhai").InnerText = "102";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("PL_BK_01AC_01/cqtChuyenDen"));
			if (hTKKCMCCellInfo == null)
			{
				return;
			}
			string text = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
			if (!HTKKFormater.isNullString(text) || HTKKFormater.isNullString(hTKKCMCCellInfo.Value))
			{
				return;
			}
			foreach (KeyValuePair<string, string> item in HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy)
			{
				if (hTKKCMCCellInfo.Value.IndexOf(item.Key) > -1 || hTKKCMCCellInfo.Value.IndexOf(item.Value) > -1)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = item.Value;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void reloadCacheTuSoDenSo(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, long iRow)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("V");
			int num3 = HTKKFile.ColumnLettersToIndex("Y");
			int num4 = HTKKFile.ColumnLettersToIndex("AA");
			int num5 = HTKKFile.ColumnLettersToIndex("S");
			int num6 = HTKKFile.ColumnLettersToIndex("O");
			int num7 = HTKKFile.ColumnLettersToIndex("C");
			string text = sView.Cells[(int)iRow, num2 - 1].Text;
			string text2 = sView.Cells[(int)iRow, num3 - 1].Text;
			string text3 = sView.Cells[(int)iRow, num7 - 1].Value.ToString();
			string text4 = sView.Cells[(int)iRow, num - 1].Text;
			string text5 = sView.Cells[(int)iRow, num6 - 1].Text;
			string text6 = sView.Cells[(int)iRow, num5 - 1].Text;
			string key = text3 + "~" + text4 + "~" + text5 + "~" + text6;
			object value = sView.Cells[(int)iRow, num4 - 1].Value;
			string text7 = "";
			if (ToKhaiInfo.SheetInfos[0].dicVar.ContainsKey(key))
			{
				string[] array = ToKhaiInfo.SheetInfos[0].dicVar[key].ToString().Split(new char[1] { '~' });
				text7 = array[0] + "~" + array[1] + "~" + array[2] + "~" + text + "~" + text2 + "~" + value;
				ToKhaiInfo.SheetInfos[0].dicVar.Remove(key);
			}
			else
			{
				text7 = "#~#~#~" + text + "~" + text2 + "~" + value;
			}
			ToKhaiInfo.SheetInfos[0].dicVar.Add(key, text7);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
