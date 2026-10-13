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
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BlpBcAc26Nckt200
{
	private Dictionary<string, string> dicBLPTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicBLP = new Dictionary<string, string[]>();

	private string pathCellNguoiDaiDien = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/nguoiDaiDien";

	private string pathCellNgayBaoCao = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayBCao";

	private string pathCellNguoiLapBieu = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/nguoiLapBieu";

	private List<Dictionary<string, string>> rowValues;

	private int[] CellId_NgayLapBC => GeneralFunctions.ArrayCellId("N_32");

	private int idxCol_SoTT => HTKKFile.ColumnLettersToIndex("B") - 1;

	private int idxCol_MaBienLai => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_MaBienLai_hide => HTKKFile.ColumnLettersToIndex("AA") - 1;

	private int idxCol_TenBienLai => HTKKFile.ColumnLettersToIndex("D") - 1;

	private int idxCol_KyHieuMau => HTKKFile.ColumnLettersToIndex("E") - 1;

	private int idxCol_KyHieuBL => HTKKFile.ColumnLettersToIndex("F") - 1;

	private int idxCol_soTonMuaTrKy_tongSo => HTKKFile.ColumnLettersToIndex("G") - 1;

	private int idxCol_soTonDauKy_tuSo => HTKKFile.ColumnLettersToIndex("H") - 1;

	private int idxCol_soTonDauKy_denSo => HTKKFile.ColumnLettersToIndex("I") - 1;

	private int idxCol_muaTrongKy_tuSo => HTKKFile.ColumnLettersToIndex("J") - 1;

	private int idxCol_muaTrongKy_denSo => HTKKFile.ColumnLettersToIndex("K") - 1;

	private int idxCol_tongSoSuDung_tuSo => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxCol_tongSoSuDung_denSo => HTKKFile.ColumnLettersToIndex("M") - 1;

	private int idxCol_tongSoSuDung_cong => HTKKFile.ColumnLettersToIndex("N") - 1;

	private int idxCol_soDaSDung => HTKKFile.ColumnLettersToIndex("P") - 1;

	private int idxCol_xoaBo_soLuong => HTKKFile.ColumnLettersToIndex("Q") - 1;

	private int idxCol_xoaBo_so => HTKKFile.ColumnLettersToIndex("R") - 1;

	private int idxCol_mat_soLuong => HTKKFile.ColumnLettersToIndex("S") - 1;

	private int idxCol_mat_so => HTKKFile.ColumnLettersToIndex("T") - 1;

	private int idxCol_huy_soLuong => HTKKFile.ColumnLettersToIndex("U") - 1;

	private int idxCol_huy_so => HTKKFile.ColumnLettersToIndex("V") - 1;

	private int idxCol_tonCuoiKy_tuSo => HTKKFile.ColumnLettersToIndex("X") - 1;

	private int idxCol_tonCuoiKy_denSo => HTKKFile.ColumnLettersToIndex("Y") - 1;

	private int idxCol_tonCuoiKy_soLuong => HTKKFile.ColumnLettersToIndex("Z") - 1;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				sView.Cells[CellId_NgayLapBC[1], CellId_NgayLapBC[0]].Text = DateTime.Now.ToString("dd/MM/yyyy");
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				LayGiaTriTonKyTruoc(sView, PeriodInfo, HeaderInfo, ToKhaiInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
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
				CellChange_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i, idxCol_muaTrongKy_denSo);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = tokhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			List<string> list = new List<string>();
			list.Add("maBienLai");
			list.Add("kyHieuMauBienLai");
			list.Add("tenBienLai");
			list.Add("kyHieuBienLai");
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
				sViewTK.Cells[num + j, idxCol_SoTT].Text = (j + 1).ToString();
				sViewTK.Cells[num + j, idxCol_MaBienLai].Text = rowValues[j]["maBienLai"];
				sViewTK.Cells[num + j, idxCol_KyHieuMau].Text = rowValues[j]["kyHieuMauBienLai"];
				sViewTK.Cells[num + j, idxCol_TenBienLai].Text = rowValues[j]["tenBienLai"];
				sViewTK.Cells[num + j, idxCol_KyHieuBL].Text = rowValues[j]["kyHieuBienLai"];
				sViewTK.Cells[num + j, idxCol_soTonDauKy_tuSo].Text = rowValues[j]["tonCuoiKy_tuSo"];
				sViewTK.Cells[num + j, idxCol_soTonDauKy_denSo].Text = rowValues[j]["tonCuoiKy_denSo"];
				sViewTK.Cells[num + j, idxCol_soTonMuaTrKy_tongSo].Value = rowValues[j]["tonCuoiKy_soLuong"].ToNumber();
				sViewTK.Cells[num + j, idxCol_tonCuoiKy_tuSo].Text = rowValues[j]["tonCuoiKy_tuSo"];
				sViewTK.Cells[num + j, idxCol_tonCuoiKy_denSo].Text = rowValues[j]["tonCuoiKy_denSo"];
				sViewTK.Cells[num + j, idxCol_tonCuoiKy_soLuong].Value = rowValues[j]["tonCuoiKy_soLuong"].ToNumber();
				Row row = sViewTK.Rows[num + j];
				float preferredHeight = row.GetPreferredHeight();
				if (row.Height < preferredHeight - 5f || row.Height > preferredHeight + 5f)
				{
					row.Height = preferredHeight;
				}
			}
			sViewTK.Cells[num + rowValues.Count, idxCol_SoTT].Text = (rowValues.Count + 1).ToString();
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
				if (idxCol == idxCol_soTonDauKy_tuSo || idxCol == idxCol_soTonDauKy_denSo)
				{
					GeneralFunctions.CompareKeycode(e, 1);
				}
				if (idxCol == idxCol_muaTrongKy_tuSo || idxCol == idxCol_muaTrongKy_denSo)
				{
					GeneralFunctions.CompareKeycode(e, 1);
				}
				if (idxCol == idxCol_huy_so || idxCol == idxCol_mat_so || idxCol == idxCol_xoaBo_so)
				{
					GeneralFunctions.CompareKeycode(e, 0);
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
			if (iRow >= dataLocationTemplateRowStart && iRow <= num && iCol - 1 == idxCol_MaBienLai)
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

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 4 && idxCol == 1)
			{
				ShowBangKe("Bao_Cao_BC26_AC_BLP.xls");
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

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			string text = "0";
			int num3 = 0;
			int num4 = 0;
			int num5 = 0;
			int num6 = 0;
			int num7 = 0;
			int num8 = 0;
			int num9 = 0;
			int num10 = 0;
			string[] array = new string[1];
			string text2 = "";
			if (iRow - 1 < num || iRow - 1 > num2)
			{
				return;
			}
			if (iCol - 1 == idxCol_KyHieuBL)
			{
				string text3 = sheetView.Cells[(int)iRow - 1, idxCol_KyHieuBL].Text;
				sheetView.Cells[(int)iRow - 1, idxCol_KyHieuBL].Text = GeneralFunctions.FormatKyHieuBLP(text3);
			}
			if (iCol - 1 == idxCol_tongSoSuDung_tuSo || iCol - 1 == idxCol_tongSoSuDung_denSo || iCol - 1 == idxCol_soTonDauKy_tuSo || iCol - 1 == idxCol_soTonDauKy_denSo || iCol - 1 == idxCol_muaTrongKy_tuSo || iCol - 1 == idxCol_muaTrongKy_denSo || iCol - 1 == idxCol_soDaSDung || iCol - 1 == idxCol_xoaBo_so || iCol - 1 == idxCol_mat_so || iCol - 1 == idxCol_huy_so)
			{
				if (iCol - 1 == idxCol_soTonDauKy_tuSo || iCol - 1 == idxCol_soTonDauKy_denSo || iCol - 1 == idxCol_muaTrongKy_tuSo || iCol - 1 == idxCol_muaTrongKy_denSo)
				{
					string text3 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
					sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text3);
				}
				string obj = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_soDaSDung);
				string obj2 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_xoaBo_soLuong);
				string obj3 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_mat_soLuong);
				string obj4 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_huy_soLuong);
				string text4 = sheetView.Cells[(int)iRow - 1, idxCol_soTonDauKy_tuSo].Text;
				string text5 = sheetView.Cells[(int)iRow - 1, idxCol_soTonDauKy_denSo].Text;
				string text6 = sheetView.Cells[(int)iRow - 1, idxCol_muaTrongKy_tuSo].Text;
				string text7 = sheetView.Cells[(int)iRow - 1, idxCol_muaTrongKy_denSo].Text;
				array = GeneralFunctions.getTusoDenso(text4, text5, text6, text7);
				string obj5 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_tongSoSuDung_cong);
				string str;
				if (!array[0].Equals("-1"))
				{
					num3 = ConvertHelpers.ConvertToInt(obj);
					num6 = ConvertHelpers.ConvertToInt(obj4);
					num4 = ConvertHelpers.ConvertToInt(obj2);
					num5 = ConvertHelpers.ConvertToInt(obj3);
					if (num6 + num5 + num4 + num3 > 0)
					{
						sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_tuSo].Text = GeneralFunctions.FormatSoHD(array[0]);
						num7 = ConvertHelpers.ConvertToInt(array[0]) + ConvertHelpers.ConvertToInt(obj5) - 1;
						str = num7.ToString();
						sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_denSo].Text = GeneralFunctions.FormatSoHD(str);
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_tuSo].Text = "";
					}
					sheetView.Cells[(int)iRow - 1, idxCol_soTonMuaTrKy_tongSo].Text = array[2];
				}
				str = sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_denSo].Text;
				obj5 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_tongSoSuDung_cong);
				string obj6 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_soTonMuaTrKy_tongSo);
				num9 = ConvertHelpers.ConvertToInt(obj6);
				num8 = ConvertHelpers.ConvertToInt(obj5);
				if (num9 == num8 && num6 + num5 + num4 + num3 > 0)
				{
					num10 = num9 - num8;
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Text = num10.ToString();
				}
				else if (num9 > num8 && num6 + num5 + num4 + num3 > 0)
				{
					num10 = num9 - num8;
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = GeneralFunctions.FormatSoHD((num7 + 1).ToString());
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = GeneralFunctions.FormatSoHD((num7 + num10).ToString());
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Text = num10.ToString();
				}
				else if (num9 > num8 && num6 + num5 + num4 + num3 > 0 && string.IsNullOrEmpty(text4) && string.IsNullOrEmpty(text5))
				{
					num10 = num9;
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = GeneralFunctions.FormatSoHD(text6);
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = GeneralFunctions.FormatSoHD(text7);
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Text = num10.ToString();
				}
				else if (num9 > num8 && num6 + num5 + num4 + num3 == 0 && string.IsNullOrEmpty(text6) && string.IsNullOrEmpty(text7))
				{
					num10 = num9;
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = GeneralFunctions.FormatSoHD(text4);
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = GeneralFunctions.FormatSoHD(text5);
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Text = num10.ToString();
				}
				else
				{
					num10 = num9;
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = GeneralFunctions.FormatSoHD(text6);
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = GeneralFunctions.FormatSoHD(text7);
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Text = num10.ToString();
				}
			}
			if (iCol - 1 == idxCol_xoaBo_so || iCol - 1 == idxCol_mat_so || iCol - 1 == idxCol_huy_so || iCol - 1 == idxCol_soDaSDung)
			{
				string text3 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				text = GeneralFunctions.GetSoLuongHD(text3);
				if (iCol - 1 == idxCol_xoaBo_so)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_xoaBo_soLuong].Text = text;
				}
				if (iCol - 1 == idxCol_mat_so)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_mat_soLuong].Text = text;
				}
				if (iCol - 1 == idxCol_huy_so)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_huy_soLuong].Text = text;
				}
				string obj = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_soDaSDung);
				string obj2 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_xoaBo_soLuong);
				string obj3 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_mat_soLuong);
				string obj4 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_huy_soLuong);
				num3 = ConvertHelpers.ConvertToInt(obj);
				num6 = ConvertHelpers.ConvertToInt(obj4);
				num5 = ConvertHelpers.ConvertToInt(obj3);
				num4 = ConvertHelpers.ConvertToInt(obj2);
				string obj5;
				if (num6 + num5 + num4 + num3 > 0)
				{
					if (!array[0].Equals("-1"))
					{
						sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_tuSo].Text = GeneralFunctions.FormatSoHD(array[0]);
					}
					string text8 = sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_tuSo].Text;
					obj5 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_tongSoSuDung_cong);
					num7 = ConvertHelpers.ConvertToInt(text8) + ConvertHelpers.ConvertToInt(obj5) - 1;
					string str = num7.ToString();
					sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_denSo].Text = GeneralFunctions.FormatSoHD(str);
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_denSo].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_tuSo].Text = "";
				}
				string obj6 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_soTonMuaTrKy_tongSo);
				obj5 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_tongSoSuDung_cong);
				num9 = ConvertHelpers.ConvertToInt(obj6);
				num8 = ConvertHelpers.ConvertToInt(obj5);
				if (num9 > num8 && num6 + num5 + num4 + num3 > 0)
				{
					num10 = num9 - num8;
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = GeneralFunctions.FormatSoHD((num7 + 1).ToString());
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = GeneralFunctions.FormatSoHD((num7 + num10).ToString());
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Text = num10.ToString();
				}
			}
			if (iCol - 1 == idxCol_soDaSDung)
			{
				string text8 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_tongSoSuDung_tuSo);
				string obj5 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_tongSoSuDung_cong);
				string obj6 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_soTonMuaTrKy_tongSo);
				int num11 = ConvertHelpers.ConvertToInt(text8);
				num8 = ConvertHelpers.ConvertToInt(obj5);
				num7 = num11 + num8 - 1;
				string str = num7.ToString();
				if (num6 + num5 + num4 + num3 > 0)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_denSo].Text = GeneralFunctions.FormatSoHD(str);
					if (!array[0].Equals("-1"))
					{
						sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_tuSo].Text = GeneralFunctions.FormatSoHD(array[0]);
					}
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_denSo].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_tuSo].Text = "";
				}
				if (num9 > num8 && num6 + num5 + num4 + num3 > 0)
				{
					num10 = num9 - num8;
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = GeneralFunctions.FormatSoHD((num7 + 1).ToString());
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = GeneralFunctions.FormatSoHD((num7 + num10).ToString());
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Text = num10.ToString();
				}
			}
			if (iCol - 1 == idxCol_soTonDauKy_tuSo || iCol - 1 == idxCol_soTonDauKy_denSo || iCol - 1 == idxCol_muaTrongKy_tuSo || iCol - 1 == idxCol_muaTrongKy_denSo || iCol - 1 == idxCol_soDaSDung || iCol - 1 == idxCol_xoaBo_so || iCol - 1 == idxCol_mat_so || iCol - 1 == idxCol_huy_so)
			{
				string obj6 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_soTonMuaTrKy_tongSo);
				string obj5 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_tongSoSuDung_cong);
				string obj = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_soDaSDung);
				string obj2 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_xoaBo_soLuong);
				string obj3 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_mat_soLuong);
				string obj4 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_huy_soLuong);
				num9 = ConvertHelpers.ConvertToInt(obj6);
				num8 = ConvertHelpers.ConvertToInt(obj5);
				num3 = ConvertHelpers.ConvertToInt(obj);
				num6 = ConvertHelpers.ConvertToInt(obj4);
				num5 = ConvertHelpers.ConvertToInt(obj3);
				num4 = ConvertHelpers.ConvertToInt(obj2);
				if (num9 + num8 + num3 + num4 + num6 + num5 == 0)
				{
					num10 = 0;
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Text = num10.ToString();
				}
				else if (num9 == num8 && num3 + num4 + num6 + num5 > 0)
				{
					num10 = 0;
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Text = num10.ToString();
				}
				else if (num9 < num8)
				{
					num10 = 0;
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Text = num10.ToString();
				}
			}
			if (iCol - 1 == idxCol_KyHieuMau)
			{
				string text3 = sheetView.Cells[(int)iRow - 1, idxCol_KyHieuMau].Text;
				sheetView.Cells[(int)iRow - 1, idxCol_KyHieuMau].Text = text3.ToUpper();
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
			checkTKChinh(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void checkTKChinh(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_21"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27"));
			string message9 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28"));
			string message10 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25"));
			bool flag = false;
			bool flag2 = false;
			bool flag3 = false;
			double num2 = 0.0;
			string[] array = GeneralFunctions.checkTrung(winFpSpread, num, "H", "I", "B", "D", "E", "F", "", isMST: false, 1);
			string[] array2 = GeneralFunctions.checkTrung(winFpSpread, num, "J", "K", "B", "D", "E", "F", "", isMST: false, 1);
			string[] array3 = GeneralFunctions.checkTrung(winFpSpread, num, "X", "Y", "B", "D", "E", "F", "", isMST: false, 1);
			string[] array4 = new string[25];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array4[0] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_MaBienLai);
				array4[1] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_TenBienLai);
				array4[2] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_KyHieuMau);
				array4[3] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_KyHieuBL);
				array4[4] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_soTonMuaTrKy_tongSo);
				array4[5] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_soTonDauKy_tuSo);
				array4[6] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_soTonDauKy_denSo);
				array4[7] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_muaTrongKy_tuSo);
				array4[8] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_muaTrongKy_denSo);
				array4[9] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_tongSoSuDung_tuSo);
				array4[10] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_tongSoSuDung_denSo);
				array4[11] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_tongSoSuDung_cong);
				array4[12] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_soDaSDung);
				array4[13] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_xoaBo_soLuong);
				array4[14] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_xoaBo_so);
				array4[15] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_mat_soLuong);
				array4[16] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_mat_so);
				array4[17] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_huy_soLuong);
				array4[18] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_huy_so);
				array4[19] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_tonCuoiKy_tuSo);
				array4[20] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_tonCuoiKy_denSo);
				array4[21] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, idxCol_tonCuoiKy_soLuong);
				array4[22] = "";
				array4[23] = "";
				array4[24] = "";
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array4))
				{
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_MaBienLai);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_TenBienLai);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_KyHieuMau);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_KyHieuBL);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_soTonDauKy_tuSo);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_soTonDauKy_denSo);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_muaTrongKy_tuSo);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_muaTrongKy_denSo);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_soDaSDung);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_mat_so);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_xoaBo_so);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_huy_so);
					continue;
				}
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_MaBienLai);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_TenBienLai);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_KyHieuMau);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_KyHieuBL);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_soTonDauKy_tuSo);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_soTonDauKy_denSo);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_muaTrongKy_tuSo);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_muaTrongKy_denSo);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_soDaSDung);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_mat_so);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_xoaBo_so);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_huy_so);
				Color backColor = sheetView.Cells[num + i, idxCol_tonCuoiKy_soLuong].BackColor;
				sheetView.Cells[num + i, idxCol_tonCuoiKy_tuSo].BackColor = backColor;
				sheetView.Cells[num + i, idxCol_tonCuoiKy_tuSo].Note = "";
				sheetView.Cells[num + i, idxCol_tonCuoiKy_denSo].BackColor = backColor;
				sheetView.Cells[num + i, idxCol_tonCuoiKy_denSo].Note = "";
				if (HTKKFormater.isNullString(array4[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_MaBienLai, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				if (HTKKFormater.isNullString(array4[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_TenBienLai, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				if (HTKKFormater.isNullString(array4[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_KyHieuMau, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				else
				{
					string strLoai = dicBLP[array4[0]][0];
					string text = GeneralFunctions.CheckMauSoBLP(array4[2], strLoai, array4[0]);
					if (text.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_KyHieuMau, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
						value = 2;
					}
					else if (text.Equals("2"))
					{
						if (array4[3].Length > 0)
						{
							if (array4[3].Substring(array4[3].Length - 1).Equals("E"))
							{
								if (GeneralFunctions.CheckSoLienBLPDT(array4[2], array4[0], array4[3]).Equals("3"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_KyHieuMau, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_53")));
									value = 2;
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_KyHieuMau, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
								value = 2;
							}
						}
					}
					else if (array4[3].Length > 0 && array4[3].Substring(array4[3].Length - 1).Equals("E") && GeneralFunctions.CheckSoLienBLPDT(array4[2], array4[0], array4[3]).Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_KyHieuMau, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_53")));
						value = 2;
					}
				}
				if (HTKKFormater.isNullString(array4[3]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_KyHieuBL, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				else
				{
					string text2 = GeneralFunctions.CheckSoBLP(array4[3], array4[0], DateTime.ParseExact(PeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture) >= new DateTime(2017, 1, 1));
					if (text2.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_KyHieuBL, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
						value = 2;
					}
					else if (text2.Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_KyHieuBL, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
						value = 2;
					}
				}
				string[] array5 = GeneralFunctions.checkSoTonSoPH(array4[5], array4[6], array4[7], array4[8]);
				if (array5[0].Equals("1"))
				{
					if (array5[1].Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soTonDauKy_tuSo, message);
						value = 2;
					}
					else if (array5[1].Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soTonDauKy_denSo, message);
						value = 2;
					}
					else if (array5[1].Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soTonDauKy_tuSo, message2);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soTonDauKy_denSo, message2);
						value = 2;
					}
					else if (array5[1].Equals("4"))
					{
						bool flag4 = GeneralFunctions.isDuLieuHL(array4[14]);
						if (!HTKKFormater.isNullString(array4[14]))
						{
							if (flag4)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[14]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message5);
										value = 2;
										flag = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message6);
								value = 2;
							}
						}
						flag4 = GeneralFunctions.isDuLieuHL(array4[16]);
						if (!HTKKFormater.isNullString(array4[16]))
						{
							if (flag4)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[16]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message5);
										value = 2;
										flag3 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message6);
								value = 2;
							}
						}
						flag4 = GeneralFunctions.isDuLieuHL(array4[18]);
						if (!HTKKFormater.isNullString(array4[18]))
						{
							if (flag4)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[18]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message5);
										value = 2;
										flag2 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message6);
								value = 2;
							}
						}
						if ((!HTKKFormater.isNullString(array4[14]) || !HTKKFormater.isNullString(array4[16]) || !HTKKFormater.isNullString(array4[18])) && !flag2 && !flag3 && !flag)
						{
							string text3 = GeneralFunctions.checkTrungXoa(array4[14], array4[16], array4[18]);
							if (text3.Equals("1"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message5);
								value = 2;
							}
							else if (text3.Equals("2"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message5);
								value = 2;
							}
							else if (text3.Equals("3"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message5);
								value = 2;
							}
							else if (text3.Equals("4"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message5);
								value = 2;
							}
						}
					}
				}
				else if (array5[0].Equals("2"))
				{
					if (array5[2].Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_muaTrongKy_tuSo, message);
						value = 2;
					}
					else if (array5[2].Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_muaTrongKy_denSo, message);
						value = 2;
					}
					else if (array5[2].Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_muaTrongKy_tuSo, message2);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_muaTrongKy_denSo, message2);
						value = 2;
					}
					else if (array5[2].Equals("4"))
					{
						bool flag4 = GeneralFunctions.isDuLieuHL(array4[14]);
						if (!HTKKFormater.isNullString(array4[14]))
						{
							if (flag4)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[14]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message5);
										value = 2;
										flag3 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message6);
								value = 2;
							}
						}
						flag4 = GeneralFunctions.isDuLieuHL(array4[16]);
						if (!HTKKFormater.isNullString(array4[16]))
						{
							if (flag4)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[16]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message5);
										value = 2;
										flag3 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message6);
								value = 2;
							}
						}
						flag4 = GeneralFunctions.isDuLieuHL(array4[18]);
						if (!HTKKFormater.isNullString(array4[18]))
						{
							if (flag4)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[18]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message5);
										value = 2;
										flag3 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message6);
								value = 2;
							}
						}
						if ((!HTKKFormater.isNullString(array4[14]) || !HTKKFormater.isNullString(array4[16]) || !HTKKFormater.isNullString(array4[18])) && !flag2 && !flag3 && !flag)
						{
							string text3 = GeneralFunctions.checkTrungXoa(array4[14], array4[16], array4[18]);
							if (text3.Equals("1"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message5);
								value = 2;
							}
							else if (text3.Equals("2"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message5);
								value = 2;
							}
							else if (text3.Equals("3"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message5);
								value = 2;
							}
							else if (text3.Equals("4"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message5);
								value = 2;
							}
						}
					}
				}
				else if (array5[0].Equals("3"))
				{
					if (array5[1].Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soTonDauKy_tuSo, message);
						value = 2;
					}
					else if (array5[1].Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soTonDauKy_denSo, message);
						value = 2;
					}
					else if (array5[1].Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soTonDauKy_tuSo, message2);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soTonDauKy_denSo, message2);
						value = 2;
					}
					if (array5[2].Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_muaTrongKy_tuSo, message);
						value = 2;
					}
					else if (array5[2].Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_muaTrongKy_denSo, message);
						value = 2;
					}
					else if (array5[2].Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_muaTrongKy_tuSo, message);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_muaTrongKy_denSo, message);
						value = 2;
					}
					if (array5[3].Equals("1") && !array4[6].Equals("0000000"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soTonDauKy_denSo, message4);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_muaTrongKy_tuSo, message4);
						value = 2;
					}
					if (array5[1].Equals("4") && array5[2].Equals("4"))
					{
						bool flag4 = GeneralFunctions.isDuLieuHL(array4[14]);
						if (!HTKKFormater.isNullString(array4[14]))
						{
							if (flag4)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[14]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message5);
										value = 2;
										flag3 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message6);
								value = 2;
							}
						}
						flag4 = GeneralFunctions.isDuLieuHL(array4[16]);
						if (!HTKKFormater.isNullString(array4[16]))
						{
							if (flag4)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[16]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message5);
										value = 2;
										flag3 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message6);
								value = 2;
							}
						}
						flag4 = GeneralFunctions.isDuLieuHL(array4[18]);
						if (!HTKKFormater.isNullString(array4[18]))
						{
							if (flag4)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[18]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message5);
										value = 2;
										flag3 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message6);
								value = 2;
							}
						}
						if ((!HTKKFormater.isNullString(array4[14]) || !HTKKFormater.isNullString(array4[16]) || !HTKKFormater.isNullString(array4[18])) && !flag2 && !flag3 && !flag)
						{
							string text3 = GeneralFunctions.checkTrungXoa(array4[14], array4[16], array4[18]);
							if (text3.Equals("1"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message5);
								value = 2;
							}
							else if (text3.Equals("2"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message5);
								value = 2;
							}
							else if (text3.Equals("3"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message5);
								value = 2;
							}
							else if (text3.Equals("4"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message5);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message5);
								value = 2;
							}
						}
					}
				}
				double num3 = ConvertHelpers.ConvertToDouble(array4[4]);
				if (ConvertHelpers.ConvertToDouble(array4[12]) > num3)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soDaSDung, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
					value = 2;
				}
				if (!HTKKFormater.isNullString(array4[14]))
				{
					num2 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetSoLuongHD(array4[14]));
					if (num2 > num3)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message7);
						value = 2;
					}
				}
				if (!HTKKFormater.isNullString(array4[16]))
				{
					num2 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetSoLuongHD(array4[16]));
					if (num2 > num3)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message8);
						value = 2;
					}
				}
				if (!HTKKFormater.isNullString(array4[18]))
				{
					num2 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetSoLuongHD(array4[18]));
					if (num2 > num3)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message9);
						value = 2;
					}
				}
				if (ConvertHelpers.ConvertToDouble(array4[11]) > num3)
				{
					if (!HTKKFormater.isNullString(array4[12]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soDaSDung, message10);
					}
					if (!HTKKFormater.isNullString(array4[14]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_xoaBo_so, message10);
					}
					if (!HTKKFormater.isNullString(array4[16]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_mat_so, message10);
					}
					if (!HTKKFormater.isNullString(array4[18]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_huy_so, message10);
					}
					value = 2;
				}
				if (!HTKKFormater.isNullString(array4[5]) && !HTKKFormater.isNullString(array4[6]) && array.Length > 1)
				{
					for (int j = 0; j < array.Length; j++)
					{
						if (num + i == Convert.ToInt32(array[j]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soTonDauKy_tuSo, message3);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soTonDauKy_denSo, message3);
							value = 2;
							break;
						}
					}
				}
				if (!HTKKFormater.isNullString(array4[7]) && !HTKKFormater.isNullString(array4[8]) && array2.Length > 1)
				{
					for (int j = 0; j < array2.Length; j++)
					{
						if (num + i == Convert.ToInt32(array2[j]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_muaTrongKy_tuSo, message3);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_muaTrongKy_denSo, message3);
							value = 2;
							break;
						}
					}
				}
				if (!HTKKFormater.isNullString(array4[19]) && !HTKKFormater.isNullString(array4[20]) && array3.Length > 1)
				{
					for (int j = 0; j < array3.Length; j++)
					{
						if (num + i == Convert.ToInt32(array3[j]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_tonCuoiKy_tuSo, message3);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_tonCuoiKy_denSo, message3);
							value = 2;
							break;
						}
					}
				}
				if (array4[5].ToIntOrDefault(-1) == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soTonDauKy_tuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_54")));
					value = 2;
				}
				if (array4[6].ToIntOrDefault(-1) == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_soTonDauKy_denSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_54")));
					value = 2;
				}
				if (array4[7].ToIntOrDefault(-1) == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_muaTrongKy_tuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_54")));
					value = 2;
				}
				if (array4[8].ToIntOrDefault(-1) == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_muaTrongKy_denSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_54")));
					value = 2;
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellNgayBaoCao));
			string text4 = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
			if (string.IsNullOrEmpty(text4))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
				value = 2;
			}
			else
			{
				HTKKFormater.IsDateTime(text4, "dd/MM/yyyy", out var dtDate);
				if (DateTime.Compare(dtDate, DateTime.Now) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					value = 2;
				}
				else if (DateTime.Compare(dtDate, DateTime.ParseExact(PeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture)) < 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					value = 2;
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

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		//IL_0043: Unknown result type (might be due to invalid IL or missing references)
		//IL_0049: Expected O, but got Unknown
		//IL_00c0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c6: Expected O, but got Unknown
		//IL_013d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0143: Expected O, but got Unknown
		//IL_01ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c0: Expected O, but got Unknown
		//IL_0237: Unknown result type (might be due to invalid IL or missing references)
		//IL_023d: Expected O, but got Unknown
		//IL_02b4: Unknown result type (might be due to invalid IL or missing references)
		//IL_02ba: Expected O, but got Unknown
		try
		{
			XmlNode val = xmlNode.OwnerDocument.CreateNode((XmlNodeType)2, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance");
			val.Value = "true";
			XmlNodeList val2 = xmlNode.SelectNodes("//CTietBKeBienLai/soTonDauKy_tuSo");
			foreach (XmlNode item in val2)
			{
				XmlNode val3 = item;
				if (val3 != null && HTKKFormater.isNullString(val3.InnerText))
				{
					((XmlNamedNodeMap)val3.Attributes).SetNamedItem(val);
				}
			}
			val2 = xmlNode.SelectNodes("//CTietBKeBienLai/soTonDauKy_denSo");
			foreach (XmlNode item2 in val2)
			{
				XmlNode val3 = item2;
				if (val3 != null && HTKKFormater.isNullString(val3.InnerText))
				{
					((XmlNamedNodeMap)val3.Attributes).SetNamedItem(val);
				}
			}
			val2 = xmlNode.SelectNodes("//CTietBKeBienLai/muaTrongKy_tuSo");
			foreach (XmlNode item3 in val2)
			{
				XmlNode val3 = item3;
				if (val3 != null && HTKKFormater.isNullString(val3.InnerText))
				{
					((XmlNamedNodeMap)val3.Attributes).SetNamedItem(val);
				}
			}
			val2 = xmlNode.SelectNodes("//CTietBKeBienLai/muaTrongKy_denSo");
			foreach (XmlNode item4 in val2)
			{
				XmlNode val3 = item4;
				if (val3 != null && HTKKFormater.isNullString(val3.InnerText))
				{
					((XmlNamedNodeMap)val3.Attributes).SetNamedItem(val);
				}
			}
			val2 = xmlNode.SelectNodes("//CTietBKeBienLai/tongSoSuDung_tuSo");
			foreach (XmlNode item5 in val2)
			{
				XmlNode val3 = item5;
				if (val3 != null && HTKKFormater.isNullString(val3.InnerText))
				{
					((XmlNamedNodeMap)val3.Attributes).SetNamedItem(val);
				}
			}
			val2 = xmlNode.SelectNodes("//CTietBKeBienLai/tongSoSuDung_denSo");
			foreach (XmlNode item6 in val2)
			{
				XmlNode val3 = item6;
				if (val3 != null && HTKKFormater.isNullString(val3.InnerText))
				{
					((XmlNamedNodeMap)val3.Attributes).SetNamedItem(val);
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
