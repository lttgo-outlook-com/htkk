using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class CnkdBcSdhd01Nckt216
{
	private List<Dictionary<string, string>> rowValues;

	private int idxCol_MaHoaDon => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_KyHieuMau => HTKKFile.ColumnLettersToIndex("D") - 1;

	private int idxCol_KyHieuHD => HTKKFile.ColumnLettersToIndex("E") - 1;

	private int idxCol_soTonMuaTrKy_tongSo => HTKKFile.ColumnLettersToIndex("F") - 1;

	private int idxCol_soTonDauKy_tuSo => HTKKFile.ColumnLettersToIndex("G") - 1;

	private int idxCol_soTonDauKy_denSo => HTKKFile.ColumnLettersToIndex("H") - 1;

	private int idxCol_muaTrongKy_tuSo => HTKKFile.ColumnLettersToIndex("I") - 1;

	private int idxCol_muaTrongKy_denSo => HTKKFile.ColumnLettersToIndex("J") - 1;

	private int idxCol_tongSoSuDung_tuSo => HTKKFile.ColumnLettersToIndex("K") - 1;

	private int idxCol_tongSoSuDung_denSo => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxCol_tongSoSuDung_cong => HTKKFile.ColumnLettersToIndex("M") - 1;

	private int idxCol_soDaSDung => HTKKFile.ColumnLettersToIndex("O") - 1;

	private int idxCol_xoaBo_soLuong => HTKKFile.ColumnLettersToIndex("P") - 1;

	private int idxCol_xoaBo_so => HTKKFile.ColumnLettersToIndex("Q") - 1;

	private int idxCol_mat_soLuong => HTKKFile.ColumnLettersToIndex("R") - 1;

	private int idxCol_mat_so => HTKKFile.ColumnLettersToIndex("S") - 1;

	private int idxCol_huy_soLuong => HTKKFile.ColumnLettersToIndex("T") - 1;

	private int idxCol_huy_so => HTKKFile.ColumnLettersToIndex("U") - 1;

	private int idxCol_tonCuoiKy_tuSo => HTKKFile.ColumnLettersToIndex("W") - 1;

	private int idxCol_tonCuoiKy_denSo => HTKKFile.ColumnLettersToIndex("X") - 1;

	private int idxCol_tonCuoiKy_soLuong => HTKKFile.ColumnLettersToIndex("Y") - 1;

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 5 && idxCol == 1)
			{
				ShowBangKe("Bao_Cao_01BC_SDHD_CNKD.xls");
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				LayGiaTriTonKyTruoc(sView, PeriodInfo, HeaderInfo, ToKhaiInfo);
			}
			sView.Cells[13, HTKKFile.ColumnLettersToIndex("E") - 1].Text = PeriodInfo.TuNgay;
			sView.Cells[13, HTKKFile.ColumnLettersToIndex("F") - 1].Text = PeriodInfo.DenNgay;
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
			list.Add("null_3");
			list.Add("ct06");
			list.Add("ct07");
			list.Add("ct23");
			list.Add("ct24");
			list.Add("ct25");
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
				sViewTK.Cells[num + j, idxCol_MaHoaDon].Text = rowValues[j]["null_3"];
				sViewTK.Cells[num + j, idxCol_KyHieuMau].Text = rowValues[j]["ct06"];
				sViewTK.Cells[num + j, idxCol_KyHieuHD].Text = rowValues[j]["ct07"];
				sViewTK.Cells[num + j, idxCol_soTonDauKy_tuSo].Text = rowValues[j]["ct23"];
				sViewTK.Cells[num + j, idxCol_soTonDauKy_denSo].Text = rowValues[j]["ct24"];
				sViewTK.Cells[num + j, idxCol_soTonMuaTrKy_tongSo].Value = ConvertHelpers.ConvertToInt(rowValues[j]["ct25"]);
				Row row = sViewTK.Rows[num + j];
				float preferredHeight = row.GetPreferredHeight();
				if (row.Height < preferredHeight - 5f || row.Height > preferredHeight + 5f)
				{
					row.Height = preferredHeight;
				}
			}
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
				list.Add(dictionary2);
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return list;
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
			string formula = "";
			int num3 = 0;
			int num4 = 0;
			int num5 = 0;
			int num6 = 0;
			int num7 = 0;
			int num8 = 0;
			int num9 = 0;
			int num10 = 0;
			string[] array = new string[1];
			if (iRow - 1 >= num && iRow - 1 <= num2)
			{
				if (iCol - 1 == idxCol_KyHieuHD)
				{
					string text2 = sheetView.Cells[(int)iRow - 1, idxCol_KyHieuHD].Text;
					sheetView.Cells[(int)iRow - 1, idxCol_KyHieuHD].Text = GeneralFunctions.FormatKyHieu(text2);
				}
				if (iCol - 1 == idxCol_tongSoSuDung_tuSo || iCol - 1 == idxCol_tongSoSuDung_denSo || iCol - 1 == idxCol_soTonDauKy_tuSo || iCol - 1 == idxCol_soTonDauKy_denSo || iCol - 1 == idxCol_muaTrongKy_tuSo || iCol - 1 == idxCol_muaTrongKy_denSo || iCol - 1 == idxCol_soDaSDung || iCol - 1 == idxCol_xoaBo_so || iCol - 1 == idxCol_mat_so || iCol - 1 == idxCol_huy_so)
				{
					if (iCol - 1 == idxCol_soTonDauKy_tuSo || iCol - 1 == idxCol_soTonDauKy_denSo || iCol - 1 == idxCol_muaTrongKy_tuSo || iCol - 1 == idxCol_muaTrongKy_denSo)
					{
						string text2 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
						sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text2);
					}
					string obj = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_soDaSDung);
					string obj2 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_xoaBo_soLuong);
					string obj3 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_mat_soLuong);
					string obj4 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_huy_soLuong);
					string text3 = sheetView.Cells[(int)iRow - 1, idxCol_soTonDauKy_tuSo].Text;
					string text4 = sheetView.Cells[(int)iRow - 1, idxCol_soTonDauKy_denSo].Text;
					string text5 = sheetView.Cells[(int)iRow - 1, idxCol_muaTrongKy_tuSo].Text;
					string text6 = sheetView.Cells[(int)iRow - 1, idxCol_muaTrongKy_denSo].Text;
					array = GeneralFunctions.getTusoDenso(text3, text4, text5, text6);
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_tongSoSuDung_cong + 1, "NumberCellType");
					if (!array[0].Equals("-1"))
					{
						num3 = ConvertHelpers.ConvertToInt(obj);
						num6 = ConvertHelpers.ConvertToInt(obj4);
						num4 = ConvertHelpers.ConvertToInt(obj2);
						num5 = ConvertHelpers.ConvertToInt(obj3);
						if (num6 + num5 + num4 + num3 > 0)
						{
							sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_tuSo].Text = GeneralFunctions.FormatSoHD(array[0]);
							string str = (ConvertHelpers.ConvertToInt(array[0]) + ConvertHelpers.ConvertToInt(valueCustom) - 1).ToString();
							sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_denSo].Text = GeneralFunctions.FormatSoHD(str);
						}
						else
						{
							sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_tuSo].Text = "";
						}
						sheetView.Cells[(int)iRow - 1, idxCol_soTonMuaTrKy_tongSo].Value = ConvertHelpers.ConvertToInt(array[2]);
					}
				}
				if (iCol - 1 == idxCol_xoaBo_so || iCol - 1 == idxCol_mat_so || iCol - 1 == idxCol_huy_so || iCol - 1 == idxCol_soDaSDung)
				{
					string text2 = sheetView.Cells[(int)iRow - 1, idxCol_xoaBo_so].Text;
					text = GeneralFunctions.GetSoLuongHD(text2);
					sheetView.Cells[(int)iRow - 1, idxCol_xoaBo_soLuong].Value = text;
					text2 = sheetView.Cells[(int)iRow - 1, idxCol_mat_so].Text;
					text = GeneralFunctions.GetSoLuongHD(text2);
					sheetView.Cells[(int)iRow - 1, idxCol_mat_soLuong].Value = text;
					text2 = sheetView.Cells[(int)iRow - 1, idxCol_huy_so].Text;
					text = GeneralFunctions.GetSoLuongHD(text2);
					sheetView.Cells[(int)iRow - 1, idxCol_huy_soLuong].Value = text;
					string obj = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_soDaSDung);
					string obj2 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_xoaBo_soLuong);
					string obj3 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_mat_soLuong);
					string obj4 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_huy_soLuong);
					num3 = ConvertHelpers.ConvertToInt(obj);
					num6 = ConvertHelpers.ConvertToInt(obj4);
					num5 = ConvertHelpers.ConvertToInt(obj3);
					num4 = ConvertHelpers.ConvertToInt(obj2);
					if (num6 + num5 + num4 + num3 > 0)
					{
						if (!array[0].Equals("-1"))
						{
							sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_tuSo].Text = GeneralFunctions.FormatSoHD(array[0]);
						}
						string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_tongSoSuDung_tuSo + 1, "NumberCellType");
						string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_tongSoSuDung_cong + 1, "NumberCellType");
						string str = (ConvertHelpers.ConvertToInt(valueCustom2) + ConvertHelpers.ConvertToInt(valueCustom) - 1).ToString();
						sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_denSo].Text = GeneralFunctions.FormatSoHD(str);
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_denSo].Text = "";
						sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_tuSo].Text = "";
					}
				}
				if (iCol - 1 == idxCol_soDaSDung)
				{
					string valueCustom2 = sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_tuSo].Text;
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_tongSoSuDung_cong + 1, "NumberCellType");
					string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_soTonMuaTrKy_tongSo + 1, "NumberCellType");
					int num11 = ConvertHelpers.ConvertToInt(valueCustom2);
					num8 = ConvertHelpers.ConvertToInt(valueCustom);
					string str = (num11 + num8 - 1).ToString();
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
				}
				if (iCol - 1 == idxCol_soTonDauKy_tuSo || iCol - 1 == idxCol_soTonDauKy_denSo || iCol - 1 == idxCol_muaTrongKy_tuSo || iCol - 1 == idxCol_muaTrongKy_denSo || iCol - 1 == idxCol_soDaSDung || iCol - 1 == idxCol_xoaBo_so || iCol - 1 == idxCol_mat_so || iCol - 1 == idxCol_huy_so)
				{
					string text3 = sheetView.Cells[(int)iRow - 1, idxCol_soTonDauKy_tuSo].Text;
					string text4 = sheetView.Cells[(int)iRow - 1, idxCol_soTonDauKy_denSo].Text;
					string text5 = sheetView.Cells[(int)iRow - 1, idxCol_muaTrongKy_tuSo].Text;
					string text6 = sheetView.Cells[(int)iRow - 1, idxCol_muaTrongKy_denSo].Text;
					array = GeneralFunctions.getTusoDenso(text3, text4, text5, text6);
					string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_soTonMuaTrKy_tongSo + 1, "NumberCellType");
					string valueCustom2 = sheetView.Cells[(int)iRow - 1, idxCol_tongSoSuDung_tuSo].Text;
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_tongSoSuDung_cong + 1, "NumberCellType");
					string obj = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_soDaSDung);
					string obj2 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_xoaBo_soLuong);
					string obj3 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_mat_soLuong);
					string obj4 = ConvertHelpers.ConvertCellValueToText(sheetView, (int)iRow - 1, idxCol_huy_soLuong);
					num9 = ConvertHelpers.ConvertToInt(valueCustom3);
					num8 = ConvertHelpers.ConvertToInt(valueCustom);
					num3 = ConvertHelpers.ConvertToInt(obj);
					num6 = ConvertHelpers.ConvertToInt(obj4);
					num5 = ConvertHelpers.ConvertToInt(obj3);
					num4 = ConvertHelpers.ConvertToInt(obj2);
					int num11 = ConvertHelpers.ConvertToInt(valueCustom2);
					num8 = ConvertHelpers.ConvertToInt(valueCustom);
					num7 = num11 + num8 - 1;
					if (!array[0].Equals("-1"))
					{
						if (num9 == num8 && num6 + num5 + num4 + num3 > 0)
						{
							num10 = num9 - num8;
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = "";
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = "";
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Value = num10;
						}
						else if (num9 > num8 && num6 + num5 + num4 + num3 > 0)
						{
							num10 = num9 - num8;
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = GeneralFunctions.FormatSoHD((num7 + 1).ToString());
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = GeneralFunctions.FormatSoHD((num7 + num10).ToString());
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Value = num10;
						}
						else if (((num9 > num8 && num6 + num5 + num4 + num3 > 0) || num6 + num5 + num4 + num3 == 0) && string.IsNullOrEmpty(text3) && string.IsNullOrEmpty(text4))
						{
							num10 = num9;
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = GeneralFunctions.FormatSoHD(text5);
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = GeneralFunctions.FormatSoHD(text6);
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Value = num10;
						}
						else if (((num9 > num8 && num6 + num5 + num4 + num3 == 0) || num6 + num5 + num4 + num3 == 0) && string.IsNullOrEmpty(text5) && string.IsNullOrEmpty(text6))
						{
							num10 = num9;
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = GeneralFunctions.FormatSoHD(text3);
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = GeneralFunctions.FormatSoHD(text4);
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Value = num10;
						}
						else
						{
							num10 = num9;
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_tuSo].Text = GeneralFunctions.FormatSoHD(text3);
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_denSo].Text = GeneralFunctions.FormatSoHD(text6);
							sheetView.Cells[(int)iRow - 1, idxCol_tonCuoiKy_soLuong].Value = num10;
						}
					}
				}
				if (iCol - 1 == idxCol_KyHieuMau)
				{
					string text2 = sheetView.Cells[(int)iRow - 1, idxCol_KyHieuMau].Text;
					sheetView.Cells[(int)iRow - 1, idxCol_KyHieuMau].Text = text2.ToUpper();
				}
				HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("path_null_100"));
				if (hTKKCMCCellInfo == null)
				{
				}
			}
			if (iRow > num2 + 9 && iRow < num2 + 14 && iCol - 1 == idxCol_tongSoSuDung_tuSo)
			{
				if (iRow == num2 + 10)
				{
					formula = "ROUND(K" + iRow + "*0.01,0)";
				}
				if (iRow == num2 + 11)
				{
					formula = "ROUND(K" + iRow + "*0.05,0)";
				}
				if (iRow == num2 + 12)
				{
					formula = "ROUND(K" + iRow + "*0.03,0)";
				}
				if (iRow == num2 + 13)
				{
					formula = "ROUND(K" + iRow + "*0.02,0)";
				}
				sheetView.Cells[(int)iRow - 1, idxCol_xoaBo_soLuong].Formula = formula;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool KeyPress(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e)
	{
		try
		{
			if (idxSheet == 0)
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
				int num12 = HTKKFile.ColumnLettersToIndex("O");
				int num13 = HTKKFile.ColumnLettersToIndex("P");
				int num14 = HTKKFile.ColumnLettersToIndex("Q");
				int num15 = HTKKFile.ColumnLettersToIndex("R");
				int num16 = HTKKFile.ColumnLettersToIndex("S");
				int num17 = HTKKFile.ColumnLettersToIndex("T");
				int num18 = HTKKFile.ColumnLettersToIndex("U");
				int num19 = HTKKFile.ColumnLettersToIndex("W");
				int num20 = HTKKFile.ColumnLettersToIndex("X");
				int num21 = HTKKFile.ColumnLettersToIndex("Y");
				SheetView sheetView = winFpSpread.Sheets[0];
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num22 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int num23 = num22 + hTKKCMCTableInfo.RowsValues.Count - 1;
				if (idxRow >= num22 && idxRow <= num23)
				{
					if (idxCol == num14 - 1 || idxCol == num16 - 1 || idxCol == num18 - 1)
					{
						GeneralFunctions.CompareKeycode(e, 0);
					}
					if (idxCol == num5 - 1 || idxCol == num6 - 1 || idxCol == num7 - 1 || idxCol == num8 - 1)
					{
						GeneralFunctions.CompareKeycode(e, 1);
					}
				}
			}
			return false;
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
			if (winFpSpread.TabStrip[0].BackColor != Color.White && winFpSpread.TabStrip[1].BackColor != Color.White && winFpSpread.TabStrip[2].BackColor != Color.White && winFpSpread.TabStrip[3].BackColor != Color.White)
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
					result = false;
					winFpSpread.TabStrip[key].BackColor = Color.Yellow;
				}
				else
				{
					result = false;
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

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = 0;
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[12, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[13, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[15, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[24, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text13 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text14 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[26, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text15 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[27, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text16 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[53, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text17 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[54, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text18 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[52, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text19 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text20 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[28, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text21 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[51, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			string[] array = GeneralFunctions.checkTrung(winFpSpread, dataLocationTemplateRowStart - 1, "G", "H", "B", "C", "D", "E", "", isMST: false, 1);
			string[] array2 = GeneralFunctions.checkTrung(winFpSpread, dataLocationTemplateRowStart - 1, "I", "J", "B", "C", "D", "E", "", isMST: false, 1);
			string[] array3 = GeneralFunctions.checkTrung(winFpSpread, dataLocationTemplateRowStart - 1, "W", "X", "B", "C", "D", "E", "", isMST: false, 1);
			List<int> list = new List<int>();
			for (int i = 0; i < rowsValues.Count; i++)
			{
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
				int num13 = HTKKFile.ColumnLettersToIndex("O");
				int iCol = HTKKFile.ColumnLettersToIndex("P");
				int num14 = HTKKFile.ColumnLettersToIndex("Q");
				int iCol2 = HTKKFile.ColumnLettersToIndex("R");
				int num15 = HTKKFile.ColumnLettersToIndex("S");
				int iCol3 = HTKKFile.ColumnLettersToIndex("T");
				int num16 = HTKKFile.ColumnLettersToIndex("U");
				int num17 = HTKKFile.ColumnLettersToIndex("W");
				int num18 = HTKKFile.ColumnLettersToIndex("X");
				int iCol4 = HTKKFile.ColumnLettersToIndex("Y");
				string text22 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text;
				string text23 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Text;
				string text24 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Text;
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, num5, "16");
				string text25 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num6 - 1].Text;
				string text26 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num7 - 1].Text;
				string text27 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num8 - 1].Text;
				string text28 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num9 - 1].Text;
				string text29 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num10 - 1].Text;
				string text30 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num11 - 1].Text;
				string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, num12, "16");
				string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, num13, "16");
				string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, iCol, "16");
				string text31 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num14 - 1].Text;
				string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, iCol2, "16");
				string text32 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num15 - 1].Text;
				string valueCustom6 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, iCol3, "16");
				string text33 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num16 - 1].Text;
				string text34 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num17 - 1].Text;
				string text35 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num18 - 1].Text;
				string valueCustom7 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, iCol4, "16");
				string text36 = "";
				string text37 = "";
				bool flag = false;
				bool flag2 = false;
				bool flag3 = false;
				if (HTKKFormater.isNullString(text23) && HTKKFormater.isNullString(text24) && ConvertHelpers.ConvertToDouble(valueCustom) == 0.0 && HTKKFormater.isNullString(text25) && HTKKFormater.isNullString(text26) && HTKKFormater.isNullString(text27) && HTKKFormater.isNullString(text28) && ConvertHelpers.ConvertToDouble(valueCustom3) == 0.0 && HTKKFormater.isNullString(text31) && HTKKFormater.isNullString(text32) && HTKKFormater.isNullString(text33))
				{
					if (rowsValues.Count == 1)
					{
						sheetView.Cells[dataLocationTemplateRowStart - 1, num3 - 1].Text = "";
						ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues[0].RowValue[1] = "";
					}
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num5 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart + i - 1, num11 - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart + i - 1, num12 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num13 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1);
					continue;
				}
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num13 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1);
				HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart + i - 1, num17 - 1);
				HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart + i - 1, num18 - 1);
				GeneralFunctions.strMauSoHD_01GTKT = "02GTGT";
				if (HTKKFormater.isNullString(text23))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text);
					list.Add(1);
				}
				else
				{
					text36 = GeneralFunctions.CheckMauSoHD(text23, "0", text22);
					if (text36 == "1")
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text3);
						list.Add(1);
					}
					else if (text36 == "2")
					{
						if (text24.Length > 0)
						{
							if (text24.Substring(text24.Length - 1).Equals("E"))
							{
								if (GeneralFunctions.CheckSoLienHDDT(text23, "0", text22, text24).Equals("3"))
								{
									HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text18);
									list.Add(0);
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text4);
								list.Add(1);
							}
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text4);
							list.Add(1);
						}
					}
					else if (text24.Length > 0 && text24.Substring(text24.Length - 1).Equals("E") && GeneralFunctions.CheckSoLienHDDT(text23, "0", text22, text24).Equals("3"))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text18);
						list.Add(0);
					}
				}
				if (HTKKFormater.isNullString(text24))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1, text);
					list.Add(1);
				}
				else
				{
					text36 = GeneralFunctions.CheckSoHD(text24, "0");
					if (text36.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1, text5);
						list.Add(1);
					}
					else if (text36.Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1, text7);
						list.Add(1);
					}
				}
				string[] array4 = GeneralFunctions.checkSoTonSoPH(text25, text26, text27, text28);
				if (array4[0].Equals("1"))
				{
					if (array4[1].Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1, text);
						list.Add(1);
					}
					else if (array4[1].Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1, text);
						list.Add(1);
					}
					else if (array4[1].Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1, text6);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1, text6);
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
									text37 = GeneralFunctions.checkHDXoa(text29, text30, text31);
									if (text37.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text9);
										list.Add(1);
										flag2 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text10);
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
									text37 = GeneralFunctions.checkHDXoa(text29, text30, text32);
									if (text37.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text9);
										list.Add(1);
										flag = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text10);
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
									text37 = GeneralFunctions.checkHDXoa(text29, text30, text33);
									if (text37.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text9);
										list.Add(1);
										flag3 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text10);
								list.Add(1);
							}
						}
						if ((!HTKKFormater.isNullString(text31) || !HTKKFormater.isNullString(text32) || !HTKKFormater.isNullString(text33)) && !flag3 && !flag && !flag2)
						{
							text37 = GeneralFunctions.checkTrungXoa(text31, text32, text33);
							if (text37.Equals("1"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text9);
								list.Add(1);
							}
							else if (text37.Equals("2"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text9);
								list.Add(1);
							}
							else if (text37.Equals("3"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text9);
								list.Add(1);
							}
							else if (text37.Equals("4"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text9);
								list.Add(1);
							}
						}
					}
				}
				else if (array4[0].Equals("2"))
				{
					if (array4[2].Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, text);
						list.Add(1);
					}
					else if (array4[2].Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1, text);
						list.Add(1);
					}
					else if (array4[2].Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, text6);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1, text6);
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
									text37 = GeneralFunctions.checkHDXoa(text29, text30, text31);
									if (text37.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text9);
										list.Add(1);
										flag2 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text10);
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
									text37 = GeneralFunctions.checkHDXoa(text29, text30, text32);
									if (text37.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text9);
										list.Add(1);
										flag = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text10);
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
									text37 = GeneralFunctions.checkHDXoa(text29, text30, text33);
									if (text37.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text9);
										list.Add(1);
										flag3 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text10);
								list.Add(1);
							}
						}
						if ((!HTKKFormater.isNullString(text31) || !HTKKFormater.isNullString(text32) || !HTKKFormater.isNullString(text33)) && !flag3 && !flag && !flag2)
						{
							text37 = GeneralFunctions.checkTrungXoa(text31, text32, text33);
							if (text37.Equals("1"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text9);
								list.Add(1);
							}
							else if (text37.Equals("2"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text9);
								list.Add(1);
							}
							else if (text37.Equals("3"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text9);
								list.Add(1);
							}
							else if (text37.Equals("4"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text9);
								list.Add(1);
							}
						}
					}
				}
				else if (array4[0].Equals("3"))
				{
					if (array4[1].Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1, text);
						list.Add(1);
					}
					else if (array4[1].Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1, text);
						list.Add(1);
					}
					else if (array4[1].Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1, text6);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1, text6);
						list.Add(1);
					}
					if (array4[2].Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, text);
						list.Add(1);
					}
					else if (array4[2].Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1, text);
						list.Add(1);
					}
					else if (array4[2].Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, text6);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1, text6);
						list.Add(1);
					}
					if (array4[3].Equals("1") && !text26.Equals("0000000"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1, text8);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, text8);
						list.Add(1);
					}
					if (array4[1].Equals("4") && array4[2].Equals("4"))
					{
						bool flag4 = GeneralFunctions.isDuLieuHL(text31);
						if (!HTKKFormater.isNullString(text31))
						{
							if (flag4)
							{
								if (!HTKKFormater.isNullString(text29) && !HTKKFormater.isNullString(text30))
								{
									text37 = GeneralFunctions.checkHDXoa(text29, text30, text31);
									if (text37.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text9);
										list.Add(1);
										flag2 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text10);
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
									text37 = GeneralFunctions.checkHDXoa(text29, text30, text32);
									if (text37.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text9);
										list.Add(1);
										flag = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text10);
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
									text37 = GeneralFunctions.checkHDXoa(text29, text30, text33);
									if (text37.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text9);
										list.Add(1);
										flag3 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text10);
								list.Add(1);
							}
						}
						if ((!HTKKFormater.isNullString(text31) || !HTKKFormater.isNullString(text32) || !HTKKFormater.isNullString(text33)) && !flag3 && !flag && !flag2)
						{
							text37 = GeneralFunctions.checkTrungXoa(text31, text32, text33);
							if (text37.Equals("1"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text9);
								list.Add(1);
							}
							else if (text37.Equals("2"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text9);
								list.Add(1);
							}
							else if (text37.Equals("3"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text9);
								list.Add(1);
							}
							else if (text37.Equals("4"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text9);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text9);
								list.Add(1);
							}
						}
					}
				}
				if (ConvertHelpers.ConvertToInt(valueCustom3) > ConvertHelpers.ConvertToInt(valueCustom))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num13 - 1, text11);
					list.Add(1);
				}
				double num19 = ConvertHelpers.ConvertToDouble(valueCustom);
				double num20 = 0.0;
				if (!HTKKFormater.isNullString(text31))
				{
					num20 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetSoLuongHD(text31));
					if (num20 > num19)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text13);
						list.Add(1);
					}
				}
				if (!HTKKFormater.isNullString(text32))
				{
					num20 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetSoLuongHD(text32));
					if (num20 > num19)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text14);
						list.Add(1);
					}
				}
				if (!HTKKFormater.isNullString(text33))
				{
					num20 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetSoLuongHD(text33));
					if (num20 > num19)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text15);
						list.Add(1);
					}
				}
				if (ConvertHelpers.ConvertToDouble(valueCustom2) > num19)
				{
					if (!HTKKFormater.isNullString(valueCustom3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num13 - 1, text12);
					}
					if (!HTKKFormater.isNullString(text31))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num14 - 1, text12);
					}
					if (!HTKKFormater.isNullString(text32))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num15 - 1, text12);
					}
					if (!HTKKFormater.isNullString(text33))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num16 - 1, text12);
					}
					list.Add(1);
				}
				if (!HTKKFormater.isNullString(text25) && !HTKKFormater.isNullString(text26) && array.Length > 1)
				{
					for (int j = 0; j < array.Length; j++)
					{
						if (dataLocationTemplateRowStart + i - 1 == Convert.ToInt32(array[j]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1, text2);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1, text2);
							list.Add(1);
							break;
						}
					}
				}
				if (!HTKKFormater.isNullString(text27) && !HTKKFormater.isNullString(text28) && array2.Length > 1)
				{
					for (int j = 0; j < array2.Length; j++)
					{
						if (dataLocationTemplateRowStart + i - 1 == Convert.ToInt32(array2[j]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, text2);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1, text2);
							list.Add(1);
							break;
						}
					}
				}
				if (HTKKFormater.isNullString(text34) || HTKKFormater.isNullString(text35) || array3.Length <= 1)
				{
					continue;
				}
				for (int j = 0; j < array3.Length; j++)
				{
					if (dataLocationTemplateRowStart + i - 1 == Convert.ToInt32(array3[j]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num17 - 1, text2);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num18 - 1, text2);
						list.Add(1);
						break;
					}
				}
			}
			int num21 = HTKKFile.ColumnLettersToIndex("M");
			string formula = sheetView.Cells[13, HTKKFile.ColumnLettersToIndex("M") - 1].Formula;
			int num22 = ConvertHelpers.ConvertToInt(formula.Replace("M", ""));
			int num23 = num22 + 2;
			HTKKFile.RemoveValidate(sheetView, num22 - 1, num21 - 1);
			HTKKFile.RemoveValidate(sheetView, num23 - 1, num21 - 1);
			if (HTKKFormater.isNullString(sheetView.Cells[num22 - 1, num21 - 1].Text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num22 - 1, num21 - 1, text19);
				list.Add(1);
			}
			if (HTKKFormater.isNullString(sheetView.Cells[num23 - 1, num21 - 1].Text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num23 - 1, num21 - 1, text19);
				list.Add(1);
			}
			else
			{
				DateTime dateTime = ConvertHelpers.ConvertToDateTime(sheetView.Cells[num23 - 1, num21 - 1].Text, "dd/MM/yyyy");
				DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(PeriodInfo.TuNgay, "dd/MM/yyyy");
				if ((dateTime2 - dateTime).TotalDays > 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num23 - 1, num21 - 1, text21);
					list.Add(1);
				}
				if ((dateTime - DateTime.Now).TotalDays > 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num23 - 1, num21 - 1, text20);
					list.Add(1);
				}
			}
			for (int i = 0; i < list.Count; i++)
			{
				if (list[i] == 1)
				{
					num = 2;
					break;
				}
			}
			switch (num)
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
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					CellChange_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, dataLocationTemplateRowStart + j, idxCol_soDaSDung + 1);
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
