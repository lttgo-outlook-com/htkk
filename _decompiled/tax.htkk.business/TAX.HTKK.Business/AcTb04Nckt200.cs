using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class AcTb04Nckt200
{
	private Dictionary<string, string> dicDMLoaiTTTDTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMLoaiTTTD = new Dictionary<string, string[]>();

	private string pathCellNguoiDaiDien = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/nguoiDaiDien";

	private string pathCellNgayTBDC = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayThongBao";

	private string pathCellNgayTBPH = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayTBaoPHanhHDon";

	private int[] CellId_NgayDaCoTB => GeneralFunctions.ArrayCellId("F_30");

	private int idxCol_C => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_D => HTKKFile.ColumnLettersToIndex("D") - 1;

	private int idxCol_M => HTKKFile.ColumnLettersToIndex("M") - 1;

	private int idxCol_AA => HTKKFile.ColumnLettersToIndex("AA") - 1;

	private int[] CellId_DVCQTen => GeneralFunctions.ArrayCellId("F_40");

	private int[] CellId_DVCQMaST => GeneralFunctions.ArrayCellId("F_42");

	private int[] CellId_TenCQTTN => GeneralFunctions.ArrayCellId("K_44");

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellNgayTBDC));
			if (hTKKCMCCellInfo != null)
			{
				sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = PeriodInfo.kyKKhai;
			}
			int count = hTKKCMCSheetInfo.Tables[0].RowsValues.Count;
			sView.Cells[CellId_DVCQTen[1] + count - 1, CellId_DVCQTen[0]].Text = HeaderInfo.TenDonViChuQuan;
			sView.Cells[CellId_DVCQMaST[1] + count - 1, CellId_DVCQMaST[0]].Text = HeaderInfo.MaSoThueDonViChuQuan;
			if (!ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				sView.Cells[CellId_TenCQTTN[1] + count - 1, CellId_TenCQTTN[0]].Text = TTinChung.tenCQTNoiNop;
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
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_21"));
			if (iRow < dataLocationTemplateRowStart || iRow > num || iCol - 1 != idxCol_C)
			{
				return;
			}
			List<string> list = new List<string>();
			for (int i = dataLocationTemplateRowStart; i <= num; i++)
			{
				if (!HTKKFormater.isNullString(sheetView.Cells[i - 1, idxCol_AA].Text))
				{
					list.Add(sheetView.Cells[i - 1, idxCol_AA].Text);
				}
			}
			string text = sheetView.Cells[(int)iRow - 1, idxCol_C].Text;
			if (HTKKFormater.isNullString(text))
			{
				sheetView.Cells[(int)iRow - 1, idxCol_AA].ResetValue();
				return;
			}
			foreach (KeyValuePair<string, string> item in dicDMLoaiTTTDTmp)
			{
				if (item.Value.Equals(text))
				{
					if (!list.Contains(item.Key))
					{
						sheetView.Cells[(int)iRow - 1, idxCol_AA].Text = item.Key;
					}
					break;
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
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			string message = HTKKFile.GetMessage("0071");
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			string text = sheetView.Cells[CellId_NgayDaCoTB[1], CellId_NgayDaCoTB[0]].Text.Trim();
			DateTime dtDate;
			if (HTKKFormater.isNullString(text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_NgayDaCoTB[1], CellId_NgayDaCoTB[0], message2);
				value = 2;
			}
			else if (!HTKKFormater.IsDateTime(text, "dd/MM/yyyy", out dtDate))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_NgayDaCoTB[1], CellId_NgayDaCoTB[0], message);
				value = 2;
			}
			else
			{
				HTKKFormater.IsDateTime(PeriodInfo.kyKKhai, "dd/MM/yyyy", out var dtDate2);
				if (DateTime.Compare(dtDate, dtDate2) >= 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_NgayDaCoTB[1], CellId_NgayDaCoTB[0], message3);
					value = 2;
				}
				else if (dtDate >= new DateTime(1900, 1, 1))
				{
					HTKKFile.RemoveValidate(sheetView, CellId_NgayDaCoTB[1], CellId_NgayDaCoTB[0]);
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string[] array = new string[rowsInfo.Cells.Count];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sheetView.Cells[num + i, idxCol_C].Text.Trim();
				array[1] = sheetView.Cells[num + i, idxCol_D].Text.Trim();
				array[2] = sheetView.Cells[num + i, idxCol_M].Text.Trim();
				array[3] = sheetView.Cells[num + i, idxCol_AA].Text.Trim();
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array))
				{
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_D);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_M);
					continue;
				}
				if (HTKKFormater.isNullString(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_C, message2);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_C);
				}
				if (HTKKFormater.isNullString(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_D, message2);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_D);
				}
				if (HTKKFormater.isNullString(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_M, message2);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_M);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellNguoiDaiDien));
			if (hTKKCMCCellInfo != null)
			{
				string str = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
				if (HTKKFormater.isNullString(str))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message2);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
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

	public Dictionary<string, string> GetDMLoaiThongTinThayDoi(string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		try
		{
			if (dicDMLoaiTTTDTmp.Count > 0)
			{
				return dicDMLoaiTTTDTmp;
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
				string key = array[0];
				string value2 = array[1];
				dictionary2.Add(key, value2);
				dictionary.Add(key, array);
			}
			dicDMLoaiTTTD = dictionary;
			dicDMLoaiTTTDTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}
}
