using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BlpBk01BcAc26Nckt234
{
	private Dictionary<string, string> dicBLPTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicBLP = new Dictionary<string, string[]>();

	private string pathNguoiKy = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/nguoiKy";

	private string pathNgayKy = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayKy";

	private string pathNguoiLapBieu = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/nguoiLapBieu";

	private int idxCol_MaBienLai => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_MaBienLai_hide => HTKKFile.ColumnLettersToIndex("AI") - 1;

	private int idxCol_TenBienLai => HTKKFile.ColumnLettersToIndex("D") - 1;

	private int idxCol_KyHieuMau => HTKKFile.ColumnLettersToIndex("O") - 1;

	private int idxCol_KyHieuBL => HTKKFile.ColumnLettersToIndex("S") - 1;

	private int idxCol_TuSo => HTKKFile.ColumnLettersToIndex("V") - 1;

	private int idxCol_DenSo => HTKKFile.ColumnLettersToIndex("Y") - 1;

	private int idxCol_SoLuong => HTKKFile.ColumnLettersToIndex("AA") - 1;

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

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
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

	public bool KeyPress(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e)
	{
		try
		{
			if (idxSheet == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
				if (idxRow >= num && idxRow <= num2 && (idxCol == idxCol_TuSo || idxCol == idxCol_DenSo))
				{
					GeneralFunctions.CompareKeycode(e, 1);
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

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (iRow - 1 < num || iRow - 1 > num2)
			{
				return;
			}
			if (iCol - 1 == idxCol_KyHieuBL)
			{
				string text = sheetView.Cells[(int)iRow - 1, idxCol_KyHieuBL].Text;
				sheetView.Cells[(int)iRow - 1, idxCol_KyHieuBL].Text = GeneralFunctions.FormatKyHieuBLP(text);
			}
			else if (iCol - 1 == idxCol_KyHieuMau)
			{
				string text2 = sheetView.Cells[(int)iRow - 1, idxCol_KyHieuMau].Text;
				sheetView.Cells[(int)iRow - 1, idxCol_KyHieuMau].Text = text2.ToUpper();
			}
			else
			{
				if (iCol - 1 != idxCol_TuSo && iCol - 1 != idxCol_DenSo)
				{
					return;
				}
				string text3 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				if (!HTKKFormater.isNullString(text3))
				{
					sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text3);
				}
				string text4 = sheetView.Cells[(int)iRow - 1, idxCol_TuSo].Text;
				string text5 = sheetView.Cells[(int)iRow - 1, idxCol_DenSo].Text;
				if (!HTKKFormater.isNullString(text4) && !HTKKFormater.isNullString(text5))
				{
					double num3 = ConvertHelpers.ConvertToDouble(text4);
					double num4 = ConvertHelpers.ConvertToDouble(text5);
					sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].ResetFormula();
					if (num3 == 0.0 && num4 == 0.0)
					{
						sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = 0;
					}
					else if (num3 == 0.0 && num4 != 0.0)
					{
						sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = num4;
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = num4 - num3 + 1.0;
					}
				}
				else if (HTKKFormater.isNullString(text4) && HTKKFormater.isNullString(text5))
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

	public void isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				int value = item.Value;
				if (value == 2 && winFpSpread.TabStrip[key].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
				}
			}
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
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string[] array = GeneralFunctions.checkTrung(winFpSpread, num2, "V", "Y", "B", "E", "O", "S", "", isMST: false, 1);
			string[] array2 = new string[rowsInfo.Cells.Count];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array2[0] = sheetView.Cells[num2 + i, idxCol_MaBienLai].Text.Trim();
				array2[1] = sheetView.Cells[num2 + i, idxCol_TenBienLai].Text.Trim();
				array2[2] = sheetView.Cells[num2 + i, idxCol_KyHieuMau].Text.Trim();
				array2[3] = sheetView.Cells[num2 + i, idxCol_KyHieuBL].Text.Trim();
				array2[4] = sheetView.Cells[num2 + i, idxCol_TuSo].Text.Trim();
				array2[5] = sheetView.Cells[num2 + i, idxCol_DenSo].Text.Trim();
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array2))
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_MaBienLai);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_KyHieuMau);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TuSo);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_DenSo);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TenBienLai);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_KyHieuBL);
					continue;
				}
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_MaBienLai);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_KyHieuMau);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TuSo);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_DenSo);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TenBienLai);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_KyHieuBL);
				if (HTKKFormater.isNullString(array2[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_MaBienLai, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					num = 2;
				}
				if (HTKKFormater.isNullString(array2[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TenBienLai, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					num = 2;
				}
				if (HTKKFormater.isNullString(array2[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					num = 2;
				}
				else
				{
					string strLoai = (HTKKFormater.isNullString(array2[0]) ? "" : dicBLP[array2[0]][0]);
					string text = GeneralFunctions.CheckMauSoBLP(array2[2], strLoai, array2[0]);
					if (text.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
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
									HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_53")));
									num = 2;
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
								num = 2;
							}
						}
					}
					else if (array2[3].Length > 0 && GeneralFunctions.CheckSoLienBLPDT(array2[2], array2[0], array2[3]).Equals("3"))
					{
						if (array2[3].Substring(array2[3].Length - 1).Equals("E"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_53")));
							num = 2;
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
							num = 2;
						}
					}
				}
				if (HTKKFormater.isNullString(array2[3]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuBL, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					num = 2;
				}
				else
				{
					string text2 = GeneralFunctions.CheckSoBLP(array2[3], array2[0], DateTime.ParseExact(PeriodInfo.kyKKhai, "dd/MM/yyyy", CultureInfo.InvariantCulture) >= new DateTime(2017, 1, 1));
					if (HTKKFormater.isNullString(text2) || text2.Equals("1") || text2.Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuBL, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_KyHieuBL);
					}
				}
				bool flag = false;
				if (array2[4].ToIntOrDefault() == 0 && array2[5].ToIntOrDefault() == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					num = 2;
					flag = true;
				}
				else if (array2[4].ToIntOrDefault() == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_DenSo);
					num = 2;
					flag = true;
				}
				else if (array2[5].ToIntOrDefault() == 0)
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TuSo);
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					num = 2;
					flag = true;
				}
				else
				{
					double num3 = ConvertHelpers.ConvertToDouble(array2[4]);
					double num4 = ConvertHelpers.ConvertToDouble(array2[5]);
					if (num3 > num4)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16")));
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16")));
						num = 2;
						flag = true;
					}
				}
				if (array2[4].ToIntOrDefault(-1) == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_56")));
					num = 2;
				}
				if (array2[5].ToIntOrDefault(-1) == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_56")));
					num = 2;
				}
				if (flag || array == null || array.Length <= 1)
				{
					continue;
				}
				for (int j = 0; j < array.Length; j++)
				{
					if (num2 + i == ConvertHelpers.ConvertToInt(array[j]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14")));
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14")));
						num = 2;
					}
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNguoiKy));
			if (hTKKCMCCellInfo != null)
			{
				string str = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
				if (HTKKFormater.isNullString(str))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNguoiLapBieu));
			if (hTKKCMCCellInfo2 != null)
			{
				string str2 = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
				if (HTKKFormater.isNullString(str2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
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

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
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

	public void ExportXML_0(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
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
}
