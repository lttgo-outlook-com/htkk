using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Utilities.Implement;

public class HTKKTraCuuImpl : HTKKBase<HTKKTraCuuInfo>
{
	public HTKKTraCuuInfo GetInfo(XmlNode xmlNode)
	{
		//IL_0245: Unknown result type (might be due to invalid IL or missing references)
		//IL_024b: Expected O, but got Unknown
		try
		{
			HTKKTraCuuInfo hTKKTraCuuInfo = new HTKKTraCuuInfo();
			hTKKTraCuuInfo.ID = ((xmlNode.Attributes["ID"] == null) ? "" : ((XmlNode)xmlNode.Attributes["ID"]).Value);
			hTKKTraCuuInfo.FinanceYear = ((xmlNode.Attributes["FinanceYear"] == null) ? "" : ((XmlNode)xmlNode.Attributes["FinanceYear"]).Value);
			hTKKTraCuuInfo.Year = ((xmlNode.Attributes["Year"] == null) ? "" : ((XmlNode)xmlNode.Attributes["Year"]).Value);
			hTKKTraCuuInfo.Month = ((xmlNode.Attributes["Month"] == null) ? "" : ((XmlNode)xmlNode.Attributes["Month"]).Value);
			hTKKTraCuuInfo.ThreeMonth = ((xmlNode.Attributes["ThreeMonth"] == null) ? "" : ((XmlNode)xmlNode.Attributes["ID"]).Value);
			hTKKTraCuuInfo.Day = ((xmlNode.Attributes["Day"] == null) ? "" : ((XmlNode)xmlNode.Attributes["Day"]).Value);
			hTKKTraCuuInfo.Caption = ((xmlNode.Attributes["Caption"] == null) ? "" : ((XmlNode)xmlNode.Attributes["Caption"]).Value);
			hTKKTraCuuInfo.LoaiTK = ((xmlNode.Attributes["LoaiTK"] == null) ? "" : ((XmlNode)xmlNode.Attributes["LoaiTK"]).Value);
			hTKKTraCuuInfo.Prefix = ((xmlNode.Attributes["Prefix"] == null) ? "" : ((XmlNode)xmlNode.Attributes["Prefix"]).Value);
			hTKKTraCuuInfo.TkThangQuy = ((xmlNode.Attributes["TkThangQuy"] == null) ? "" : ((XmlNode)xmlNode.Attributes["TkThangQuy"]).Value);
			XmlNodeList val = xmlNode.SelectNodes("Validity");
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				HTKKTraCuuValidity hTKKTraCuuValidity = new HTKKTraCuuValidity();
				hTKKTraCuuValidity.StartDate = ((val2.Attributes["StartDate"] == null) ? "" : ((XmlNode)val2.Attributes["StartDate"]).Value);
				hTKKTraCuuValidity.EndDate = ((val2.Attributes["EndDate"] == null) ? "" : ((XmlNode)val2.Attributes["EndDate"]).Value);
				hTKKTraCuuValidity.ThueKhauTru = ((val2.Attributes["ThueKhauTru"] == null) ? "" : ((XmlNode)val2.Attributes["ThueKhauTru"]).Value);
				hTKKTraCuuValidity.ThuePhaiNop = ((val2.Attributes["ThuePhaiNop"] == null) ? "" : ((XmlNode)val2.Attributes["ThuePhaiNop"]).Value);
				string text = ((val2.Attributes["DataFile"] == null) ? "" : ((XmlNode)val2.Attributes["DataFile"]).Value);
				if (text.IndexOf(",") > 0)
				{
					hTKKTraCuuValidity.DataFile = text.Split(new char[1] { ',' })[0];
				}
				else if (text.IndexOf(";") > 0)
				{
					hTKKTraCuuValidity.DataFile = text.Split(new char[1] { ';' })[0];
				}
				else
				{
					hTKKTraCuuValidity.DataFile = text;
				}
				if (HTKKFormater.isNullString(hTKKTraCuuValidity.DataFile))
				{
					return null;
				}
				hTKKTraCuuInfo.Validity.Add(hTKKTraCuuValidity);
			}
			return hTKKTraCuuInfo;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDicSacThue(XmlNodeList xnList)
	{
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_003f: Expected O, but got Unknown
		string text = "";
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			if (xnList.Count > 0)
			{
				foreach (XmlNode xn in xnList)
				{
					XmlNode val = xn;
					string text2 = ((val.Attributes["ID"] == null) ? "" : ((XmlNode)val.Attributes["ID"]).Value);
					string str = ((val.Attributes["PopID"] == null) ? "" : ((XmlNode)val.Attributes["PopID"]).Value);
					string text3 = ((val.Attributes["ParentID"] == null) ? "" : ((XmlNode)val.Attributes["ParentID"]).Value);
					string text4 = ((val.Attributes["Caption"] == null) ? "" : ((XmlNode)val.Attributes["Caption"]).Value);
					string value = text4;
					if (text2 == "101_6")
					{
						text = text2;
					}
					if (HTKKFormater.isNullString(str))
					{
						if (!HTKKFormater.isNullString(text3) && text3.IndexOf("_") == -1 && !HTKKFormater.isNullString(text2))
						{
							dictionary.Add(text2, value);
						}
					}
					else
					{
						if (HTKKFormater.isNullString(text3) || text3.IndexOf("_") <= -1 || HTKKFormater.isNullString(text2) || dictionary.Count <= 0)
						{
							continue;
						}
						foreach (KeyValuePair<string, string> item in dictionary)
						{
							string key = item.Key;
							value = item.Value;
							if (key.Equals(text3))
							{
								value = value + "$" + text2;
								dictionary.Remove(key);
								dictionary.Add(key, value);
								break;
							}
						}
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, HTKKTraCuuInfo> GetDictDataByMapFileXml(string pathFile)
	{
		//IL_001b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0021: Expected O, but got Unknown
		//IL_004c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0052: Expected O, but got Unknown
		try
		{
			Dictionary<string, HTKKTraCuuInfo> dictionary = new Dictionary<string, HTKKTraCuuInfo>();
			if (File.Exists(pathFile))
			{
				XmlDocument val = new XmlDocument();
				val.LoadXml(HTKKFile.LoadFile(pathFile));
				XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Root/Map");
				foreach (XmlNode item in val2)
				{
					XmlNode xmlNode = item;
					HTKKTraCuuInfo info = GetInfo(xmlNode);
					if (info != null)
					{
						dictionary.Add(info.ID, info);
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public DataTable GetDataTable(List<HTKKTraCuuInfo> listDataReturn)
	{
		try
		{
			return HTKKBase<HTKKTraCuuInfo>.ToDataTable(listDataReturn);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public List<HTKKTraCuuInfo> GetListDataSelect(DataTable tabSelect, Dictionary<string, HTKKTraCuuInfo> dictData)
	{
		try
		{
			List<HTKKTraCuuInfo> list = new List<HTKKTraCuuInfo>();
			List<string> list2 = new List<string>();
			foreach (KeyValuePair<string, HTKKTraCuuInfo> dictDatum in dictData)
			{
				HTKKTraCuuInfo hTKKTraCuuInfo = new HTKKTraCuuInfo();
				hTKKTraCuuInfo = dictDatum.Value;
				List<HTKKTraCuuValidity> validity = hTKKTraCuuInfo.Validity;
				if (validity.Count <= 0)
				{
					continue;
				}
				for (int i = 0; i < validity.Count; i++)
				{
					string dataFile = validity[i].DataFile;
					DataRow[] array = tabSelect.Select("FileName LIKE '%" + dataFile + "%'");
					DataRow[] array2 = array;
					foreach (DataRow dataRow in array2)
					{
						if (list2.Contains(dataRow["FileName"].ToString()))
						{
							continue;
						}
						string text = dataRow["FileName"].ToString();
						string versionXML = HTKKFormater.GetVersionXML(dataRow["FileName"].ToString());
						string text2 = dataFile;
						string empty = string.Empty;
						string text3 = string.Empty;
						string kyKeKhai = string.Empty;
						string text4 = string.Empty;
						string text5 = string.Empty;
						string text6 = string.Empty;
						string text7 = string.Empty;
						string text8 = string.Empty;
						string text9 = string.Empty;
						string text10 = string.Empty;
						string text11 = string.Empty;
						string empty2 = string.Empty;
						string text12 = string.Empty;
						string empty3 = string.Empty;
						string empty4 = string.Empty;
						string empty5 = string.Empty;
						string empty6 = string.Empty;
						string text13 = string.Empty;
						string cqtCapTion = string.Empty;
						string text14 = string.Empty;
						string empty7 = string.Empty;
						string empty8 = string.Empty;
						string text15 = string.Empty;
						string empty9 = string.Empty;
						string text16 = string.Empty;
						string text17 = string.Empty;
						bool flag = false;
						string text18 = string.Empty;
						if (HTKKFormater.isNullString(hTKKTraCuuInfo.Prefix) || hTKKTraCuuInfo.Prefix.Equals("MADADT"))
						{
						}
						if (text.IndexOf("bs") == 0)
						{
							text12 = text[..text.IndexOf("_")];
							text12 = text12.Replace("bs", "");
							empty4 = "B";
							empty5 = "B";
							text13 = ((!(hTKKTraCuuInfo.ID == "G4") && !(hTKKTraCuuInfo.ID == "N5") && !(hTKKTraCuuInfo.ID == "H5") && !(hTKKTraCuuInfo.ID == "O7") && !(hTKKTraCuuInfo.ID == "P5") && !(hTKKTraCuuInfo.ID == "U6")) ? (" - Bổ sung lần thứ " + text12) : " - Thay thế");
						}
						else if (text.IndexOf("tt") == 0)
						{
							text12 = text[..text.IndexOf("_")];
							text12 = text12.Replace("tt", "");
							empty4 = "T";
							empty5 = "T";
							text13 = "Thay thế lần thứ " + text12;
						}
						else
						{
							empty4 = "C";
							empty5 = "C";
						}
						string text19 = "";
						if ("~K2~K3~J7~E6~E7~D7~R6~X5~R9~E9~".IndexOf("~" + hTKKTraCuuInfo.ID.Substring(0, 2) + "~") >= 0)
						{
							text19 = text.Replace(text2 + "_", "").Replace("_" + versionXML + ".xml", "");
							if (empty4.Equals("B"))
							{
								text19 = text19.Replace("bs" + text12 + "_", "");
							}
							if (text19.Substring(0, 3) == "USD" || text19.Substring(0, 3) == "VND")
							{
								empty9 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ma_DonVi']");
								text16 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ten_DonVi']");
								hTKKTraCuuInfo.DonViTinhMa = empty9;
								hTKKTraCuuInfo.DonViTinhTen = text16;
								if (string.IsNullOrEmpty(hTKKTraCuuInfo.DonViTinhMa))
								{
									if (text19.Substring(0, 3) == "USD")
									{
										hTKKTraCuuInfo.DonViTinhMa = "USD";
										hTKKTraCuuInfo.DonViTinhTen = "Đô la Mỹ";
									}
									if (text19.Substring(0, 3) == "VND")
									{
										hTKKTraCuuInfo.DonViTinhMa = "VND";
										hTKKTraCuuInfo.DonViTinhTen = "Đồng Việt Nam";
									}
									empty9 = hTKKTraCuuInfo.DonViTinhMa;
									text16 = hTKKTraCuuInfo.DonViTinhTen;
								}
								text19 = text19.Substring(4);
								flag = true;
							}
						}
						else
						{
							text19 = text.Replace(text2 + "_", "").Replace("_" + versionXML + ".xml", "").Replace("_VND", "")
								.Replace("_NT", "");
						}
						hTKKTraCuuInfo = RevertNameNopTheoTinhCQT(hTKKTraCuuInfo, text, ref text19, ref cqtCapTion);
						if (text19.Contains("CN"))
						{
							string[] array3 = text19.Split(new char[1] { '_' });
							string[] array4 = array3;
							foreach (string text20 in array4)
							{
								if (text20.Contains("CN"))
								{
									string tieuMucHachToan = text20.Replace("CN", "");
									hTKKTraCuuInfo.TieuMucHachToan = tieuMucHachToan;
								}
							}
							text19 = text19.Replace("CN" + hTKKTraCuuInfo.TieuMucHachToan + "_", "");
						}
						if (hTKKTraCuuInfo.ID.Substring(0, 2) == "D1" || hTKKTraCuuInfo.ID.Substring(0, 2) == "Z5")
						{
							string[] array5 = text19.Split(new char[1] { '_' });
							if (array5 != null && array5.Length > 1)
							{
								hTKKTraCuuInfo.Prefix = array5[^2];
							}
						}
						if (hTKKTraCuuInfo.ID.Substring(0, 2) == "R5")
						{
							string[] array5 = text19.Split(new char[1] { '_' });
							if (array5 != null && array5.Length > 1)
							{
								hTKKTraCuuInfo.Prefix = array5[^2];
							}
						}
						if (!hTKKTraCuuInfo.Prefix.Equals("") && (hTKKTraCuuInfo.ID.Substring(0, 2) == "71" || hTKKTraCuuInfo.ID.Substring(0, 2) == "D4" || hTKKTraCuuInfo.ID.Substring(0, 2) == "01" || hTKKTraCuuInfo.ID.Substring(0, 2) == "D1" || hTKKTraCuuInfo.ID.Substring(0, 2) == "02" || hTKKTraCuuInfo.ID.Substring(0, 2) == "A4" || hTKKTraCuuInfo.ID.Substring(0, 2) == "D2" || hTKKTraCuuInfo.ID.Substring(0, 2) == "E8" || hTKKTraCuuInfo.ID.Substring(0, 2) == "I5" || hTKKTraCuuInfo.ID.Substring(0, 2) == "P6" || hTKKTraCuuInfo.ID.Substring(0, 2) == "P1" || hTKKTraCuuInfo.ID.Substring(0, 2) == "R5" || hTKKTraCuuInfo.ID.Substring(0, 2) == "X7" || hTKKTraCuuInfo.ID.Substring(0, 2) == "Y9" || hTKKTraCuuInfo.ID.Substring(0, 2) == "Z3" || hTKKTraCuuInfo.ID.Substring(0, 2) == "Z5"))
						{
							text19 = text19.Replace(hTKKTraCuuInfo.Prefix + "_", "");
						}
						string text21 = "";
						if (hTKKTraCuuInfo.ID.Substring(0, 2) == "02" || hTKKTraCuuInfo.ID.Substring(0, 2) == "A4" || hTKKTraCuuInfo.ID.Substring(0, 2) == "99")
						{
							text21 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/maHoSoKhaiDuAnDauTu']");
						}
						else if (hTKKTraCuuInfo.ID.Substring(0, 2) == "D2")
						{
							text21 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct01a']");
						}
						else if (hTKKTraCuuInfo.ID.Substring(0, 2) == "P6" || hTKKTraCuuInfo.ID.Substring(0, 2) == "P1" || hTKKTraCuuInfo.ID.Substring(0, 2) == "X7" || hTKKTraCuuInfo.ID.Substring(0, 2) == "Y9")
						{
							text21 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/soChungTu']");
						}
						else if ("~O1~O2~O3~O4~O5~".IndexOf("~" + hTKKTraCuuInfo.ID + "~") >= 0)
						{
							text21 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/maSoThue']");
						}
						if (hTKKTraCuuInfo.Prefix.Equals("MADADT"))
						{
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "A4" || hTKKTraCuuInfo.ID.Substring(0, 2) == "99")
							{
								string[] array6 = text19.Split(new char[1] { '_' });
								if (array6.Length == 3)
								{
									text19 = array6[1] + "_" + array6[2];
								}
							}
							else if (hTKKTraCuuInfo.ID.Substring(0, 2) == "P6" || hTKKTraCuuInfo.ID.Substring(0, 2) == "P1" || hTKKTraCuuInfo.ID.Substring(0, 2) == "X7" || hTKKTraCuuInfo.ID.Substring(0, 2) == "Y9")
							{
								if (!text21.Equals(""))
								{
									text19 = text19.Replace("_" + text21 + "_", "_");
								}
							}
							else
							{
								string[] array6 = text19.Split(new char[1] { '_' });
								if (array6.Length == 2 || array6.Length == 3)
								{
									text19 = text19.Replace(array6[0] + "_", "");
								}
								if (!text21.Equals(""))
								{
									text19 = text19.Replace(text21 + "_", "");
								}
							}
						}
						string text22 = "";
						if (empty4 == "B")
						{
							text19 = text19.Replace("bs" + text12 + "_", "");
						}
						string text23 = "";
						text23 = text.Replace(text2 + "_", "").Replace("bs" + text12 + "_", "");
						if ("~K2~K3~J7~E6~E7~D7~R6~X5~R9~E9~".IndexOf("~" + hTKKTraCuuInfo.ID.Substring(0, 2) + "~") >= 0 && (text23.Substring(0, 3) == "USD" || text23.Substring(0, 3) == "VND"))
						{
							text23 = text23.Substring(4);
						}
						if (empty4 == "T")
						{
							text19 = text19.Replace("tt" + text12 + "_", "");
							text23 = text.Replace(text2 + "_", "").Replace("tt" + text12 + "_", "");
						}
						if (!string.IsNullOrEmpty(hTKKTraCuuInfo.NopTheoTinhCQT))
						{
							string text20 = "";
							if (text.Contains("TINH") && !string.IsNullOrEmpty(hTKKTraCuuInfo.maTinh))
							{
								text20 = "TINH" + hTKKTraCuuInfo.maTinh;
							}
							if (text.Contains("CQTQL") && !string.IsNullOrEmpty(hTKKTraCuuInfo.maCQTQuanLy))
							{
								text20 = "CQTQL" + hTKKTraCuuInfo.maCQTQuanLy;
							}
							if (!string.IsNullOrEmpty(text20))
							{
								text23 = text23.Replace(text20 + "_", "");
							}
						}
						if (hTKKTraCuuInfo.ID.Substring(0, 2) == "T6" || hTKKTraCuuInfo.ID.Substring(0, 2) == "U2")
						{
							text23 = text23.Replace("CN" + hTKKTraCuuInfo.TieuMucHachToan + "_", "");
						}
						if (hTKKTraCuuInfo.ID.Substring(0, 2) == "Y5")
						{
							text23 = text23.Replace("CN" + hTKKTraCuuInfo.TieuMucHachToan + "_", "");
							string[] array7 = text23.Split(new char[1] { '_' });
							text19 = array7[1];
							text4 = array7[0].Substring(0, 2) + "/" + array7[0].Substring(2, 2) + "/" + array7[0].Substring(4);
							text13 = text13 + " - " + ((hTKKTraCuuInfo.TieuMucHachToan == "1") ? "Bên nhận chuyển nhượng vốn" : "Bên doanh nghiệp có vốn được chuyển nhượng");
							string text24 = text13;
							text13 = text24 + " - Ngày hợp đồng " + array7[0].Substring(0, 2) + "/" + array7[0].Substring(2, 2) + "/" + array7[0].Substring(4);
						}
						if (hTKKTraCuuInfo.ID.Substring(0, 2) == "U1")
						{
							string[] array7 = text23.Split(new char[1] { '_' });
							hTKKTraCuuInfo.TieuMucHachToan = array7[0].Replace("NN", "");
							hTKKTraCuuInfo.ToChucCNDNVonChuyenNhuong = Convert.ToInt32(array7[1].Replace("L", ""));
							text23 = text23.Replace(array7[0] + "_" + array7[1] + "_", "");
							text19 = array7[2];
							text13 = text13 + " - " + ((hTKKTraCuuInfo.TieuMucHachToan == "1") ? "Hoạt động chuyển nhượng bất động sản" : ((hTKKTraCuuInfo.ToChucCNDNVonChuyenNhuong == 1) ? "Bán đấu giá - Bất động sản" : "Bán đấu giá - Tài sản khác"));
						}
						if (hTKKTraCuuInfo.ID.Substring(0, 2) == "Y3")
						{
							string[] array7 = text23.Split(new char[1] { '_' });
							hTKKTraCuuInfo.TieuMucHachToan = array7[0].Replace("NN", "");
							hTKKTraCuuInfo.ToChucCNDNVonChuyenNhuong = Convert.ToInt32(array7[1].Replace("L", ""));
							text23 = text23.Replace(array7[0] + "_" + array7[1] + "_", "");
							text19 = array7[2];
							text13 = text13 + " - " + ((hTKKTraCuuInfo.TieuMucHachToan == "2") ? "Tài sản khác" : ((hTKKTraCuuInfo.ToChucCNDNVonChuyenNhuong == 1) ? "Bất động sản - Tổ chức áp dụng phương pháp tỷ lệ trên doanh thu theo quy định của pháp luật về thuế TNDN có hoạt động chuyển nhượng bất động sản" : "Bất động sản - Cơ quan, tổ chức bán hoặc bán đấu giá bất động sản, tài sản khác là tài sản bảo đảm thi hành án, xử lý nợ"));
						}
						if (text2.Contains("NN") && "N7~M1~M2~C6~C7~T5~H4~O1~O2~O3~O4~O5~O6~Q4~P7~U1~U8~X5~E9~R9~Y3~Y4~Z1~".Contains(hTKKTraCuuInfo.ID.Substring(0, 2)) && !string.IsNullOrEmpty(GetCatalogueDMNNKDItem(text2)))
						{
							string[] array7 = text23.Split(new char[1] { '_' });
							hTKKTraCuuInfo.TieuMucHachToan = array7[0].Replace("NN", "");
							text23 = text23.Replace(array7[0] + "_", "");
							text19 = array7[1];
							string catalogueDMNNKDItem = GetCatalogueDMNNKDItem(text2 + "###" + hTKKTraCuuInfo.TieuMucHachToan);
							if (!string.IsNullOrEmpty(catalogueDMNNKDItem))
							{
								string[] array8 = catalogueDMNNKDItem.Split(new string[1] { "###" }, StringSplitOptions.None);
								if (array8.Length > 2)
								{
									hTKKTraCuuInfo.NganhNgheKD = array8[2];
									text13 += array8[2];
								}
							}
						}
						if (!hTKKTraCuuInfo.Prefix.Equals("") && (hTKKTraCuuInfo.ID.Substring(0, 2) == "71" || hTKKTraCuuInfo.ID.Substring(0, 2) == "D4" || hTKKTraCuuInfo.ID.Substring(0, 2) == "01" || hTKKTraCuuInfo.ID.Substring(0, 2) == "D1" || hTKKTraCuuInfo.ID.Substring(0, 2) == "R5" || hTKKTraCuuInfo.ID.Substring(0, 2) == "02" || hTKKTraCuuInfo.ID.Substring(0, 2) == "A4" || hTKKTraCuuInfo.ID.Substring(0, 2) == "99" || hTKKTraCuuInfo.ID.Substring(0, 2) == "H6" || hTKKTraCuuInfo.ID.Substring(0, 2) == "D2" || hTKKTraCuuInfo.ID.Substring(0, 2) == "E8" || hTKKTraCuuInfo.ID.Substring(0, 2) == "I5" || hTKKTraCuuInfo.ID.Substring(0, 2) == "Z3" || hTKKTraCuuInfo.ID.Substring(0, 2) == "Z5"))
						{
							text23 = text23.Replace(hTKKTraCuuInfo.Prefix + "_", "");
						}
						if (hTKKTraCuuInfo.Prefix.Equals("MADADT"))
						{
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "A4" || hTKKTraCuuInfo.ID.Substring(0, 2) == "99" || hTKKTraCuuInfo.ID.Substring(0, 2) == "H6")
							{
								string[] array9 = text23.Split(new char[1] { '_' });
								if (array9.Length == 4)
								{
									text23 = array9[1] + "_" + array9[2] + "_" + array9[3];
								}
							}
							else
							{
								string[] array9 = text23.Split(new char[1] { '_' });
								if (array9.Length == 2 || array9.Length == 3)
								{
									text23 = text19.Replace(array9[0] + "_", "");
								}
							}
						}
						if (!(hTKKTraCuuInfo.ID.Substring(0, 2) == "99") && !(hTKKTraCuuInfo.ID.Substring(0, 2) == "H6") && !(hTKKTraCuuInfo.ID == "N7") && !text23.StartsWith("DT") && !text23.StartsWith("L") && !text23.StartsWith("KTN") && !text23.StartsWith("CD") && !text23.StartsWith("KDDT") && !text23.StartsWith("KDCD") && !text23.StartsWith("KDKTN") && !text23.StartsWith("Q") && !text23.StartsWith("K") && text23.Length > 2 && ConvertHelpers.ConvertToInt(text23.Substring(0, 2)) == 0)
						{
							continue;
						}
						if ("66;67;09;10;60;61;G2;".IndexOf(hTKKTraCuuInfo.ID) > -1)
						{
							text19 = "";
							text22 = "Y";
						}
						else
						{
							string text25 = "";
							text25 = ((hTKKTraCuuInfo.ID == null || hTKKTraCuuInfo.ID.Length <= 2) ? hTKKTraCuuInfo.ID : hTKKTraCuuInfo.ID.Substring(0, 2));
							if ("~98~92~08~40~42~44~M8~M9~K2~K3~J9~J8~J7~E2~E3~E4~E5~E6~E7~E9~D7~D8~D9~R9~R6~S1~".IndexOf("~" + text25 + "~") > 0)
							{
								string[] array10 = text19.Split(new char[1] { '_' });
								if (array10 != null && array10.Length > 0)
								{
									if (array10[0].Equals("DT"))
									{
										hTKKTraCuuInfo.TieuMucHachToan = "1";
										if (hTKKTraCuuInfo.Prefix.Equals(""))
										{
											text18 = " - Dầu thô ";
										}
									}
									else if (array10[0].Equals("KTN"))
									{
										hTKKTraCuuInfo.TieuMucHachToan = "2";
										if (hTKKTraCuuInfo.Prefix.Equals(""))
										{
											text18 = " - Khí thiên nhiên ";
										}
									}
									else if (array10[0].Equals("CD"))
									{
										hTKKTraCuuInfo.TieuMucHachToan = "3";
										if (hTKKTraCuuInfo.Prefix.Equals(""))
										{
											text18 = " - Condensate ";
										}
									}
									else if (array10[0].Equals("KDDT"))
									{
										hTKKTraCuuInfo.TieuMucHachToan = "4";
										if (hTKKTraCuuInfo.Prefix.Equals(""))
										{
											text18 = " - Kết dư dầu ";
										}
									}
									else if (array10[0].Equals("KDKTN"))
									{
										hTKKTraCuuInfo.TieuMucHachToan = "5";
										if (hTKKTraCuuInfo.Prefix.Equals(""))
										{
											text18 = " - Kết dư khí ";
										}
									}
									else if (array10[0].Equals("KDCD"))
									{
										hTKKTraCuuInfo.TieuMucHachToan = "6";
										if (hTKKTraCuuInfo.Prefix.Equals(""))
										{
											text18 = " - Kết dư Condensate ";
										}
									}
								}
								string[] array11 = new string[6] { "DT_", "KTN_", "CD_", "KDDT_", "KDKTN_", "KDCD_" };
								string[] array12 = text19.Split(new char[1] { '_' });
								if (!hTKKTraCuuInfo.Prefix.Equals(array12[0]) && !hTKKTraCuuInfo.Prefix.Equals(""))
								{
									continue;
								}
								if (!hTKKTraCuuInfo.Prefix.Equals(""))
								{
									if (hTKKTraCuuInfo.ID.Substring(0, 2) == "R6")
									{
										string[] array7 = text23.Split(new char[1] { '_' });
										hTKKTraCuuInfo.ToChucCNDNVonChuyenNhuong = Convert.ToInt32(array7[1].Replace("LHD", ""));
										text23 = text23.Replace(array7[0] + "_" + array7[1] + "_", "");
										text19 = array7[2];
										text18 = text18 + " - " + ((hTKKTraCuuInfo.ToChucCNDNVonChuyenNhuong == 1) ? "Hợp đồng dầu khí (PSC)" : "Hiệp đinh liên chính phủ (VSP)");
									}
									else
									{
										text19 = text19.Replace(hTKKTraCuuInfo.Prefix + "_", "");
									}
								}
								else
								{
									for (int l = 0; l < array11.Length; l++)
									{
										text19 = text19.Replace(array11[l], "");
									}
								}
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2).Equals("D1") || hTKKTraCuuInfo.ID.Substring(0, 2) == "Z5")
							{
								hTKKTraCuuInfo.maXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/DiaChiHDSXKDKhacTinhNDTSC/ct11a_phuongXa_ma']");
								hTKKTraCuuInfo.tenXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/DiaChiHDSXKDKhacTinhNDTSC/ct11a_phuongXa_ten']");
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2).Equals("D4") || hTKKTraCuuInfo.ID.Substring(0, 2) == "Z3")
							{
								hTKKTraCuuInfo.maXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct11a_ma']");
								hTKKTraCuuInfo.tenXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct11a_ten']");
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2).Equals("Q9"))
							{
								hTKKTraCuuInfo.maXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='maxapl2']");
								hTKKTraCuuInfo.tenXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='tenxapl2']");
								hTKKTraCuuInfo.maTinh = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='matinhpl2']");
								hTKKTraCuuInfo.tenTinh = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='tentinhpl2']");
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2).Equals("D5") || hTKKTraCuuInfo.ID.Substring(0, 2).Equals("Z4"))
							{
								hTKKTraCuuInfo.maXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct09a_ma']");
								hTKKTraCuuInfo.tenXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct09a_ten']");
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2).Equals("H2"))
							{
								hTKKTraCuuInfo.maXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT09/ct09a_phuongXa_ma']");
								hTKKTraCuuInfo.tenXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT09/ct09a_phuongXa_ten']");
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2).Equals("H4"))
							{
								hTKKTraCuuInfo.maXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT09/ct09a_phuongXa_ma']");
								hTKKTraCuuInfo.tenXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT09/ct09a_phuongXa_ten']");
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2).Equals("G9"))
							{
								hTKKTraCuuInfo.maXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/DiaChiHDSXKD/ct11b_ma']");
								hTKKTraCuuInfo.tenXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/DiaChiHDSXKD/ct11b_ten']");
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2).Equals("T8") || hTKKTraCuuInfo.ID.Substring(0, 2).Equals("R5"))
							{
								hTKKTraCuuInfo.maXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/DiaChiHDSXKDKhacTinhNDTSC/ct11a_phuongXa_ma']");
								hTKKTraCuuInfo.tenXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/DiaChiHDSXKDKhacTinhNDTSC/ct11a_phuongXa_ten']");
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2).Equals("99"))
							{
								hTKKTraCuuInfo.maXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/maXa']");
								hTKKTraCuuInfo.tenXa = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/tenXa']");
								hTKKTraCuuInfo.NganhNgheKD = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/tenLoaiHoan']");
								hTKKTraCuuInfo.TieuMucHachToan = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/maLoaiHoan']");
							}
							if (hTKKTraCuuInfo.ID == "03" || hTKKTraCuuInfo.ID == "79" || hTKKTraCuuInfo.ID == "791" || hTKKTraCuuInfo.ID == "K8" || hTKKTraCuuInfo.ID == "K1" || hTKKTraCuuInfo.ID == "K11")
							{
								hTKKTraCuuInfo.TieuMucHachToan = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/tieuMucHachToan']");
								hTKKTraCuuInfo.NganhNgheKD = "";
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "T8" || hTKKTraCuuInfo.ID.Substring(0, 2) == "R5")
							{
								hTKKTraCuuInfo.TieuMucHachToan = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ma_NganhNghe']");
								hTKKTraCuuInfo.NganhNgheKD = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ten_NganhNghe']");
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "01" || hTKKTraCuuInfo.ID.Substring(0, 2) == "I5")
							{
								hTKKTraCuuInfo.TieuMucHachToan = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ma_NganhNghe']");
								hTKKTraCuuInfo.NganhNgheKD = "";
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "E8")
							{
								hTKKTraCuuInfo.TieuMucHachToan = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/apDung_ma']");
								hTKKTraCuuInfo.NganhNgheKD = "";
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "D1" || hTKKTraCuuInfo.ID.Substring(0, 2) == "Z5")
							{
								hTKKTraCuuInfo.TieuMucHachToan = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ma_NganhNghe']");
								text15 = " - " + HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ten_NganhNghe']");
								hTKKTraCuuInfo.NganhNgheKD = "";
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "R5")
							{
								hTKKTraCuuInfo.TieuMucHachToan = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ma_NganhNghe']");
								text15 = (hTKKTraCuuInfo.NganhNgheKD = " - " + HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ten_NganhNghe']"));
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "02")
							{
								hTKKTraCuuInfo.MaHSDuAnDT = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/maHoSoKhaiDuAnDauTu']");
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "D2")
							{
								hTKKTraCuuInfo.MaHSDuAnDT = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct01a']");
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "P6" || hTKKTraCuuInfo.ID.Substring(0, 2) == "P1" || hTKKTraCuuInfo.ID.Substring(0, 2) == "X7" || hTKKTraCuuInfo.ID.Substring(0, 2) == "Y9")
							{
								hTKKTraCuuInfo.MaHSDuAnDT = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/soChungTu']");
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "A4" || hTKKTraCuuInfo.ID.Substring(0, 2) == "99")
							{
								hTKKTraCuuInfo.MaHSDuAnDT = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/maHoSoKhaiDuAnDauTu']");
							}
							if ("~O1~O2~O3~O4~O5~".IndexOf("~" + hTKKTraCuuInfo.ID + "~") >= 0)
							{
								hTKKTraCuuInfo.MaHSDuAnDT = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/maSoThue']");
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "71")
							{
								string valueByAttributes = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/hoatDongThuHo']");
								if (valueByAttributes.ToUpper().Equals("TRUE"))
								{
									hTKKTraCuuInfo.ThuHo = "TH";
								}
								else
								{
									hTKKTraCuuInfo.ThuHo = "";
								}
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "D4" || hTKKTraCuuInfo.ID.Substring(0, 2) == "Z3")
							{
								string valueByAttributes = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/khaiRiengThuHo']");
								if (valueByAttributes.ToUpper().Equals("TRUE"))
								{
									hTKKTraCuuInfo.ThuHo = "TH";
								}
								else
								{
									hTKKTraCuuInfo.ThuHo = "";
								}
							}
							if (hTKKTraCuuInfo.ID == "47" || hTKKTraCuuInfo.ID == "K7")
							{
								if (text.IndexOf("_CTDC") == -1)
								{
									continue;
								}
								HTKKFile.AddOrUpdateDicStr(hTKKTraCuuInfo.dicVar, "A2_CTDC", "True");
								text19 = text19.Replace("_CTDC", "").Replace("_LNDC", "");
							}
							if (hTKKTraCuuInfo.ID == "471" || hTKKTraCuuInfo.ID == "K71")
							{
								if (text.IndexOf("_LNDC") == -1)
								{
									continue;
								}
								HTKKFile.AddOrUpdateDicStr(hTKKTraCuuInfo.dicVar, "A2_CTDC", "False");
								text19 = text19.Replace("_CTDC", "").Replace("_LNDC", "");
							}
							if (hTKKTraCuuInfo.ID == "L9" || hTKKTraCuuInfo.ID == "I1" || hTKKTraCuuInfo.ID == "I2")
							{
								string[] array9 = text23.Split(new char[1] { '_' });
								if (array9.Length == 3)
								{
									text23 = array9[1] + "_" + array9[2];
									text19 = text19.Replace(array9[0] + "_", "");
									empty3 = array9[0].Replace("L", "");
									text14 = " - lần " + empty3;
									hTKKTraCuuInfo.LanXuatBan = array9[0].Replace("L", "");
								}
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "M4" || hTKKTraCuuInfo.ID.Substring(0, 2) == "G3")
							{
								string[] array9 = text23.Split(new char[1] { '_' });
								if (array9.Length == 3)
								{
									text23 = array9[1] + "_" + array9[2];
									text19 = text19.Replace(array9[0] + "_", "");
									empty3 = array9[0].Replace("L", "");
									text14 = " - lần " + empty3;
									hTKKTraCuuInfo.LanXuatBan = array9[0].Replace("L", "");
									hTKKTraCuuInfo.Lan = hTKKTraCuuInfo.LanXuatBan;
								}
							}
							if ("U8~".Contains(hTKKTraCuuInfo.ID.Substring(0, 2)))
							{
								string[] array9 = text23.Split(new char[1] { '_' });
								if (array9.Length == 3)
								{
									text23 = array9[1] + "_" + array9[2];
									text19 = text19.Replace(array9[0] + "_", "");
									empty3 = array9[0].Replace("L", "");
									text14 = ((hTKKTraCuuInfo.ID.Substring(0, 2) == "U8") ? " - File số " : " - lần ") + empty3;
									hTKKTraCuuInfo.LanXuatBan = array9[0].Replace("L", "");
								}
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "99" || hTKKTraCuuInfo.ID.Substring(0, 2) == "H6")
							{
								if (text19.IndexOf("TT") == 0)
								{
									text19 = text19.Replace("TT", "");
									hTKKTraCuuInfo.KieuKyHoanThue = "1";
									hTKKTraCuuInfo.LoaiTK = "M";
									text22 = "M";
								}
								if (text19.IndexOf("QQ") == 0)
								{
									text19 = text19.Replace("QQ", "");
									hTKKTraCuuInfo.KieuKyHoanThue = "2";
									hTKKTraCuuInfo.LoaiTK = "Q";
									text22 = "Q";
								}
								if (text19.IndexOf("TQ") == 0)
								{
									text19 = text19.Replace("TQ", "");
									hTKKTraCuuInfo.KieuKyHoanThue = "3";
									hTKKTraCuuInfo.LoaiTK = "Q";
									text22 = "Q";
								}
								if (text19.IndexOf("QT") == 0)
								{
									text19 = text19.Replace("QT", "");
									hTKKTraCuuInfo.KieuKyHoanThue = "4";
									hTKKTraCuuInfo.LoaiTK = "M";
									text22 = "M";
								}
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "R9")
							{
								string[] array7 = text19.Split(new char[1] { '_' });
								if (array7.Length >= 2)
								{
									text19 = array7[1];
									hTKKTraCuuInfo.LanXuatBan = array7[0].Replace("L", "");
								}
							}
							if (!(hTKKTraCuuInfo.ID.Substring(0, 2) == "99") && !(hTKKTraCuuInfo.ID.Substring(0, 2) == "H6"))
							{
								text22 = HTKKFormater.GetKieuKyByLengh(text19, hTKKTraCuuInfo.ID);
								if (hTKKTraCuuInfo.ID == "P7")
								{
									text22 = "D";
								}
							}
						}
						if (!(hTKKTraCuuInfo.LoaiTK == text22))
						{
							continue;
						}
						try
						{
							List<int> list3 = new List<int>();
							if ("66;67;09;10;60;61;G2;".IndexOf(hTKKTraCuuInfo.ID) > -1)
							{
								text8 = DateTime.Now.ToString("yyyy");
								goto IL_4649;
							}
							if (hTKKTraCuuInfo.LoaiTK == "M")
							{
								if (text19.Length == 6)
								{
									text3 = HTKKFormater.toDateMMYYYY(text19, isMess: false);
									kyKeKhai = text3;
									list3.Add(HTKKFormater.FormatStringToDateNumberBegin(text19));
									text9 = text19.Substring(0, 2);
									text8 = text19.Substring(2, 4);
									text4 = "01/" + text9 + "/" + text8;
									text5 = DateTime.DaysInMonth(ConvertHelpers.ConvertToInt(text8), ConvertHelpers.ConvertToInt(text9)) + "/" + text9 + "/" + text8;
									goto IL_4649;
								}
								if (text19.Length == 13)
								{
									if (text19.IndexOf("_") > -1)
									{
										string[] array13 = text19.Split(new char[1] { '_' });
										text3 = ((!(hTKKTraCuuInfo.ID == "99") && !(hTKKTraCuuInfo.ID == "H6")) ? (HTKKFormater.toDateMMYYYY(array13[0], isMess: false) + " - " + HTKKFormater.toDateMMYYYY(array13[1], isMess: false)) : ((!(hTKKTraCuuInfo.KieuKyHoanThue == "1")) ? ("Quý " + array13[0].Substring(1, 1) + "/" + array13[0].Substring(2, 4) + " - Tháng " + HTKKFormater.toDateMMYYYY(array13[1], isMess: false)) : ("Tháng " + HTKKFormater.toDateMMYYYY(array13[0], isMess: false) + " - Tháng " + HTKKFormater.toDateMMYYYY(array13[1], isMess: false))));
										list3.Add(HTKKFormater.FormatStringToDateNumberBegin(array13[0]));
										list3.Add(HTKKFormater.FormatStringToDateNumberEnd(array13[1]));
										text8 = array13[0].Substring(2, 4);
										text6 = HTKKFormater.toDateMMYYYY(array13[0], isMess: false);
										text7 = HTKKFormater.toDateMMYYYY(array13[1], isMess: false);
										kyKeKhai = ((!(hTKKTraCuuInfo.ID.Substring(0, 2) == "99") && !(hTKKTraCuuInfo.ID.Substring(0, 2) == "H6")) ? text6 : text7);
									}
									goto IL_4649;
								}
								continue;
							}
							if (hTKKTraCuuInfo.LoaiTK == "D")
							{
								if (text19.Length == 8)
								{
									text3 = HTKKFormater.toDateDDMMYYYY(text19, isMess: false);
									kyKeKhai = text3;
									if (hTKKTraCuuInfo.ID.Substring(0, 2) != "Y5")
									{
										text4 = text3;
									}
									text5 = text3;
									list3.Add(HTKKFormater.FormatStringToDateNumberBegin(text19));
									text11 = text19.Substring(0, 2);
									text9 = text19.Substring(2, 2);
									text8 = text19.Substring(4, 4);
									goto IL_4649;
								}
								if (text19.Length == 17)
								{
									if (text19.IndexOf("_") > 0)
									{
										if (hTKKTraCuuInfo.ID == "631" || hTKKTraCuuInfo.ID == "L11")
										{
											string[] array13 = text19.Split(new char[1] { '_' });
											text3 = array13[0];
											kyKeKhai = text3;
											text11 = HTKKFormater.FormatStringToDate(array13[2]);
											list3.Add(HTKKFormater.FormatStringToDateNumberBegin(array13[2]));
											list3.Add(HTKKFormater.FormatStringToDateNumberEnd(array13[2]));
											text8 = array13[0];
											text14 = " - Cơ sở mới ";
											text22 = "Y";
											if (hTKKTraCuuInfo.ID == "L11")
											{
												text14 = " - Ngày phát sinh: " + text11;
												HTKKFile.AddOrUpdateDicStr(hTKKTraCuuInfo.dicVar, "CoSoMoiThanhLap", "true");
												HTKKFile.AddOrUpdateDicStr(hTKKTraCuuInfo.dicVar, "NgayThanhLap", text11);
											}
										}
										else
										{
											string[] array13 = text19.Split(new char[1] { '_' });
											if (array13[0].Replace("/", "").Length >= 8)
											{
												text8 = array13[0].Replace("/", "").Substring(4, 4);
												text3 = text8;
												kyKeKhai = text3;
											}
											text4 = HTKKFormater.toDateDDMMYYYY(array13[0], isMess: false);
											text5 = HTKKFormater.toDateDDMMYYYY(array13[1], isMess: false);
											list3.Add(HTKKFormater.FormatStringToDateNumberBegin(array13[0]));
											list3.Add(HTKKFormater.FormatStringToDateNumberEnd(array13[1]));
											text22 = "Y";
										}
									}
									goto IL_4649;
								}
								if (text19.Length == 11)
								{
									if (text19.IndexOf("_") > 0)
									{
										string[] array13 = text19.Split(new char[1] { '_' });
										if (array13[1].Replace("/", "").Length >= 8)
										{
											text3 = array13[1].Substring(0, 2) + "/" + array13[1].Substring(2, 2) + "/" + array13[1].Substring(4, 4);
											kyKeKhai = text3;
											text11 = array13[1].Substring(0, 2);
											text9 = array13[1].Substring(2, 2);
											text8 = array13[1].Substring(4, 4);
											text4 = text3;
											text5 = text3;
										}
										if (array13[0].Length > 1)
										{
											empty3 = array13[0].Replace("L", "");
											text14 = " - lần " + empty3;
											hTKKTraCuuInfo.LanXuatBan = empty3;
											if (hTKKTraCuuInfo.ID.Substring(0, 2) == "P1" || hTKKTraCuuInfo.ID.Substring(0, 2) == "P6" || hTKKTraCuuInfo.ID.Substring(0, 2) == "X7" || hTKKTraCuuInfo.ID.Substring(0, 2).Equals("Y9"))
											{
												text12 = empty3;
											}
										}
										list3.Add(HTKKFormater.FormatStringToDateNumberBegin(array13[1]));
										list3.Add(HTKKFormater.FormatStringToDateNumberEnd(array13[1]));
									}
									goto IL_4649;
								}
								if (hTKKTraCuuInfo.ID == "P7")
								{
									text19 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/ngayLapTKhai']");
									text19 = text19.Replace("/", "");
									text3 = HTKKFormater.toDateDDMMYYYY(text19, isMess: false);
									kyKeKhai = text3;
									text4 = text3;
									text5 = text3;
									list3.Add(HTKKFormater.FormatStringToDateNumberBegin(text19));
									text11 = text19.Substring(0, 2);
									text9 = text19.Substring(2, 2);
									text8 = text19.Substring(4, 4);
									goto IL_4649;
								}
								continue;
							}
							if (hTKKTraCuuInfo.LoaiTK == "Y")
							{
								if (text19.Length == 4)
								{
									text3 = text19;
									kyKeKhai = text3;
									list3.Add(HTKKFormater.FormatStringToDateNumberBegin(text19));
									text8 = text19;
									if (hTKKTraCuuInfo.ID == "46")
									{
										text4 = "01/01/" + text8;
										text5 = "31/12/" + text8;
									}
									goto IL_4649;
								}
								if (text19.Length == 18)
								{
									if (text19.IndexOf("_") > 0)
									{
										string[] array13 = text19.Split(new char[1] { '_' });
										text3 = HTKKFormater.toDateMMYYYY(array13[1], isMess: false) + " - " + HTKKFormater.toDateMMYYYY(array13[2], isMess: false);
										kyKeKhai = array13[0];
										text6 = array13[1].Substring(0, 2) + "/" + array13[1].Substring(2, 4);
										text7 = array13[2].Substring(0, 2) + "/" + array13[2].Substring(2, 4);
										list3.Add(HTKKFormater.FormatStringToDateNumberBegin(array13[1]));
										list3.Add(HTKKFormater.FormatStringToDateNumberEnd(array13[2]));
										text8 = array13[0];
										goto IL_4649;
									}
								}
								else if (text19.Length == 7)
								{
									text3 = text19;
									if (text19.IndexOf("_") >= 1)
									{
										text3 = text19.Split(new char[1] { '_' })[1];
										kyKeKhai = text3;
										list3.Add(HTKKFormater.FormatStringToDateNumberBegin(text19.Split(new char[1] { '_' })[1]));
										text8 = text19.Split(new char[1] { '_' })[1];
										if (text19.Split(new char[1] { '_' })[0].Length > 1)
										{
											text12 = text19.Split(new char[1] { '_' })[0].Substring(1, 1);
											text14 = " - lần " + text12;
										}
										goto IL_4649;
									}
								}
								else if (text19.Length == 9)
								{
									text3 = text19;
									text3 = "Từ năm " + text19.Split(new char[1] { '_' })[0] + " đến năm " + text19.Split(new char[1] { '_' })[1];
									kyKeKhai = text3;
									list3.Add(HTKKFormater.FormatStringToDateNumberBegin(text19.Split(new char[1] { '_' })[1]));
									text8 = text19.Split(new char[1] { '_' })[1];
									text6 = text19.Split(new char[1] { '_' })[0];
									text7 = text19.Split(new char[1] { '_' })[1];
									goto IL_4649;
								}
								continue;
							}
							if (hTKKTraCuuInfo.LoaiTK == "Q")
							{
								if (text19.Length == 6)
								{
									text8 = text19.Substring(2, 4);
									text10 = text19.Substring(1, 1);
									text3 = "Quý " + text10 + "/" + text8;
									kyKeKhai = text10 + "/" + text8;
									if (hTKKTraCuuInfo.ID == "18")
									{
										text4 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayDauKyBC']");
										text5 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayCuoiKyBC']");
										list3.Add(HTKKFormater.FormatStringToDateNumberBegin(text4));
										list3.Add(HTKKFormater.FormatStringToDateNumberEnd(text5));
									}
									else
									{
										HTKKPeriodInfo hTKKPeriodInfo = new HTKKPeriodInfo();
										hTKKPeriodInfo.Nam = text8;
										hTKKPeriodInfo.KieuKy = "Q";
										hTKKPeriodInfo.Quy = text10;
										list3.Add(HTKKFormater.FormatStringToDateNumberBegin(NgayDauKyTinhThue(hTKKPeriodInfo).ToString("dd/MM/yyyy")));
										list3.Add(HTKKFormater.FormatStringToDateNumberEnd(NgayCuoiKyTinhThue(hTKKPeriodInfo).ToString("dd/MM/yyyy")));
									}
									goto IL_4649;
								}
								if (text19.Length == 7)
								{
									text8 = text19.Substring(3, 4);
									text10 = text19.Substring(2, 1);
									text3 = "Quý " + text19.Substring(2, 1) + "/" + text8;
									kyKeKhai = text19.Substring(2, 1) + "/" + text8;
									HTKKPeriodInfo hTKKPeriodInfo = new HTKKPeriodInfo();
									hTKKPeriodInfo.Nam = text8;
									hTKKPeriodInfo.KieuKy = "Q";
									hTKKPeriodInfo.Quy = text10;
									text4 = NgayDauKyTinhThue(hTKKPeriodInfo).ToString("dd/MM/yyyy");
									text5 = NgayCuoiKyTinhThue(hTKKPeriodInfo).ToString("dd/MM/yyyy");
									list3.Add(HTKKFormater.FormatStringToDateNumberBegin(NgayDauKyTinhThue(hTKKPeriodInfo).ToString("dd/MM/yyyy")));
									list3.Add(HTKKFormater.FormatStringToDateNumberEnd(NgayCuoiKyTinhThue(hTKKPeriodInfo).ToString("dd/MM/yyyy")));
									goto IL_4649;
								}
								if (text19.Length == 14)
								{
									if (text19.IndexOf("_") > -1)
									{
										string text27 = text19.Replace("Q", "");
										string[] array13 = text27.Split(new char[1] { '_' });
										text10 = array13[0].Substring(1, 1);
										text8 = array13[0].Substring(2, 4);
										text6 = array13[0].Substring(1, 1) + "/" + array13[0].Substring(2, 4);
										text7 = array13[1].Substring(1, 1) + "/" + array13[1].Substring(2, 4);
										text3 = "Quý " + text6 + " - " + text7;
										kyKeKhai = text6;
										list3.Add(HTKKFormater.FormatStringToDateNumberBegin(ConvertHelpers.GetNgayDauQuy(ConvertHelpers.ConvertToInt(array13[0].Substring(1, 1)), ConvertHelpers.ConvertToInt(array13[0].Substring(2, 4)), 1, 1).ToString("dd/MM/yyyy")));
										list3.Add(HTKKFormater.FormatStringToDateNumberEnd(ConvertHelpers.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(array13[1].Substring(1, 1)), ConvertHelpers.ConvertToInt(array13[1].Substring(2, 4)), 1, 1).ToString("dd/MM/yyyy")));
									}
									goto IL_4649;
								}
								if (text19.Length == 13)
								{
									if (text19.IndexOf("_") > -1)
									{
										string text27 = text19;
										string[] array13 = text27.Split(new char[1] { '_' });
										text10 = array13[0].Substring(1, 1);
										text8 = array13[0].Substring(2, 4);
										text6 = array13[0].Substring(1, 1) + "/" + array13[0].Substring(2, 4);
										text7 = array13[1].Substring(1, 1) + "/" + array13[1].Substring(2, 4);
										if (hTKKTraCuuInfo.ID == "99" || hTKKTraCuuInfo.ID == "H6")
										{
											text6 = array13[0].Substring(1, 1) + "/" + array13[0].Substring(2, 4);
											text7 = array13[1].Substring(1, 1) + "/" + array13[1].Substring(2, 4);
											if (hTKKTraCuuInfo.KieuKyHoanThue == "2")
											{
												text3 = "Quý " + text6 + " - Quý " + text7;
											}
											else
											{
												text6 = array13[0].Substring(0, 2) + "/" + array13[0].Substring(2, 4);
												text7 = array13[1].Substring(1, 1) + "/" + array13[1].Substring(2, 4);
												text3 = "Tháng " + text6 + " - Quý " + text7;
											}
										}
										else
										{
											text3 = "Quý " + text6 + " - " + text7;
										}
										kyKeKhai = ((!(hTKKTraCuuInfo.ID.Substring(0, 2) == "99") && !(hTKKTraCuuInfo.ID.Substring(0, 2) == "H6")) ? text6 : text7);
										list3.Add(HTKKFormater.FormatStringToDateNumberBegin(ConvertHelpers.GetNgayDauQuy(ConvertHelpers.ConvertToInt(array13[0].Substring(1, 1)), ConvertHelpers.ConvertToInt(array13[0].Substring(2, 4)), 1, 1).ToString("dd/MM/yyyy")));
										list3.Add(HTKKFormater.FormatStringToDateNumberEnd(ConvertHelpers.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(array13[1].Substring(1, 1)), ConvertHelpers.ConvertToInt(array13[1].Substring(2, 4)), 1, 1).ToString("dd/MM/yyyy")));
									}
									goto IL_4649;
								}
								if (text19.Length == 21)
								{
									if (text19.IndexOf("_") >= 0)
									{
										string[] array14 = text19.Split(new char[1] { '_' });
										text8 = array14[0].Substring(3, 4);
										text10 = array14[0].Substring(2, 1);
										text3 = "Quý " + array14[0].Substring(2, 1) + "/" + text8;
										kyKeKhai = array14[0].Substring(2, 1) + "/" + text8;
										text6 = HTKKFormater.FormatStringToDate(array14[1]);
										text7 = HTKKFormater.FormatStringToDate(array14[2]);
										HTKKPeriodInfo hTKKPeriodInfo = new HTKKPeriodInfo();
										hTKKPeriodInfo.Nam = text8;
										hTKKPeriodInfo.KieuKy = "Q";
										hTKKPeriodInfo.Quy = text10;
										list3.Add(HTKKFormater.FormatStringToDateNumberBegin(NgayDauKyTinhThue(hTKKPeriodInfo).ToString("dd/MM/yyyy")));
										list3.Add(HTKKFormater.FormatStringToDateNumberEnd(NgayCuoiKyTinhThue(hTKKPeriodInfo).ToString("dd/MM/yyyy")));
										goto IL_4649;
									}
								}
								else
								{
									if (text19.Length != 25 || text19.IndexOf("_") < 0)
									{
										continue;
									}
									string[] array14 = text19.Split(new char[1] { '_' });
									text8 = array14[0].Substring(3, 4);
									text10 = array14[0].Substring(2, 1);
									text3 = "Quý " + array14[0].Substring(2, 1) + "/" + text8;
									kyKeKhai = array14[0].Substring(2, 1) + "/" + text8;
									text4 = HTKKFormater.FormatStringToDate(array14[1]);
									text5 = HTKKFormater.FormatStringToDate(array14[2]);
									HTKKPeriodInfo hTKKPeriodInfo = new HTKKPeriodInfo();
									hTKKPeriodInfo.Nam = text8;
									hTKKPeriodInfo.KieuKy = "Q";
									hTKKPeriodInfo.Quy = text10;
									list3.Add(HTKKFormater.FormatStringToDateNumberBegin(NgayDauKyTinhThue(hTKKPeriodInfo).ToString("dd/MM/yyyy")));
									list3.Add(HTKKFormater.FormatStringToDateNumberEnd(NgayCuoiKyTinhThue(hTKKPeriodInfo).ToString("dd/MM/yyyy")));
									if (hTKKTraCuuInfo.ID == "A1")
									{
										hTKKTraCuuInfo.DoanhNghiepHoatDongLienTuc = "LT";
										string valueByAttributes2 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='null_4']");
										if (HTKKFormater.isNullString(valueByAttributes2))
										{
											hTKKTraCuuInfo.DoanhNghiepHoatDongLienTuc = "KL";
										}
									}
									if (dataRow["PathFile"].ToString().IndexOf("VND") > 0)
									{
										hTKKTraCuuInfo.DonViTinhTen = "Việt Nam Đồng";
										hTKKTraCuuInfo.DonViTinhMa = "VND";
									}
									else
									{
										hTKKTraCuuInfo.DonViTinhTen = "Ngoại tệ";
										hTKKTraCuuInfo.DonViTinhMa = "NT";
									}
									goto IL_4649;
								}
								continue;
							}
							if (hTKKTraCuuInfo.LoaiTK == "K")
							{
								text8 = text19.Substring(3, 4);
								text10 = text19.Substring(2, 1);
								text3 = "Kỳ " + text10 + "/" + text8;
								kyKeKhai = text10 + "/" + text8;
								text4 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/tuNgay']");
								text5 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/denNgay']");
								if (text4.IsNullOrEmpty())
								{
									text4 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayDauKyBaoCao']");
									text5 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayCuoiKyBaoCao']");
								}
								list3.Add(HTKKFormater.FormatStringToDateNumberBegin(text4));
								list3.Add(HTKKFormater.FormatStringToDateNumberEnd(text5));
								goto IL_4649;
							}
							if (hTKKTraCuuInfo.LoaiTK == "D_Y")
							{
								if (hTKKTraCuuInfo.ID == "78" || hTKKTraCuuInfo.ID == "58")
								{
									string[] array13 = text19.Split(new char[1] { '_' });
									text3 = array13[0];
									kyKeKhai = text3;
									text8 = array13[0];
									HTKKPeriodInfo hTKKPeriodInfo = new HTKKPeriodInfo();
									hTKKPeriodInfo.Nam = text8;
									hTKKPeriodInfo.KieuKy = "Y";
									text22 = "Y";
									hTKKPeriodInfo.TuNgay = array13[1].Substring(0, 2) + "/" + array13[1].Substring(2, 2) + "/" + array13[1].Substring(4, 4);
									hTKKPeriodInfo.DenNgay = array13[2].Substring(0, 2) + "/" + array13[2].Substring(2, 2) + "/" + array13[2].Substring(4, 4);
									list3.Add(HTKKFormater.FormatStringToDateNumberBegin(array13[1]));
									list3.Add(HTKKFormater.FormatStringToDateNumberBegin(array13[2]));
									text4 = hTKKPeriodInfo.TuNgay;
									text5 = hTKKPeriodInfo.DenNgay;
									hTKKTraCuuInfo.DoanhNghiepHoatDongLienTuc = "LT";
									string valueByAttributes2 = HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='null_6']");
									if (valueByAttributes2.Equals("1"))
									{
										hTKKTraCuuInfo.DoanhNghiepHoatDongLienTuc = "KL";
									}
									if (dataRow["PathFile"].ToString().IndexOf("VND") > 0)
									{
										hTKKTraCuuInfo.DonViTinhTen = "Việt Nam Đồng";
										hTKKTraCuuInfo.DonViTinhMa = "VND";
									}
									else
									{
										hTKKTraCuuInfo.DonViTinhTen = "Ngoại tệ";
										hTKKTraCuuInfo.DonViTinhMa = "NT";
									}
									goto IL_4649;
								}
								if (text19.Length == 22)
								{
									string[] array13 = text19.Split(new char[1] { '_' });
									text3 = array13[0];
									kyKeKhai = text3;
									text8 = array13[0];
									HTKKPeriodInfo hTKKPeriodInfo = new HTKKPeriodInfo();
									hTKKPeriodInfo.Nam = text8;
									hTKKPeriodInfo.KieuKy = "Y";
									text22 = "Y";
									hTKKPeriodInfo.TuNgay = array13[1].Substring(0, 2) + "/" + array13[1].Substring(2, 2) + "/" + array13[1].Substring(4, 4);
									hTKKPeriodInfo.DenNgay = array13[2].Substring(0, 2) + "/" + array13[2].Substring(2, 2) + "/" + array13[2].Substring(4, 4);
									list3.Add(HTKKFormater.FormatStringToDateNumberBegin(array13[1]));
									list3.Add(HTKKFormater.FormatStringToDateNumberBegin(array13[2]));
									text4 = hTKKPeriodInfo.TuNgay;
									text5 = hTKKPeriodInfo.DenNgay;
									goto IL_4649;
								}
								continue;
							}
							if (!(hTKKTraCuuInfo.LoaiTK == "D_D"))
							{
								goto IL_4649;
							}
							if (text19.Length != 17)
							{
								continue;
							}
							if (text19.IndexOf("_") > 0)
							{
								string[] array13 = text19.Split(new char[1] { '_' });
								text4 = array13[0].Substring(0, 2) + "/" + array13[0].Substring(2, 2) + "/" + array13[0].Substring(4, 4);
								text5 = array13[1].Substring(0, 2) + "/" + array13[1].Substring(2, 2) + "/" + array13[1].Substring(4, 4);
								list3.Add(HTKKFormater.FormatStringToDateNumberBegin(array13[0]));
								list3.Add(HTKKFormater.FormatStringToDateNumberEnd(array13[1]));
								text3 = text4 + " - " + text5;
								kyKeKhai = text4;
								if (array13[0].Replace("/", "").Length >= 8)
								{
									text8 = array13[0].Replace("/", "").Substring(4, 4);
								}
							}
							goto IL_4649;
							IL_4649:
							list2.Add(dataRow["FileName"].ToString());
							empty8 = (hTKKTraCuuInfo.MaHSDuAnDT.Equals("") ? "" : ((hTKKTraCuuInfo.ID.Substring(0, 2) == "D2") ? (" - Mã DAĐT: " + hTKKTraCuuInfo.MaHSDuAnDT + " - Tên DAĐT: " + HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct06']") + " ") : ((hTKKTraCuuInfo.ID.Substring(0, 2) == "P6" || hTKKTraCuuInfo.ID.Substring(0, 2) == "P1" || hTKKTraCuuInfo.ID.Substring(0, 2) == "X7" || hTKKTraCuuInfo.ID.Substring(0, 2) == "Y9") ? (" - Số chứng từ nộp NSNN: " + hTKKTraCuuInfo.MaHSDuAnDT + " ") : (("~O1~O2~O3~O4~O5~".IndexOf("~" + hTKKTraCuuInfo.ID + "~") < 0) ? (" - Mã DAĐT: " + hTKKTraCuuInfo.MaHSDuAnDT + " - Tên DAĐT: " + HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TTDuAnDauTu/tenDuAn']") + " ") : (" - Mã số thuế: " + hTKKTraCuuInfo.MaHSDuAnDT + " - Tên đơn vị: " + HTKKFile.GetValueByAttributes(dataRow["PathFile"].ToString(), "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ten_DN_DVi']") + " ")))));
							if (!cqtCapTion.Equals(""))
							{
								if (text.Contains("TINH"))
								{
									cqtCapTion = " - Tên tỉnh: " + cqtCapTion;
								}
								else if (text.Contains("CQTQL"))
								{
									cqtCapTion = " - Tên CQT: " + cqtCapTion;
								}
							}
							if ("~K2~K3~J7~E6~E7~D7~R6~X5~R9~E9~".IndexOf("~" + hTKKTraCuuInfo.ID.Substring(0, 2) + "~") >= 0 && flag)
							{
								text17 = " - Loại đơn vị tiền: " + text16;
							}
							if (hTKKTraCuuInfo.ID.Substring(0, 2) == "U8" && text.IndexOf("tt") == 0)
							{
								text13 = text13.Substring(0, 8);
								text12 = "";
							}
							HTKKTraCuuInfo hTKKTraCuuInfo2 = new HTKKTraCuuInfo();
							hTKKTraCuuInfo2.ID = hTKKTraCuuInfo.ID;
							hTKKTraCuuInfo2.FinanceYear = hTKKTraCuuInfo.FinanceYear;
							hTKKTraCuuInfo2.Year = text8;
							hTKKTraCuuInfo2.Month = hTKKTraCuuInfo.Month;
							hTKKTraCuuInfo2.ThreeMonth = hTKKTraCuuInfo.Month;
							hTKKTraCuuInfo2.Day = hTKKTraCuuInfo.Day;
							hTKKTraCuuInfo2.Caption = hTKKTraCuuInfo.Caption + text18 + text17 + text15 + cqtCapTion + empty8 + text14 + text13;
							hTKKTraCuuInfo2.LoaiTK = empty4;
							hTKKTraCuuInfo2.LoaiTKShow = empty5;
							hTKKTraCuuInfo2.TkThangQuy = hTKKTraCuuInfo.TkThangQuy;
							hTKKTraCuuInfo2.StartDate = validity[i].StartDate;
							hTKKTraCuuInfo2.EndDate = validity[i].EndDate;
							hTKKTraCuuInfo2.DataFile = validity[i].DataFile;
							hTKKTraCuuInfo2.ThueKhauTru = validity[i].ThueKhauTru;
							hTKKTraCuuInfo2.ThuePhaiNop = validity[i].ThuePhaiNop;
							hTKKTraCuuInfo2.FileName = dataRow["FileName"].ToString();
							hTKKTraCuuInfo2.KieuKy = text22;
							hTKKTraCuuInfo2.KyTinhThueShow = text3;
							hTKKTraCuuInfo2.KyKeKhai = kyKeKhai;
							hTKKTraCuuInfo2.VersionAPP = empty;
							hTKKTraCuuInfo2.VersionXML = versionXML;
							hTKKTraCuuInfo2.TenToKhai = text2;
							hTKKTraCuuInfo2.Lan = text12;
							hTKKTraCuuInfo2.TuNgay = text4;
							hTKKTraCuuInfo2.DenNgay = text5;
							hTKKTraCuuInfo2.TuThang = text6;
							hTKKTraCuuInfo2.DenThang = text7;
							hTKKTraCuuInfo2.Thang = text9;
							hTKKTraCuuInfo2.Quy = text10;
							hTKKTraCuuInfo2.Ngay = text11;
							hTKKTraCuuInfo2.NgayLapKHBS = empty2;
							hTKKTraCuuInfo2.ListDateNumber = list3;
							hTKKTraCuuInfo2.DoanhNghiepHoatDongLienTuc = hTKKTraCuuInfo.DoanhNghiepHoatDongLienTuc;
							hTKKTraCuuInfo2.DonViTinhMa = hTKKTraCuuInfo.DonViTinhMa;
							hTKKTraCuuInfo2.DonViTinhTen = hTKKTraCuuInfo.DonViTinhTen;
							hTKKTraCuuInfo2.TieuMucHachToan = hTKKTraCuuInfo.TieuMucHachToan;
							hTKKTraCuuInfo2.NganhNgheKD = hTKKTraCuuInfo.NganhNgheKD;
							hTKKTraCuuInfo2.ThuHo = hTKKTraCuuInfo.ThuHo;
							hTKKTraCuuInfo2.MaHSDuAnDT = hTKKTraCuuInfo.MaHSDuAnDT;
							hTKKTraCuuInfo2.LanXuatBan = hTKKTraCuuInfo.LanXuatBan;
							hTKKTraCuuInfo2.Validity = hTKKTraCuuInfo.Validity;
							hTKKTraCuuInfo2.dicVar = hTKKTraCuuInfo.dicVar;
							hTKKTraCuuInfo2.tenCQTCapCuc = hTKKTraCuuInfo.tenCQTCapCuc;
							hTKKTraCuuInfo2.maCQTCapCuc = hTKKTraCuuInfo.maCQTCapCuc;
							hTKKTraCuuInfo2.maCQTQuanLy = hTKKTraCuuInfo.maCQTQuanLy;
							hTKKTraCuuInfo2.tenCQTQuanLy = hTKKTraCuuInfo.tenCQTQuanLy;
							hTKKTraCuuInfo2.maTinh = hTKKTraCuuInfo.maTinh;
							hTKKTraCuuInfo2.tenTinh = hTKKTraCuuInfo.tenTinh;
							hTKKTraCuuInfo2.maHuyen = hTKKTraCuuInfo.maHuyen;
							hTKKTraCuuInfo2.tenHuyen = hTKKTraCuuInfo.tenHuyen;
							hTKKTraCuuInfo2.maXa = hTKKTraCuuInfo.maXa;
							hTKKTraCuuInfo2.tenXa = hTKKTraCuuInfo.tenXa;
							hTKKTraCuuInfo2.NopTheoTinhCQT = hTKKTraCuuInfo.NopTheoTinhCQT;
							hTKKTraCuuInfo2.KieuKyHoanThue = hTKKTraCuuInfo.KieuKyHoanThue;
							hTKKTraCuuInfo2.ToChucCNDNVonChuyenNhuong = hTKKTraCuuInfo.ToChucCNDNVonChuyenNhuong;
							HTKKTraCuuInfo item = hTKKTraCuuInfo2;
							list.Add(item);
						}
						catch
						{
							Exception ex = new Exception("Loi boc tach ten dataFile: " + text);
							HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
						}
					}
				}
			}
			return list;
		}
		catch (Exception ex2)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex2);
			throw;
		}
	}

	public static string GetCatalogueDMNNKDItem(string textSearch)
	{
		//IL_0031: Unknown result type (might be due to invalid IL or missing references)
		//IL_0037: Expected O, but got Unknown
		//IL_0061: Unknown result type (might be due to invalid IL or missing references)
		//IL_0067: Expected O, but got Unknown
		try
		{
			if (string.IsNullOrEmpty(textSearch))
			{
				return null;
			}
			if (!File.Exists("..\\InterfaceIni\\Catalogue_DM_NNKD.xml"))
			{
				return null;
			}
			XmlDocument val = new XmlDocument();
			val.Load("..\\InterfaceIni\\Catalogue_DM_NNKD.xml");
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				XmlAttribute val3 = val2.Attributes["Value"];
				if (val3 != null && ((XmlNode)val3).Value.IndexOf(textSearch, StringComparison.Ordinal) >= 0)
				{
					return ((XmlNode)val3).Value;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return null;
	}

	public static HTKKTraCuuInfo RevertNameNopTheoTinhCQT(HTKKTraCuuInfo info, string fullDataFileName, ref string kyKeKhai, ref string cqtCapTion)
	{
		//IL_006f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0075: Expected O, but got Unknown
		try
		{
			if (fullDataFileName.Contains("TINH") || fullDataFileName.Contains("CQTQL"))
			{
				string text = "menu.xml";
				if (File.Exists(text))
				{
					XmlDocument val = new XmlDocument();
					val.LoadXml(HTKKFile.LoadFile(text));
					XmlNodeList source = ((XmlNode)val).SelectNodes("/Root/Menu");
					string idToKhai = "";
					if (info.ID != null && info.ID.Length > 2)
					{
						idToKhai = info.ID.Substring(0, 2);
					}
					else
					{
						idToKhai = info.ID;
					}
					XmlNode val2 = ((IEnumerable)source).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["ID"]).Value == idToKhai));
					if (val2 != null && val2.Attributes["NopTheoTinhCQT"] != null)
					{
						info.NopTheoTinhCQT = ((XmlNode)val2.Attributes["NopTheoTinhCQT"]).Value.Trim();
						if (fullDataFileName.Contains("TINH"))
						{
							string[] array = fullDataFileName.Split(new char[1] { '_' });
							string[] array2 = array;
							foreach (string text2 in array2)
							{
								if (!text2.Contains("TINH"))
								{
									continue;
								}
								info.maTinh = text2.Replace("TINH", "");
								KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == info.maTinh);
								info.tenTinh = keyValuePair.Key;
								cqtCapTion = keyValuePair.Key;
								string text3 = "TINH" + info.maTinh;
								kyKeKhai = kyKeKhai.Replace(text3 + "_", "");
								string maCQTCapCuc = info.maTinh + "00";
								if (HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.ContainsValue(maCQTCapCuc))
								{
									if (info.NopTheoTinhCQT == "3")
									{
										CQTDiaBan cQTDiaBan = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == maCQTCapCuc);
										if (cQTDiaBan != null)
										{
											info.maCQTCapCuc = maCQTCapCuc;
											info.tenCQTCapCuc = cQTDiaBan.TenCQT;
										}
									}
									else
									{
										KeyValuePair<string, string> keyValuePair2 = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maCQTCapCuc);
										if (!string.IsNullOrEmpty(keyValuePair2.Value))
										{
											info.maCQTCapCuc = maCQTCapCuc;
											info.tenCQTCapCuc = keyValuePair2.Key;
										}
									}
								}
								string maCQTQuanLy = info.maTinh + "00";
								if (!HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy.ContainsKey(maCQTQuanLy))
								{
									break;
								}
								if (info.NopTheoTinhCQT == "3")
								{
									CQTDiaBan cQTDiaBan = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == maCQTQuanLy);
									if (cQTDiaBan != null)
									{
										info.maCQTQuanLy = maCQTQuanLy;
										info.tenCQTQuanLy = cQTDiaBan.TenCQT;
									}
								}
								else
								{
									info.maCQTQuanLy = maCQTQuanLy;
									info.tenCQTQuanLy = HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy[maCQTQuanLy];
								}
								break;
							}
						}
						if (fullDataFileName.Contains("CQTQL"))
						{
							string[] array = fullDataFileName.Split(new char[1] { '_' });
							string[] array2 = array;
							foreach (string text2 in array2)
							{
								if (!text2.Contains("CQTQL"))
								{
									continue;
								}
								info.maTinh = "";
								info.tenTinh = "";
								info.maCQTQuanLy = text2.Replace("CQTQL", "");
								info.maTinh = HTKKWinDictionaryDiaDanh.GetMaTinhMoiByMaTinhCu(info.maCQTQuanLy.Substring(0, 3));
								string key = HTKKWinDictionaryDiaDanh.dicTinhTPMoi.FirstOrDefault((KeyValuePair<string, string> kv) => kv.Value == info.maTinh).Key;
								info.tenTinh = key;
								if (info.NopTheoTinhCQT == "3")
								{
									CQTDiaBan cQTDiaBan = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == info.maCQTQuanLy);
									if (cQTDiaBan != null)
									{
										info.tenCQTQuanLy = cQTDiaBan.TenCQT;
									}
								}
								else
								{
									KeyValuePair<string, string> keyValuePair3 = HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy.FirstOrDefault((KeyValuePair<string, string> x) => x.Key == info.maCQTQuanLy);
									info.tenCQTQuanLy = keyValuePair3.Value;
								}
								string maCQTCapCuc2 = info.maCQTQuanLy.Substring(0, info.maCQTQuanLy.Length - 2) + "00";
								if (info.NopTheoTinhCQT == "3")
								{
									CQTDiaBan cQTDiaBan = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == maCQTCapCuc2);
									if (cQTDiaBan != null)
									{
										info.maCQTCapCuc = maCQTCapCuc2;
										info.tenCQTCapCuc = cQTDiaBan.TenCQT;
									}
								}
								else
								{
									KeyValuePair<string, string> keyValuePair4 = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maCQTCapCuc2);
									if (!string.IsNullOrEmpty(keyValuePair4.Value))
									{
										info.maCQTCapCuc = maCQTCapCuc2;
										info.tenCQTCapCuc = keyValuePair4.Key;
									}
								}
								cqtCapTion = info.tenCQTQuanLy;
								string text3 = "CQTQL" + info.maCQTQuanLy;
								kyKeKhai = kyKeKhai.Replace(text3 + "_", "");
								break;
							}
						}
					}
				}
			}
			return info;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static DateTime NgayDauKyTinhThue(HTKKPeriodInfo periodInfo)
	{
		try
		{
			DateTime result = default(DateTime);
			switch (periodInfo.KieuKy)
			{
			case "M":
				result = new DateTime(int.Parse(periodInfo.Nam), int.Parse(periodInfo.Thang), 1);
				break;
			case "Q":
			{
				int month = 3 * (int.Parse(periodInfo.Quy) - 1) + 1;
				result = new DateTime(int.Parse(periodInfo.Nam), month, 1);
				break;
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

	public static DateTime NgayCuoiKyTinhThue(HTKKPeriodInfo periodInfo)
	{
		try
		{
			DateTime result = default(DateTime);
			switch (periodInfo.KieuKy)
			{
			case "M":
				result = new DateTime(int.Parse(periodInfo.Nam), int.Parse(periodInfo.Thang), DateTime.DaysInMonth(int.Parse(periodInfo.Nam), int.Parse(periodInfo.Thang)));
				break;
			case "Q":
			{
				int month = 3 * int.Parse(periodInfo.Quy);
				result = new DateTime(int.Parse(periodInfo.Nam), month, DateTime.DaysInMonth(int.Parse(periodInfo.Nam), month));
				break;
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
}
