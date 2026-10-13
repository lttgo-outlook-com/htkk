using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace TAX.HTKK.Utilities.Core;

public class HTKKWinDictionaryDiaDanh
{
	public static string[] TinhTP;

	public static string[] TinhTP_VN;

	public static Dictionary<string, string> dicTinhTP = new Dictionary<string, string>();

	public static Dictionary<string, string> dicQuanHuyen = new Dictionary<string, string>();

	public static Dictionary<string, string[]> dicQuanHuyenByTinhTP = new Dictionary<string, string[]>();

	public static Dictionary<string, string[]> dicPhuongXaByQuanHuyen = new Dictionary<string, string[]>();

	public static Dictionary<string, string> dicPhuongXa = new Dictionary<string, string>();

	public static Dictionary<string, string> CoQuanThueCapCuc = new Dictionary<string, string>();

	public static Dictionary<string, string> CoQuanThueQuanLy = new Dictionary<string, string>();

	public static Dictionary<string, string> QuanHe = new Dictionary<string, string>();

	public static Dictionary<string, string> QuocTich = new Dictionary<string, string>();

	public static Dictionary<string, string> QuocGia = new Dictionary<string, string>();

	public static Dictionary<string, string> QuocTichFull = new Dictionary<string, string>();

	public static Dictionary<string, string> QuocTichLangGieng = new Dictionary<string, string>();

	public static Dictionary<string, string> DMNganhNgheCoDoanhThuCaoNhat = new Dictionary<string, string>();

	public static Dictionary<string, string> PhuongPhapXacDinhGia = new Dictionary<string, string>();

	public static Dictionary<string, string> LyDoGiaHan = new Dictionary<string, string>();

	public static Dictionary<string, string> GioiTinh = new Dictionary<string, string>();

	public static Dictionary<string, string> LoaiGiayToTuyThan = new Dictionary<string, string>();

	public static Dictionary<string, string> NoiCap = new Dictionary<string, string>();

	public static Dictionary<string, string> LinhVucHD = new Dictionary<string, string>();

	public static Dictionary<string, string> dicMaTieuMuc = new Dictionary<string, string>();

	public static Dictionary<string, string> dicMaChuong = new Dictionary<string, string>();

	public static string[] NoiCapCCCD;

	public static Dictionary<string, string> LoaiHoanThue = new Dictionary<string, string>();

	public static Dictionary<string, string> LyDoHoanThue = new Dictionary<string, string>();

	public static Dictionary<string, string> KhoanNopNSSHoanThue = new Dictionary<string, string>();

	public static Dictionary<string, string> KyPhatSinhHoanThue = new Dictionary<string, string>();

	public static Dictionary<string, string> dicTinhTPMoi = new Dictionary<string, string>();

	public static Dictionary<string, string> dicThueTinh = new Dictionary<string, string>();

	public static List<DiaBanXa> _diaBanXa = new List<DiaBanXa>();

	public static List<CQTDiaBan> _cqtDiaBan = new List<CQTDiaBan>();

	public static List<CoQuanThu> coQuanThus = new List<CoQuanThu>();

	public static List<ThueTinh> thueTinhs = new List<ThueTinh>();

	public static Dictionary<string, string> ThueSuatHoanThue = new Dictionary<string, string>();

	public HTKKWinDictionaryDiaDanh()
	{
		CreateQuanHe();
		CreateQuocTich();
		CreateQuocGia();
		CreateQuocTichFull();
		CreateTinhTP();
		CreateNoiCap();
		CreateNoiCapCCCD();
		CreatedicTinhTP();
		CreateQuanHuyen();
		CreatePhuongXa();
		CreateQuanHuyenByTinhTP();
		CreatePhuongXaByQuanHuyen();
		CreateCoQuanThueCapCuc();
		CreateCoQuanThueQuanLy();
		CreateDMNganhNgheCoDoanhThuCaoNhat();
		CreatePhuongPhapXacDinhGia();
		CreateLyDoGiaHan();
		CreateGioiTinh();
		CreateLoaiGiayToTuyThan();
		CreateLoaiHoanThue();
		CreateLyDoHoanThue();
		CreateKhoanNopNSSHoanThue();
		CreateKyPhatSinhHoanThue();
		CreateTinhTP_VN();
		CreateLinhVucHD();
		CreateQuocTichLangGieng();
		CreateThueSuatHoanThue();
		getDanhMucTinhMoi();
		getDanhMucXaMoi();
		getDanhMucCQTDiaBan();
		LoadDanhMucCoQuanThu();
		LoadDanhMucThueTinh();
	}

	private void CreateLyDoGiaHan()
	{
		try
		{
			LyDoGiaHan = CreateDictionaryDm(1, 0, "Catalogue_DM_Lydo_GH_TT16.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CreateQuocGia()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Quoc_Gia_full.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary.Add(key, value2);
				}
			}
			QuocGia = dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CreateQuocTichFull()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0047: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Quoc_Tich_full.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			int num = 0;
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary.Add(key, value2);
				}
			}
			QuocTichFull = dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CreateGioiTinh()
	{
		try
		{
			GioiTinh = CreateDictionaryDm(0, 1, "Catalogue_DM_Gioi_Tinh.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CreateLoaiGiayToTuyThan()
	{
		try
		{
			LoaiGiayToTuyThan = CreateDictionaryDm(0, 1, "Catalogue_DM_Loai_Giay_To_Tuy_Than.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CreateLinhVucHD()
	{
		try
		{
			LinhVucHD = CreateDictionaryDm(0, 1, "Catalogue_Linh_Vuc_HD.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public Dictionary<string, string> getQuanHuyen(string keyTinhTP)
	{
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			foreach (KeyValuePair<string, string> item in dicQuanHuyen)
			{
				if (item.Key.Substring(0, 3).Equals(keyTinhTP))
				{
					dictionary.Add(item.Key, item.Value);
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

	public static Dictionary<string, string> getQuanHuyenByTinh(string keyTinhTP)
	{
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			foreach (KeyValuePair<string, string> item in dicQuanHuyen)
			{
				if (item.Key.Substring(0, 3).Equals(keyTinhTP))
				{
					dictionary.Add(item.Key, item.Value);
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

	public Dictionary<string, string> getPhuongXa(string KeyQuanHuyen)
	{
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			foreach (KeyValuePair<string, string> item in dicPhuongXa)
			{
				if (item.Key.Substring(0, 5).Equals(KeyQuanHuyen))
				{
					dictionary.Add(item.Key, item.Value);
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

	public Dictionary<string, string> getQuocTich()
	{
		return QuocTich;
	}

	private void CreateQuocTich()
	{
		QuocTich.Clear();
		QuocTich.Add("Việt Nam", "01");
		QuocTich.Add("Khác", "02");
		try
		{
			QuocTich = CreateDictionaryDm(1, 0, "Catalogue_Quoc_Gia.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public Dictionary<string, string> getQuanHe()
	{
		return QuanHe;
	}

	public void CreateQuanHe()
	{
		try
		{
			QuanHe = CreateDictionaryDm(1, 0, "Catalogue_QHNNT.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CreateCoQuanThueCapCuc()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0055: Unknown result type (might be due to invalid IL or missing references)
		//IL_005c: Expected O, but got Unknown
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Tinh_Thanh.xml";
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 3 && array[0] == "1")
				{
					string key = array[3];
					string newValue = array[1];
					HTKKFile.AddOrUpdateDicStr(dictionary, key, newValue);
				}
			}
			CoQuanThueCapCuc = dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CreateCoQuanThueQuanLy()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0055: Unknown result type (might be due to invalid IL or missing references)
		//IL_005c: Expected O, but got Unknown
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Tinh_Thanh.xml";
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 3 && array[0] == "0")
				{
					string key = array[1];
					string newValue = array[3];
					HTKKFile.AddOrUpdateDicStr(dictionary, key, newValue);
				}
			}
			CoQuanThueQuanLy = dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CreateDMNganhNgheCoDoanhThuCaoNhat()
	{
		try
		{
			DMNganhNgheCoDoanhThuCaoNhat = CreateDictionaryDm(0, 1, "DM_nganh_TLDT_cao.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CreatePhuongPhapXacDinhGia()
	{
		try
		{
			PhuongPhapXacDinhGia = CreateDictionaryDm(0, 1, "Catalogue_PPDG_TNDN.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static Dictionary<string, string> getTinhTP(string maQuocGia)
	{
		if (maQuocGia != "VN" && maQuocGia != "01")
		{
			return null;
		}
		return dicTinhTPMoi;
	}

	public static string getMaQuanHuyen(string strTenQuanHuyen, string strMaTinh)
	{
		foreach (KeyValuePair<string, string> item in dicQuanHuyen)
		{
			if (item.Value == strTenQuanHuyen && item.Key.Substring(0, 3) == strMaTinh)
			{
				return item.Key;
			}
		}
		return string.Empty;
	}

	public static string getMaPhuongXa(string strTenPhuongXa, string strMaQuanHuyen)
	{
		foreach (KeyValuePair<string, string> item in dicPhuongXa)
		{
			if (item.Value.Trim() == strTenPhuongXa.Trim() && item.Key.Substring(0, 5) == strMaQuanHuyen)
			{
				return item.Key;
			}
		}
		return string.Empty;
	}

	public static Dictionary<string, string> getCoQuanThueQuanLybyMaCoQuanThueCapCuc(string strMaCoQuanThueCapCuc)
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (KeyValuePair<string, string> item in CoQuanThueQuanLy)
		{
			if (item.Key.Substring(0, 3).Equals(strMaCoQuanThueCapCuc.Substring(0, 3)))
			{
				dictionary.Add(item.Key, item.Value);
			}
		}
		return dictionary;
	}

	public static Dictionary<string, string> getCoQuanThueQuanLyCucbyMaCoQuanThueCapCuc(string strMaCoQuanThueCapCuc)
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (KeyValuePair<string, string> item in CoQuanThueQuanLy)
		{
			if (item.Key.Substring(0, 3).Equals(strMaCoQuanThueCapCuc.Substring(0, 3)) && item.Key.Substring(3, 2).Equals("00"))
			{
				dictionary.Add(item.Key, item.Value);
			}
		}
		return dictionary;
	}

	public static Dictionary<string, string> getCoQuanThueQuanLybyMaCoQuanThueCapCucAutocomplete(string strMaCoQuanThueCapCuc)
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (KeyValuePair<string, string> item in CoQuanThueQuanLy)
		{
			if (item.Key.Substring(0, 3).Equals(strMaCoQuanThueCapCuc.Substring(0, 3)))
			{
				dictionary.Add(item.Value, item.Key);
			}
		}
		return dictionary;
	}

	public static string getMaCoQuanThueQuanLy(string strMaCoQuanThueCapCuc, string strTenCoQuanThueQuanLy)
	{
		string result = string.Empty;
		foreach (KeyValuePair<string, string> item in CoQuanThueQuanLy)
		{
			if (item.Key.Substring(0, 3).Equals(strMaCoQuanThueCapCuc.Substring(0, 3)) & strTenCoQuanThueQuanLy.Equals(item.Value))
			{
				result = item.Key;
				break;
			}
		}
		return result;
	}

	public static Dictionary<string, string> CreateQuocTichLangGieng()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Quoc_Gia_Lang_Gieng.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary.Add(key, value2);
				}
			}
			QuocTichLangGieng = dictionary;
			return QuocTichLangGieng;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return new Dictionary<string, string>();
		}
	}

	public static string GetTenCoQuanThueCapCuc(string strMaCoQuanThueCapCuc, bool isGetFullName)
	{
		string text = string.Empty;
		if (CoQuanThueCapCuc.ContainsValue(strMaCoQuanThueCapCuc))
		{
			foreach (KeyValuePair<string, string> item in CoQuanThueCapCuc)
			{
				if (!item.Value.Equals(strMaCoQuanThueCapCuc))
				{
					continue;
				}
				text = item.Key;
				if (!string.IsNullOrEmpty(text) && !isGetFullName)
				{
					string[] array = text.Split(new char[1] { '-' });
					if (array.Length > 1 && array[0].Length == 4)
					{
						text = text.Replace(array[0] + "- ", "");
					}
				}
				break;
			}
		}
		return text;
	}

	public static string GetTenCoQuanThueQuanLy(string strMaCoQuanThue)
	{
		string result = string.Empty;
		if (CoQuanThueQuanLy.ContainsKey(strMaCoQuanThue))
		{
			foreach (KeyValuePair<string, string> item in CoQuanThueQuanLy)
			{
				if (item.Key.Equals(strMaCoQuanThue))
				{
					result = item.Value;
					break;
				}
			}
		}
		return result;
	}

	private void CreateTinhTP()
	{
		TinhTP = new string[64]
		{
			"An Giang", "Bà Rịa - Vũng Tàu", "Bắc Kạn", "Bắc Giang", "Bạc Liêu", "Bắc Ninh", "Bến Tre", "Bình Định", "Bình Dương", "Bình Phước",
			"Bình Thuận", "Cà Mau", "Cần Thơ", "Cao Bằng", "Đà Nẵng", "Đắk Lắk", "Đắk Nông", "Điện Biên", "Đồng Nai", "Đồng Tháp",
			"Gia Lai", "Hà Giang", "Hà Nam", "Hà Nội", "Hà Tĩnh", "Hải Dương", "Hải Phòng", "Hậu Giang", "Hòa Bình", "Hưng Yên",
			"Khánh Hòa", "Kiên Giang", "Kon Tum", "Lai Châu", "Lâm Đồng", "Lạng Sơn", "Lào Cai", "Long An", "Nam Định", "Nghệ An",
			"Ninh Bình", "Ninh Thuận", "Phú Thọ", "Phú Yên", "Quảng Bình", "Quảng Nam", "Quảng Ngãi", "Quảng Ninh", "Quảng Trị", "Sóc Trăng",
			"Sơn La", "Tây Ninh", "Thái Bình", "Thái Nguyên", "Thanh Hoá", "Thành phố Huế", "Tiền Giang", "TP Hồ Chí Minh", "Trà Vinh", "Tuyên Quang",
			"Vĩnh Long", "Vĩnh Phúc", "Yên Bái", "Khác"
		};
	}

	private void CreateTinhTP_VN()
	{
		TinhTP_VN = new string[63]
		{
			"An Giang", "Bà Rịa - Vũng Tàu", "Bắc Kạn", "Bắc Giang", "Bạc Liêu", "Bắc Ninh", "Bến Tre", "Bình Định", "Bình Dương", "Bình Phước",
			"Bình Thuận", "Cà Mau", "Cần Thơ", "Cao Bằng", "Đà Nẵng", "Đắk Lắk", "Đắk Nông", "Điện Biên", "Đồng Nai", "Đồng Tháp",
			"Gia Lai", "Hà Giang", "Hà Nam", "Hà Nội", "Hà Tĩnh", "Hải Dương", "Hải Phòng", "Hậu Giang", "Hòa Bình", "Hưng Yên",
			"Khánh Hòa", "Kiên Giang", "Kon Tum", "Lai Châu", "Lâm Đồng", "Lạng Sơn", "Lào Cai", "Long An", "Nam Định", "Nghệ An",
			"Ninh Bình", "Ninh Thuận", "Phú Thọ", "Phú Yên", "Quảng Bình", "Quảng Nam", "Quảng Ngãi", "Quảng Ninh", "Quảng Trị", "Sóc Trăng",
			"Sơn La", "Tây Ninh", "Thái Bình", "Thái Nguyên", "Thanh Hoá", "Thành phố Huế", "Tiền Giang", "TP Hồ Chí Minh", "Trà Vinh", "Tuyên Quang",
			"Vĩnh Long", "Vĩnh Phúc", "Yên Bái"
		};
	}

	private void CreateNoiCap()
	{
		try
		{
			NoiCap = CreateDictionaryDm(2, 1, "Catalogue_Noi_Cap.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CreateNoiCapCCCD()
	{
		try
		{
			NoiCapCCCD = new string[3] { "Cục c.sát ĐKQL cư trú và DLQG về dân cư", "Cục cảnh sát quản lý hành chính về trật tự xã hội", "Bộ Công an" };
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public Dictionary<string, string> getdicTinhTP()
	{
		return dicTinhTP;
	}

	private void CreatedicTinhTP()
	{
		try
		{
			dicTinhTP = CreateDictionaryDm(2, 1, "Catalogue_Dia_Ban_Tinh.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CreateQuanHuyenByTinhTP()
	{
		try
		{
			List<string> list = new List<string>();
			if (dicTinhTP.Count > 0)
			{
				foreach (KeyValuePair<string, string> item in dicTinhTP)
				{
					list = new List<string>();
					if (item.Value.ToString().Trim().IndexOf("999") > -1)
					{
						continue;
					}
					string text = "Catalogue_Dia_Ban_Huyen_" + item.Value.ToString().Trim() + ".xml";
					if (!File.Exists(HTKKConfig.INTERFACE_INI + "\\" + text))
					{
						continue;
					}
					Dictionary<string, string> dictionary = new Dictionary<string, string>();
					dictionary = CreateDictionaryDm(2, 3, text);
					foreach (KeyValuePair<string, string> item2 in dictionary)
					{
						list.Add(item2.Value);
					}
					if (dicQuanHuyenByTinhTP.TryGetValue(item.Value, out var _))
					{
						dicQuanHuyenByTinhTP[item.Value] = list.ToArray();
					}
					else
					{
						dicQuanHuyenByTinhTP.Add(item.Value, list.ToArray());
					}
				}
			}
			dicQuanHuyenByTinhTP.Add("999", new string[1] { "Khác" });
			list = new List<string>();
			foreach (KeyValuePair<string, string[]> item3 in dicQuanHuyenByTinhTP)
			{
				list.AddRange(item3.Value);
			}
			dicQuanHuyenByTinhTP.Add("ALL", list.ToArray());
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CreatePhuongXaByQuanHuyen()
	{
		try
		{
			List<string> list = new List<string>();
			if (dicTinhTP.Count > 0)
			{
				foreach (KeyValuePair<string, string> item in dicQuanHuyen)
				{
					list = new List<string>();
					if (item.Key.ToString().Trim().IndexOf("999") > -1)
					{
						continue;
					}
					string text = "Catalogue_Dia_Ban_Xa_" + item.Key.ToString().Trim().Substring(0, 3) + ".xml";
					if (!File.Exists(HTKKConfig.INTERFACE_INI + "\\" + text))
					{
						continue;
					}
					Dictionary<string, string> dictionary = new Dictionary<string, string>();
					dictionary = CreateDictionaryDm(3, 4, text);
					foreach (KeyValuePair<string, string> item2 in dictionary)
					{
						if (item2.Key.ToString().Trim().Substring(0, 5) == item.Key)
						{
							list.Add(item2.Value);
						}
					}
					if (dicPhuongXaByQuanHuyen.TryGetValue(item.Key, out var _))
					{
						dicPhuongXaByQuanHuyen[item.Key] = list.ToArray();
					}
					else
					{
						dicPhuongXaByQuanHuyen.Add(item.Key, list.ToArray());
					}
				}
			}
			dicPhuongXaByQuanHuyen.Add("99999", new string[1] { "Khác" });
			list = new List<string>();
			foreach (KeyValuePair<string, string[]> item3 in dicPhuongXaByQuanHuyen)
			{
				list.AddRange(item3.Value);
			}
			dicPhuongXaByQuanHuyen.Add("ALL", list.ToArray());
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CreateQuanHuyen()
	{
		try
		{
			dicQuanHuyen = CreateDictionaryDm(2, 3, "Catalogue_Dia_Ban_Huyen.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CreatePhuongXa()
	{
		try
		{
			dicPhuongXa = CreateDictionaryDm(3, 4, "Catalogue_Dia_Ban_Xa.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string[] getPhuongXa(string keyTinhTP, string tenQuanHuyen)
	{
		string text = string.Empty;
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		try
		{
			foreach (KeyValuePair<string, string> item in dicQuanHuyen)
			{
				if (item.Key.Substring(0, 3).Equals(keyTinhTP) && item.Value.Equals(tenQuanHuyen))
				{
					text = item.Key;
				}
			}
			if (text.Length > 0)
			{
				return dicPhuongXaByQuanHuyen[text];
			}
			return null;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	private Dictionary<string, string> CreateDictionaryDm(int indexKey, int indexValue, string fileNameCatalogue)
	{
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_001f: Expected O, but got Unknown
		//IL_0059: Unknown result type (might be due to invalid IL or missing references)
		//IL_0060: Expected O, but got Unknown
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\" + fileNameCatalogue;
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
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[indexKey];
					string newValue = array[indexValue];
					HTKKFile.AddOrUpdateDicStr(dictionary, key, newValue);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			dictionary = null;
		}
		return dictionary;
	}

	public void CreateLoaiHoanThue()
	{
		try
		{
			LoaiHoanThue = CreateDictionaryDm(1, 2, "Catalogue_Loai_Hoan_Thue.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CreateLyDoHoanThue()
	{
		try
		{
			LyDoHoanThue = CreateDictionaryDm(1, 2, "Catalogue_Ly_Do_Hoan_Thue.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CreateThueSuatHoanThue()
	{
		try
		{
			ThueSuatHoanThue = CreateDictionaryDm(0, 1, "Catalogue_Thue_Suat_Hoan_Thue.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CreateKhoanNopNSSHoanThue()
	{
		try
		{
			KhoanNopNSSHoanThue = CreateDictionaryDm(1, 2, "Catalogue_Hoan_Thue_Khoan_Phai_Nop_NSNN.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CreateKyPhatSinhHoanThue()
	{
		try
		{
			KyPhatSinhHoanThue = CreateDictionaryDm(1, 2, "Catalogue_Ky_Phat_Sinh_Hoan_Thue.xml");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static Dictionary<string, string> GetDataTruongHopQuyetToan(string menuId)
	{
		//IL_004e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0055: Expected O, but got Unknown
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		try
		{
			XmlDataDocument xmlDataDocument = new XmlDataDocument();
			FileStream fileStream = new FileStream(HTKKConfig.INTERFACE_INI + "\\DM_TruongHopQuyetToan.xml", FileMode.Open, FileAccess.Read);
			((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
			XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val = item;
				string text = ((XmlNode)val.Attributes["MenuId"]).Value.ToString().Trim();
				string value = ((XmlNode)val.Attributes["TenQT"]).Value.ToString().Trim();
				string key = ((XmlNode)val.Attributes["MaQT"]).Value.ToString().Trim();
				if (menuId == text && !dictionary.ContainsKey(key))
				{
					dictionary.Add(key, value);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return dictionary;
		}
	}

	private void getDanhMucTinhMoi()
	{
		dicTinhTPMoi = CreateDictionaryDm(2, 1, "Catalogue_Dia_Ban_Tinh_Moi.xml");
	}

	private List<DiaBanXa> getDanhMucXaMoi()
	{
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Expected O, but got Unknown
		//IL_0060: Unknown result type (might be due to invalid IL or missing references)
		//IL_0066: Expected O, but got Unknown
		_diaBanXa.Clear();
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Dia_Ban_Xa_Moi.xml";
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
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					DiaBanXa diaBanXa = new DiaBanXa();
					diaBanXa.MaTinh = array[0].Trim();
					diaBanXa.MaXa = array[1].Trim();
					diaBanXa.TenXa = array[2].Trim();
					_diaBanXa.Add(diaBanXa);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			_diaBanXa.Clear();
		}
		return _diaBanXa;
	}

	public static string getMaPhuongXaMoi(string strTenPhuongXa, string strMaTinh)
	{
		if (_diaBanXa.Count == 0 || string.IsNullOrEmpty(strTenPhuongXa) || string.IsNullOrEmpty(strMaTinh))
		{
			return string.Empty;
		}
		DiaBanXa diaBanXa = _diaBanXa.Find((DiaBanXa x) => x.MaTinh == strMaTinh && x.TenXa == strTenPhuongXa);
		if (diaBanXa != null)
		{
			return diaBanXa.MaXa;
		}
		return string.Empty;
	}

	public static string[] getDanhMucPhuongXaMoi(string keyTinhTP)
	{
		try
		{
			if (_diaBanXa.Count == 0)
			{
				return null;
			}
			if (string.IsNullOrEmpty(keyTinhTP))
			{
				return _diaBanXa.Select((DiaBanXa x) => x.TenXa).ToArray();
			}
			string[] array = (from x in _diaBanXa
				where x.MaTinh == keyTinhTP
				select x.TenXa).ToArray();
			if (array.Length > 0)
			{
				return new string[1] { "" }.Concat(array).ToArray();
			}
			return null;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> getDanhMucPhuongXaMoi1(string keyTinhTP)
	{
		Dictionary<string, string> result = new Dictionary<string, string>();
		try
		{
			if (_diaBanXa.Count == 0)
			{
				return null;
			}
			if (string.IsNullOrEmpty(keyTinhTP))
			{
				return _diaBanXa.ToDictionary((DiaBanXa x) => x.MaXa, (DiaBanXa x) => x.TenXa);
			}
			return _diaBanXa.Where((DiaBanXa x) => x.MaTinh == keyTinhTP).ToDictionary((DiaBanXa x) => x.MaXa, (DiaBanXa x) => x.TenXa);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return result;
		}
	}

	public static string[] getDanhMucCqtCapCucTheoDiaBanXa(string maXa)
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		if (string.IsNullOrEmpty(maXa))
		{
			return null;
		}
		CQTDiaBan cQTDiaBan = _cqtDiaBan.Find((CQTDiaBan x) => x.MaXa == maXa);
		if (cQTDiaBan != null)
		{
			string maCuc = cQTDiaBan.MaCQT.Substring(0, 3) + "00";
			CQTDiaBan cQTDiaBan2 = _cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == maCuc);
			dictionary.Add(maCuc, cQTDiaBan2.TenCQT);
			Dictionary<string, string> dictionary2 = _cqtDiaBan.Where((CQTDiaBan x) => x.MaCQT == "82500" || x.MaCQT == "82700").ToDictionary((CQTDiaBan x) => x.MaCQT, (CQTDiaBan x) => x.TenCQT);
			foreach (KeyValuePair<string, string> item in dictionary2)
			{
				dictionary.Add(item.Key, item.Value);
			}
		}
		return dictionary.Values.ToArray();
	}

	private List<CQTDiaBan> getDanhMucCQTDiaBan()
	{
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Expected O, but got Unknown
		//IL_0060: Unknown result type (might be due to invalid IL or missing references)
		//IL_0066: Expected O, but got Unknown
		_cqtDiaBan.Clear();
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_CQT_Dia_Ban.xml";
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
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					CQTDiaBan cQTDiaBan = new CQTDiaBan();
					cQTDiaBan.MaXa = array[0].Trim();
					cQTDiaBan.MaCQT = array[1].Trim();
					cQTDiaBan.TenCQT = array[2].Trim();
					_cqtDiaBan.Add(cQTDiaBan);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			_cqtDiaBan.Clear();
		}
		return _cqtDiaBan;
	}

	public static string[] getAllTextCqtCucMoi()
	{
		List<string> list = (from x in _cqtDiaBan
			where x.MaXa == "" && x.MaCQT.EndsWith("00")
			select x.TenCQT).ToList();
		if (!list.Contains("Chi cục Thuế doanh nghiệp lớn"))
		{
			list.Add("Chi cục Thuế doanh nghiệp lớn");
		}
		if (!list.Contains("Chi cục Thuế Thương mại điện tử"))
		{
			list.Add("Chi cục Thuế Thương mại điện tử");
		}
		return new string[1] { "" }.Concat(list).ToArray();
	}

	public static string getTenCqtMoiByMaCqt(string maCqt)
	{
		CQTDiaBan cQTDiaBan = _cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == maCqt);
		if (cQTDiaBan != null)
		{
			return cQTDiaBan.TenCQT;
		}
		return "";
	}

	public static string getMaCqtMoiByTenCqt(string tenCqt)
	{
		CQTDiaBan cQTDiaBan = _cqtDiaBan.Find((CQTDiaBan x) => x.TenCQT == tenCqt);
		if (cQTDiaBan != null)
		{
			return cQTDiaBan.MaCQT;
		}
		return "";
	}

	public static string[] getAllTextCqtChiCucMoi()
	{
		List<string> list = _cqtDiaBan.Select((CQTDiaBan x) => x.TenCQT).Distinct().ToList();
		if (!list.Contains("Chi cục Thuế doanh nghiệp lớn"))
		{
			list.Add("Chi cục Thuế doanh nghiệp lớn");
		}
		if (!list.Contains("Chi cục Thuế Thương mại điện tử"))
		{
			list.Add("Chi cục Thuế Thương mại điện tử");
		}
		return new string[1] { "" }.Concat(list).ToArray();
	}

	private void LoadDanhMucCoQuanThu()
	{
		try
		{
			coQuanThus.Clear();
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_CoQuanThu.xml";
			XDocument val = XDocument.Load(text);
			IEnumerable<XElement> source = ((XContainer)val).Descendants(XName.op_Implicit("Item"));
			coQuanThus = (from x in source
				select (string)x.Attribute(XName.op_Implicit("Value")) into v
				where !string.IsNullOrEmpty(v) && v.Contains("###")
				select v).Select(delegate(string v)
			{
				string[] array = v.Split(new string[1] { "###" }, StringSplitOptions.None);
				return new CoQuanThu
				{
					MaCQT = array[0].Trim(),
					Ma = array[1].Trim(),
					Ten = array[2].Trim()
				};
			}).ToList();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			coQuanThus.Clear();
		}
	}

	private void LoadDanhMucThueTinh()
	{
		try
		{
			dicThueTinh = CreateDictionaryDm(1, 0, "Catalogue_Thue_Tinh.xml");
			thueTinhs.Clear();
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Thue_Tinh.xml";
			XDocument val = XDocument.Load(text);
			IEnumerable<XElement> source = ((XContainer)val).Descendants(XName.op_Implicit("Item"));
			thueTinhs = (from x in source
				select (string)x.Attribute(XName.op_Implicit("Value")) into v
				where !string.IsNullOrEmpty(v) && v.Contains("###")
				select v).Select(delegate(string v)
			{
				string[] array = v.Split(new string[1] { "###" }, StringSplitOptions.None);
				return new ThueTinh
				{
					Ma = array[0].Trim(),
					Ten = array[1].Trim()
				};
			}).ToList();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			coQuanThus.Clear();
		}
	}

	public static Dictionary<string, string> GetDictDiaBanXaByMaCQT(string maCQT, List<DiaBanXa> danhSachXa, List<CQTDiaBan> danhSachCQT)
	{
		List<string> list = new List<string>();
		foreach (CQTDiaBan item in danhSachCQT)
		{
			if (item.MaCQT == maCQT && !list.Contains(item.MaXa))
			{
				list.Add(item.MaXa);
			}
		}
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (DiaBanXa item2 in danhSachXa)
		{
			if (list.Contains(item2.MaXa) && !dictionary.ContainsKey(item2.MaXa))
			{
				dictionary.Add(item2.MaXa, item2.TenXa);
			}
		}
		return dictionary;
	}

	public static string GetMaTinhMoiByMaTinhCu(string maTinhCu)
	{
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		//IL_0041: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Dia_Ban_Tinh_Map.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				XmlAttribute val3 = val2.Attributes["MaCu"];
				XmlAttribute val4 = val2.Attributes["MaMoi"];
				if (val3 == null || val4 == null)
				{
					continue;
				}
				string[] array = ((XmlNode)val3).Value.Split(new char[1] { ',' });
				string[] array2 = array;
				foreach (string text2 in array2)
				{
					if (maTinhCu.Equals(text2.Trim()))
					{
						return ((XmlNode)val4).Value;
					}
				}
			}
			return "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}
}
