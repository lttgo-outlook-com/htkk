using System;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

internal class BctcQd15Nckt200
{
	public bool ValidatePeriod_ALL(HTKKPeriodInfo PeriodInfo)
	{
		//IL_003b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0085: Unknown result type (might be due to invalid IL or missing references)
		bool result = true;
		try
		{
			if (int.Parse(PeriodInfo.Nam) > 2014)
			{
				frmMessage frmMessage = new frmMessage("0440", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return false;
			}
			if (int.Parse(PeriodInfo.Nam) > DateTime.Now.Year)
			{
				frmMessage frmMessage = new frmMessage("0063", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public bool ValidatePeriod_axfpsPhuLuc(FpSpread winFpSpread)
	{
		//IL_0095: Unknown result type (might be due to invalid IL or missing references)
		bool result = true;
		try
		{
			if ((bool)winFpSpread.Sheets[0].Cells[2, 0].Value == (bool)winFpSpread.Sheets[0].Cells[1, 0].Value && (bool)winFpSpread.Sheets[0].Cells[1, 0].Value)
			{
				frmMessage frmMessage = new frmMessage("0314", "2", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return frmMessage.Yes;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		return false;
	}
}
