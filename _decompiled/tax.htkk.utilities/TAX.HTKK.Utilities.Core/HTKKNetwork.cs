using System;
using System.Collections.Generic;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;

namespace TAX.HTKK.Utilities.Core;

public class HTKKNetwork
{
	public static string[] GetPhysicalAddresses()
	{
		try
		{
			List<string> list = new List<string>();
			NetworkInterface[] allNetworkInterfaces = NetworkInterface.GetAllNetworkInterfaces();
			NetworkInterface[] array = allNetworkInterfaces;
			foreach (NetworkInterface networkInterface in array)
			{
				PhysicalAddress physicalAddress = networkInterface.GetPhysicalAddress();
				byte[] addressBytes = physicalAddress.GetAddressBytes();
				if (addressBytes.Length == 0)
				{
					continue;
				}
				StringBuilder stringBuilder = new StringBuilder();
				for (int j = 0; j < addressBytes.Length; j++)
				{
					stringBuilder.Append(addressBytes[j].ToString("X2"));
					if (j != addressBytes.Length - 1)
					{
						stringBuilder.Append("-");
					}
				}
				list.Add(stringBuilder.ToString());
			}
			return list.ToArray();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static string[] GetListMacMac()
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		//IL_002b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0031: Expected O, but got Unknown
		//IL_0033: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_005d: Expected O, but got Unknown
		try
		{
			List<string> list = new List<string>();
			ManagementScope val = new ManagementScope("\\\\computerName\\root\\cimv2");
			StringBuilder stringBuilder = new StringBuilder();
			stringBuilder.Append("SELECT MACAddress FROM Win32_NetworkAdapter");
			ObjectQuery val2 = new ObjectQuery(stringBuilder.ToString());
			ManagementObjectSearcher val3 = new ManagementObjectSearcher(val, val2);
			ManagementObjectCollection val4 = val3.Get();
			ManagementObjectEnumerator enumerator = val4.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					ManagementObject val5 = (ManagementObject)enumerator.Current;
					string item = "MAC Address: " + ((ManagementBaseObject)val5)["MACAddress"].ToString();
					list.Add(item);
				}
			}
			finally
			{
				((IDisposable)enumerator)?.Dispose();
			}
			return list.ToArray();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static string GetComputerLanIP()
	{
		string hostName = Dns.GetHostName();
		IPHostEntry hostEntry = Dns.GetHostEntry(hostName);
		try
		{
			IPAddress[] addressList = hostEntry.AddressList;
			foreach (IPAddress iPAddress in addressList)
			{
				if (iPAddress.AddressFamily.ToString() == "InterNetwork")
				{
					return iPAddress.ToString();
				}
			}
			return "-1";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static bool PingHost(string nameOrAddress)
	{
		bool result = false;
		Ping ping = new Ping();
		try
		{
			PingReply pingReply = ping.Send(nameOrAddress);
			result = pingReply.Status == IPStatus.Success;
		}
		catch (PingException)
		{
		}
		return result;
	}

	public static bool ConnectToInternet(int timeout_per_host_millis)
	{
		if (NetworkInterface.GetIsNetworkAvailable())
		{
			string[] array = new string[2] { "www.google.com", "www.facebook.com" };
			Ping ping = new Ping();
			string[] array2 = array;
			foreach (string hostNameOrAddress in array2)
			{
				try
				{
					PingReply pingReply = ping.Send(hostNameOrAddress, timeout_per_host_millis);
					if (pingReply.Status == IPStatus.Success)
					{
						return true;
					}
				}
				catch
				{
				}
			}
		}
		return false;
	}
}
