using System;
using System.Collections;
using System.Configuration.Install;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ServiceProcess;

namespace TAX.HTKK.Utilities.Core;

public class HTKKWindowService
{
	private string _servicePath;

	private string _serviceName;

	private string _serviceDisplayName;

	[DllImport("advapi32.dll")]
	public static extern IntPtr OpenSCManager(string lpMachineName, string lpSCDB, int scParameter);

	[DllImport("Advapi32.dll")]
	public static extern IntPtr CreateService(IntPtr SC_HANDLE, string lpSvcName, string lpDisplayName, int dwDesiredAccess, int dwServiceType, int dwStartType, int dwErrorControl, string lpPathName, string lpLoadOrderGroup, int lpdwTagId, string lpDependencies, string lpServiceStartName, string lpPassword);

	[DllImport("advapi32.dll")]
	public static extern void CloseServiceHandle(IntPtr SCHANDLE);

	[DllImport("advapi32.dll")]
	public static extern int StartService(IntPtr SVHANDLE, int dwNumServiceArgs, string lpServiceArgVectors);

	[DllImport("advapi32.dll", SetLastError = true)]
	public static extern IntPtr OpenService(IntPtr SCHANDLE, string lpSvcName, int dwNumServiceArgs);

	[DllImport("advapi32.dll")]
	public static extern int DeleteService(IntPtr SVHANDLE);

	[DllImport("kernel32.dll")]
	public static extern int GetLastError();

	public bool InstallService(string svcPath, string svcName, string svcDispName)
	{
		int scParameter = 2;
		int dwServiceType = 16;
		int dwErrorControl = 1;
		int num = 983040;
		int num2 = 1;
		int num3 = 2;
		int num4 = 4;
		int num5 = 8;
		int num6 = 16;
		int num7 = 32;
		int num8 = 64;
		int num9 = 128;
		int num10 = 256;
		int dwDesiredAccess = num | num2 | num3 | num4 | num5 | num6 | num7 | num8 | num9 | num10;
		int dwStartType = 2;
		try
		{
			IntPtr intPtr = OpenSCManager(null, null, scParameter);
			if (intPtr.ToInt32() != 0)
			{
				IntPtr sVHANDLE = CreateService(intPtr, svcName, svcDispName, dwDesiredAccess, dwServiceType, dwStartType, dwErrorControl, svcPath, null, 0, null, null, null);
				if (sVHANDLE.ToInt32() == 0)
				{
					CloseServiceHandle(intPtr);
					return false;
				}
				if (StartService(sVHANDLE, 0, null) == 0)
				{
					return false;
				}
				CloseServiceHandle(intPtr);
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	public bool UnInstallService(string svcName)
	{
		int scParameter = 1073741824;
		IntPtr sCHANDLE = OpenSCManager(null, null, scParameter);
		try
		{
			if (sCHANDLE.ToInt32() != 0)
			{
				int dwNumServiceArgs = 65536;
				IntPtr sVHANDLE = OpenService(sCHANDLE, svcName, dwNumServiceArgs);
				if (sVHANDLE.ToInt32() != 0)
				{
					if (DeleteService(sVHANDLE) != 0)
					{
						CloseServiceHandle(sCHANDLE);
						return true;
					}
					CloseServiceHandle(sCHANDLE);
					return false;
				}
				return false;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public void Uninstall(string ServiceName)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		ServiceInstaller val = new ServiceInstaller();
		InstallContext context = new InstallContext("c:\\install.log", (string[])null);
		((Installer)val).Context = context;
		val.ServiceName = $"{ServiceName}";
		((Installer)val).Uninstall((IDictionary)null);
	}

	public static void StartService(string serviceName, int timeoutMilliseconds)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		//IL_000a: Unknown result type (might be due to invalid IL or missing references)
		ServiceController val = new ServiceController(serviceName);
		try
		{
			if (((object)val.Status).Equals((object)(ServiceControllerStatus)1))
			{
				TimeSpan timeSpan = TimeSpan.FromMilliseconds(timeoutMilliseconds);
				val.Start();
				val.WaitForStatus((ServiceControllerStatus)4, timeSpan);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void StopService(string serviceName, int timeoutMilliseconds)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		//IL_000a: Unknown result type (might be due to invalid IL or missing references)
		ServiceController val = new ServiceController(serviceName);
		try
		{
			if (((object)val.Status).Equals((object)(ServiceControllerStatus)4))
			{
				TimeSpan timeSpan = TimeSpan.FromMilliseconds(timeoutMilliseconds);
				val.Stop();
				val.WaitForStatus((ServiceControllerStatus)1, timeSpan);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool IsServiceInstalled(string serviceName)
	{
		bool result = false;
		try
		{
			ServiceController[] services = ServiceController.GetServices();
			ServiceController[] array = services;
			foreach (ServiceController val in array)
			{
				if (val.ServiceName == serviceName)
				{
					result = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static bool IsServiceRuning(string serviceName)
	{
		//IL_0005: Unknown result type (might be due to invalid IL or missing references)
		//IL_000b: Expected O, but got Unknown
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		bool result = false;
		try
		{
			ServiceController val = new ServiceController(serviceName);
			if (((object)val.Status).Equals((object)(ServiceControllerStatus)4))
			{
				result = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}
}
