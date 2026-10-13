using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

namespace HTKK.Class;

public static class CertificateCallback
{
	static CertificateCallback()
	{
		ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
	}

	public static void Initialize()
	{
	}

	private static bool CertificateValidationCallBack(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
	{
		if (sslPolicyErrors == SslPolicyErrors.None)
		{
			return true;
		}
		if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) != SslPolicyErrors.None)
		{
			if (chain != null && chain.ChainStatus != null)
			{
				X509ChainStatus[] chainStatus = chain.ChainStatus;
				for (int i = 0; i < chainStatus.Length; i++)
				{
					X509ChainStatus x509ChainStatus = chainStatus[i];
					if ((!(certificate.Subject == certificate.Issuer) || x509ChainStatus.Status != X509ChainStatusFlags.UntrustedRoot) && x509ChainStatus.Status != X509ChainStatusFlags.NoError)
					{
						return false;
					}
				}
			}
			return true;
		}
		return false;
	}
}
