using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Management;
using System.Runtime.InteropServices;

namespace KnockOutlook.Operations
{
	class Check
	{
		private static string ExceptionMessage;
		private static string MajorVersion;
		private static string Architecture;
		private static bool ClickToRun;
		private static string ProgrammaticAccess;
		private static List<Dictionary<string, string>> AntivirusProducts;

		public static void Run()
		{
			if (CheckVersion())
			{
				if (CheckProgrammaticAccess())
				{
					CheckAntivirus();
				}
			}

			ShowResults();
		}

		private static void ShowResults()
		{
			string output = Program.Banner();

			if (string.IsNullOrEmpty(ExceptionMessage))
			{
				output += string.Format("Major Version : {0}\r\n", MajorVersion);
				output += string.Format("Architecture  : {0}\r\n", Architecture);
				output += string.Format("Click-to-Run  : {0}\r\n", ClickToRun);
				output += string.Format("\r\nProgrammatic Access Security is set to '{0}'\r\n", ProgrammaticAccess);

				if (ProgrammaticAccess.Contains("antivirus"))
				{
					if (AntivirusProducts.Count > 0)
					{
						output += string.Format("\r\nIdentified antivirus products:\r\n");
						int index = 0;

						foreach (Dictionary<string, string> antivirus in AntivirusProducts)
						{
							index++;
							output += string.Format("\r\n{0}. {1}\r\n", index, antivirus["DisplayName"]);
							output += string.Format(" {0}Engine State     : {1}\r\n", "\u251C\u2574", antivirus["EngineState"]);
							output += string.Format(" {0}Signature Status : {1}\r\n", "\u2514\u2574", antivirus["SignatureStatus"]);
						}
					}
				}
			}
			else
			{
				output += ExceptionMessage + "\r\n";
			}

			Console.Write(output);
		}

		private static bool CheckVersion()
		{
			try
			{
				Application Outlook = new Application();
				string productCode = Outlook.ProductCode;
				MajorVersion = Utilities.GetMajorVersion(productCode);
				Architecture = GetArchitecture(productCode);
				return true;
			}
			catch (COMException exception)
			{
				if (exception.HResult == -2147221164)
				{
					ExceptionMessage = "Outlook is not installed on the system";
				}
				else
				{
					ExceptionMessage = exception.ToString();
				}

				return false;
			}
		}

		private static bool CheckProgrammaticAccess()
		{
			RegistryKey outlookSecurity;

			if (Architecture == "x64")
			{
				if (Registry.LocalMachine.OpenSubKey(string.Format(@"SOFTWARE\Microsoft\Office\{0}\Common\InstallRoot\Virtual", MajorVersion)) == null)
				{
					ClickToRun = false;
					outlookSecurity = Registry.LocalMachine.OpenSubKey(string.Format(@"SOFTWARE\Microsoft\Office\{0}\Outlook\Security", MajorVersion));
				}
				else
				{
					ClickToRun = true;
					outlookSecurity = Registry.LocalMachine.OpenSubKey(string.Format(@"SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\{0}\Outlook\Security", MajorVersion));
				}
			}
			else
			{
				if (Registry.LocalMachine.OpenSubKey(string.Format(@"SOFTWARE\WOW6432Node\Microsoft\Office\{0}\Common\InstallRoot\Virtual", MajorVersion)) == null)
				{
					ClickToRun = false;
					outlookSecurity = Registry.LocalMachine.OpenSubKey(string.Format(@"SOFTWARE\WOW6432Node\Microsoft\Office\{0}\Outlook\Security", MajorVersion));
				}
				else
				{
					ClickToRun = true;
					outlookSecurity = Registry.LocalMachine.OpenSubKey(string.Format(@"SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\{0}\Outlook\Security", MajorVersion));
				}
			}

			int objectModelGuard = 0;

			if (outlookSecurity != null)
			{
				object value = outlookSecurity.GetValue("ObjectModelGuard");

				if (value != null)
				{
					objectModelGuard = Convert.ToInt32(value.ToString());
				}
			}

			if (objectModelGuard == 1)
			{
				ProgrammaticAccess = "Always warn";
				return false;
			}
			else if (objectModelGuard == 2)
			{
				ProgrammaticAccess = "Never warn";
				return false;
			}
			else
			{
				ProgrammaticAccess = "Warn when antivirus is inactive or out-of-date";
				return true;
			}
		}

		private static string GetArchitecture(string productCode)
		{
			return productCode.Split('-')[3][0] == '1' ? "x64" : "x86";
		}

		private static void CheckAntivirus()
		{
			AntivirusProducts = new List<Dictionary<string, string>>();
			ManagementObjectSearcher mos = new ManagementObjectSearcher(@"root\SecurityCenter2", "SELECT * FROM AntivirusProduct");

			foreach (ManagementObject mo in mos.Get())
			{
				Tuple<string, string> parsedProductState = ParseProductState(int.Parse(mo["productState"].ToString()));
				Dictionary<string, string> antivirus = new Dictionary<string, string>
				{
					{ "DisplayName", mo["displayName"].ToString() },
					{ "EngineState", parsedProductState.Item1 },
					{ "SignatureStatus", parsedProductState.Item2 }
				};

				AntivirusProducts.Add(antivirus);
			}
		}

		private static Tuple<string, string> ParseProductState(int productState)
		{
			return Tuple.Create(
				((EngineState)(productState & (int)ProductFlags.EngineState)).ToString(),
				((SignatureStatus)(productState & (int)ProductFlags.SignatureStatus)).ToString()
			);
		}

		[Flags]
		private enum ProductFlags
		{
			SignatureStatus = 0x00F0,
			EngineState = 0xF000
		}

		[Flags]
		private enum EngineState
		{
			Disabled = 0x0000,
			Enabled = 0x1000,
			Snoozed = 0x2000,
			Expired = 0x3000
		}

		[Flags]
		private enum SignatureStatus
		{
			Updated = 0x00,
			Outdated = 0x10
		}
	}
}
