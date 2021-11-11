using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace KnockOutlook.Operations
{
	class Contacts
	{
		private static readonly List<Classes.Account> OutlookAccounts = new List<Classes.Account>();
		private static readonly List<Classes.Account> EmptyAccounts = new List<Classes.Account>();
		private static string ExportPath;
		private static string ExceptionMessage;

		public static void Run(bool bypass)
		{
			if (GetContacts(bypass))
			{
				if (OutlookAccounts.Count > 0)
				{
					foreach (Classes.Account account in OutlookAccounts)
					{
						if (account.Contacts.Count == 0)
						{
							EmptyAccounts.Add(account);
						}
					}

					foreach (Classes.Account emptyAccount in EmptyAccounts)
					{
						OutlookAccounts.Remove(emptyAccount);
					}

					if (OutlookAccounts.Count > 0)
					{
						Utilities.Export(OutlookAccounts, ExportPath);
					}
				}
			}

			ShowResults();
		}

		private static void ShowResults()
		{
			string output = Program.Banner();

			if (string.IsNullOrEmpty(ExceptionMessage))
			{
				if (OutlookAccounts.Count > 0 || EmptyAccounts.Count > 0)
				{
					int index = 0;
					output += "Identified accounts: \r\n";

					if (OutlookAccounts.Count > 0)
					{
						foreach (Classes.Account account in OutlookAccounts)
						{
							index++;
							output += string.Format("\r\n{0}. {1}\r\n", index, account.Address);
							output += string.Format(" {0}Number of contacts : {1}\r\n", "\u2514\u2574", account.Contacts.Count);
						}
					}

					if (EmptyAccounts.Count > 0)
					{
						foreach (Classes.Account emptyAccount in EmptyAccounts)
						{
							index++;
							output += string.Format("\r\n{0}. {1}\r\n", index, emptyAccount.Address);
							output += string.Format(" {0}Number of contacts : {1}\r\n", "\u2514\u2574", emptyAccount.Contacts.Count);
						}
					}

					if (!string.IsNullOrEmpty(ExportPath))
					{
						output += string.Format("\r\nExported at {0}\r\n", ExportPath);
					}
				}
				else
				{
					output += "No configured accounts were identified\r\n";
				}
			}
			else
			{
				output += ExceptionMessage + "\r\n";
			}

			Console.Write(output);
		}

		private static bool GetContacts(bool bypass)
		{
			try
			{
				Application Outlook = new Application();
				string majorVersion = Utilities.GetMajorVersion(Outlook.ProductCode);
				ExportPath = Utilities.GenerateExportPath(majorVersion);
				List<int> snapshot = null;

				if (bypass)
				{
					snapshot = Utilities.Bypass(majorVersion);
				}

				foreach (Account sessionAccount in Outlook.Session.Accounts)
				{
					string emailAddress = Utilities.GetAccountEmailAddress(sessionAccount);

					if (!string.IsNullOrEmpty(emailAddress))
					{
						List<Folder> contactFolders = Utilities.GetFolders((Folder)sessionAccount.DeliveryStore.GetRootFolder());

						if (contactFolders.Count > 0)
						{
							Classes.Account account = new Classes.Account(emailAddress)
							{
								Contacts = new List<Classes.Contact>()
							};

							foreach (Folder folder in contactFolders)
							{
								foreach (object item in folder.Items)
								{
									try
									{
										ContactItem contactItem = (ContactItem)item;
										Classes.Contact contact = new Classes.Contact(contactItem.FullName, contactItem.Email1Address);
										account.Contacts.Add(contact);
									}
									catch (InvalidCastException)
									{
										continue;
									}
								}
							}

							OutlookAccounts.Add(account);
						}
					}
				}

				if (bypass)
				{
					Utilities.Revert(majorVersion, snapshot);
				}

				return true;
			}
			catch (COMException exception)
			{
				if (exception.HResult == -2147221164)
				{
					ExceptionMessage = "Outlook is not installed on the system";
				}
				else if (exception.HResult == -2079195127)
				{
					ExceptionMessage = "Outlook is not connected";
				}
				else if (exception.HResult == -2147467260)
				{
					ExceptionMessage = "Operation aborted by the user";
				}
				else
				{
					ExceptionMessage = exception.ToString();
				}

				return false;
			}
		}
	}
}
