using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace KnockOutlook.Operations
{
	class Mails
	{
		private static readonly List<Classes.Account> OutlookAccounts = new List<Classes.Account>();
		private static readonly List<Classes.Account> EmptyAccounts = new List<Classes.Account>();
		private static string ExportPath;
		private static string ExceptionMessage;

		public static void Run(bool bypass)
		{
			if (GetMails(bypass))
			{
				if (OutlookAccounts.Count > 0)
				{
					foreach (Classes.Account account in OutlookAccounts)
					{
						if (account.Mails.Count == 0)
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
							output += string.Format(" {0}Number of mails : {1}\r\n", "\u2514\u2574", account.Mails.Count);
						}
					}

					if (EmptyAccounts.Count > 0)
					{
						foreach (Classes.Account emptyAccount in EmptyAccounts)
						{
							index++;
							output += string.Format("\r\n{0}. {1}\r\n", index, emptyAccount.Address);
							output += string.Format(" {0}Number of mails : {1}\r\n", "\u2514\u2574", emptyAccount.Mails.Count);
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

		private static bool GetMails(bool bypass)
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
						List<Folder> mailFolders = Utilities.GetFolders((Folder)sessionAccount.DeliveryStore.GetRootFolder());

						if (mailFolders.Count > 0)
						{
							Classes.Account account = new Classes.Account(emailAddress)
							{
								Mails = new List<Classes.Mail>()
							};

							foreach (Folder folder in mailFolders)
							{
								foreach (object item in folder.Items)
								{
									try
									{
										MailItem mailItem = (MailItem)item;
										Classes.Mail Mail = new Classes.Mail(mailItem.EntryID, mailItem.ReceivedTime.ToString("dd MMMM yyyy - HH:mm:ss"), mailItem.Subject, GetSenderSMTPAddress(mailItem, emailAddress));

										foreach (Recipient recipient in mailItem.Recipients)
										{
											Mail.To.Add(recipient.PropertyAccessor.GetProperty(@"http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString());
										}

										if (mailItem.Attachments.Count > 0)
										{
											Mail.Attachments = new List<string>();

											foreach (Attachment attachment in mailItem.Attachments)
											{
												Mail.Attachments.Add(attachment.DisplayName);
											}
										}

										account.Mails.Add(Mail);
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

		private static string GetSenderSMTPAddress(MailItem mail, string currentUser)
		{
			if (mail.SenderEmailType == "EX")
			{
				AddressEntry sender = mail.Sender;

				if (sender != null)
				{
					if (sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry || sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
					{
						ExchangeUser exchangeUser = sender.GetExchangeUser();

						if (exchangeUser != null)
						{
							string primarySmtpAddress = exchangeUser.PrimarySmtpAddress;

							if (string.IsNullOrEmpty(primarySmtpAddress))
							{
								return currentUser;
							}
							else
							{
								return exchangeUser.PrimarySmtpAddress;
							}
						}
						else
						{
							return null;
						}
					}
					else
					{
						return sender.PropertyAccessor.GetProperty(@"http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString();
					}
				}
				else
				{
					return null;
				}
			}
			else
			{
				return mail.SenderEmailAddress;
			}
		}
	}
}
