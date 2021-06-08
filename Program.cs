using CommandLine;
using System;
using System.Security.Principal;

namespace KnockOutlook
{
	class Program
	{
		public static string Banner()
		{
			string banner = "\r\n";
			banner += @"      __ __                  __   ____        __  __            __  " + "\r\n";
			banner += @"     / //_/____  ____  _____/ /__/ __ \__  __/ /_/ /___  ____  / /__" + "\r\n";
			banner += @"    / ,<  / __ \/ __ \/ ___/ //_/ / / / / / / __/ / __ \/ __ \/ //_/" + "\r\n";
			banner += @"   / /| |/ / / / /_/ / /__/ ,< / /_/ / /_/ / /_/ / /_/ / /_/ / ,<   " + "\r\n";
			banner += @"  /_/ |_/_/ /_/\____/\___/_/\_\\____/\__,_/\__/_/\____/\____/_/\_\  " + "\r\n";
			banner += "\r\n\r\n";
			return banner;
		}

		private static void ShowHelp()
		{
			string help = Banner();
			help += "Parameters:\r\n";
			help += "    --operation :  specify the operation to run\r\n";
			help += "    --keyword   :  specify a keyword for the 'search' operation\r\n";
			help += "    --id        :  specify an EntryID for the 'save' operation\r\n";
			help += "    --bypass    :  bypass the Programmatic Access Security settings (requires admin)\r\n";
			help += "\r\nOperations:\r\n";
			help += "    check       :  perform a number of checks to ensure operational security\r\n";
			help += "    contacts    :  extract all contacts of every account\r\n";
			help += "    mails       :  extract mailbox metadata of every account\r\n";
			help += "    search      :  search for the provided keyword in every mailbox\r\n";
			help += "    save        :  save a specified mail by its EntryID\r\n";
			help += "\r\nExamples:\r\n";
			help += "    KnockOutlook.exe --operation check\r\n";
			help += "    KnockOutlook.exe --operation contacts\r\n";
			help += "    KnockOutlook.exe --operation mails --bypass\r\n";
			help += "    KnockOutlook.exe --operation search --keyword password\r\n";
			help += "    KnockOutlook.exe --operation save --id {EntryID} --bypass\r\n";
			Console.Write(help);
		}

		private static void ProcessArguments(Options options)
		{
			if (options.Bypass)
			{
				if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator))
				{
					string output = Banner();
					output += "The process is not running with high integrity level\r\n";
					Console.Write(output);
					Environment.Exit(0);
				}
			}

			if (options.Operation == "check")
			{
				Operations.Check.Run();
			}
			else if (options.Operation == "contacts")
			{
				Operations.Contacts.Run(options.Bypass);
			}
			else if (options.Operation == "mails")
			{
				Operations.Mails.Run(options.Bypass);
			}
			else if (options.Operation == "search")
			{
				if (string.IsNullOrEmpty(options.Keyword))
				{
					ShowHelp();
				}
				else
				{
					Operations.Search.Run(options.Keyword, options.Bypass);
				}
			}
			else if (options.Operation == "save")
			{
				if (string.IsNullOrEmpty(options.EntryID))
				{
					ShowHelp();
				}
				else
				{
					Operations.Save.Run(options.EntryID, options.Bypass);
				}
			}
			else
			{
				ShowHelp();
			}
		}

		static void Main(string[] args)
		{
			Console.OutputEncoding = System.Text.Encoding.UTF8;
			new Parser(config => config.HelpWriter = null).ParseArguments<Options>(args).WithParsed(options => ProcessArguments(options)).WithNotParsed(errors => ShowHelp());
		}
	}
}
