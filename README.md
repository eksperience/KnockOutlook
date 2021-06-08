# KnockOutlook

*"The best feeling is to win by knockout." - Nonito Donaire*



## Overview

KnockOutlook is a C# project that interacts with Outlook's COM object in order to perform a number of operations useful in red team engagements.



## Command Line Usage

```
      __ __                  __   ____        __  __            __
     / //_/____  ____  _____/ /__/ __ \__  __/ /_/ /___  ____  / /__
    / ,<  / __ \/ __ \/ ___/ //_/ / / / / / / __/ / __ \/ __ \/ //_/
   / /| |/ / / / /_/ / /__/ ,< / /_/ / /_/ / /_/ / /_/ / /_/ / ,<
  /_/ |_/_/ /_/\____/\___/_/\_\\____/\__,_/\__/_/\____/\____/_/\_\


Parameters:
    --operation :  specify the operation to run
    --keyword   :  specify a keyword for the 'search' operation
    --id        :  specify an EntryID for the 'save' operation
    --bypass    :  bypass the Programmatic Access Security settings (requires admin)

Operations:
    check       :  perform a number of checks to ensure operational security
    contacts    :  extract all contacts of every account
    mails       :  extract mailbox metadata of every account
    search      :  search for the provided keyword in every mailbox
    save        :  save a specified mail by its EntryID

Examples:
    KnockOutlook.exe --operation check
    KnockOutlook.exe --operation contacts
    KnockOutlook.exe --operation mails --bypass
    KnockOutlook.exe --operation search --keyword password
    KnockOutlook.exe --operation save --id {EntryID} --bypass
```



## Operations

* **check**

  Enumerates the Outlook installation details in order to construct the correct registry key and retrieve the Programmatic Access Security setting.

  If this value is set to `Warn when antivirus is inactive or out-of-date` it queries WMI for any installed antivirus products and parses their current state.

  

* **contacts**

  Enumerates the contacts of every configured account and extracts the following information:

  * Full Name
  * Email Address

  

* **mails**

  Enumerates the mails of every configured account and extracts the following metadata:

  * ID
  * Timestamp
  * Subject
  * From
  * To
  * Attachments



* **search**

  Searches inside the mailbox of every configured account using Outlook's built-in search engine and returns the `EntryID` of mails that contain the provided keyword in their body.



* **save**

  Uses Outlook's built-in `Save As` mechanism to export a mail referenced by its `EntryID`.



## Object Model Guard Bypass

The `--bypass` switch can be used in conjunction with `contacts`, `mails`, `search` and `save` operations given the fact that the current process is running with high integrity level.

It will attempt to snapshot the current security policy of Outlook, patch it in a way that the Programmatic Access Security prompt is auto-allowed and finally revert it to its initial state after the operation has finished.



## Output

All operations will output basic information on screen.

The `contacts` and `mails` operations will output results in JSON format to a Gzip compressed file.

The `save` operation will export the requested mail in `.MSG` format.

All filenames are randomly generated during runtime.

By default, Outlook's Secure Temp Folder is used as a destination for all exported files.



## Authors

* [eks](https://twitter.com/eks_perience)
* [psof](https://github.com/psof)
