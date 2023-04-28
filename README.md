
# Outlook-Automation-With-Python

If you work with Microsoft Outlook and are looking for ways to automate common email-related tasks, the pypiwin32 Python library is a great tool to have in your arsenal. You can easily access the Outlook object model and write Python code to create, read, update, and delete emails, contacts appointments, download email attachments or send automated reports to your contacts.

This post is focused for people who have a Microsoft 365 License (provided by your company or university - if you have a license you can use multiple accounts included personal accounts)


![outlook image (5)](https://user-images.githubusercontent.com/74120313/235248173-16484c8c-bd38-48d6-94bb-e49b006aae56.png)

For more specific information about pypiwin32 check the following link: https://pypi.org/project/pywin32/
To start using this great tool you have to install the requirements shown below: 
* Outlook Desktop Application
* pywin32==306
* pypiwin32==223

In this post we will cover two main sections: how to download and send emails by customizing the principal parameters.

### Download Files:

#### **From your inbox:** 
It is important to mentionate that you have to specify the location where emails are stored.

```python
# Make a connection with your outlook account

# We import the main library
import win32com.client as win32

outlook = win32.Dispatch('Outlook.Application')

# Accesss to the Account:
namespace = outlook.GetNamespace("MAPI")

# Provide your email account
email_account = "example@outlook.com"

# We check the access:
for account in namespace.Accounts:
 if account.DisplayName == email_account:
  break

#We access into the inbox-mail
inbox = account.DeliveryStore.GetDefaultFolder(6)

#Get all emails in the inbox section:
emails = inbox.Items
```



You can download an attachment from any email or parts from it by some conditionals:
It is a best practice to keep your mailbox clean by deleting emails you no longer need and create different folders to store emails by categories. Before we start looking for emails, we can sort them by date, either from newest to oldest or vice versa.


**From newest to oldest**
```python
emails = inbox.Items
emails.Sort("[ReceivedTime]", True)
```

**From oldest to newest**
```python
emails = inbox.Items
emails.Sort("[ReceivedTime]", False)
```
Now we identified the emails we want to manage. 

#### **Received Time**
```python
from_date = datetime(1985,4,23).date()
to_date = datetime(1985,4,25).date()

for email in emails:
    email_date = datetime.strptime(str(email.ReceivedTime).split()[0], "%Y-%m-%d").date()
    if ( from_date <= email_date <= to_date):
    	#Code

```

#### **Specific Sender**

```python
name_sender = "Barry Allen".upper()
email_sender = "Barry.Allen@justiceleague.com"

for email in emails:
    if (
        email.SenderEmailAddress == email_sender
        or email.SenderName == name_sender
    ):
		#Code
```

#### **Matching Subjects**
```python
#Library to search for a part of the string in other string
import re

#We look for emails that contains "Process 1" in the subject
match_key = "Process 1".upper()

#Matching the subjects:
for email in emails:
	subject = email.Subject
	match = re.search(fr"\b{match_key}\b", paragraph)

	if match:
		#Code
```

#### **Matching Bodies**

```python
#Library to search for a part of the string in other string
import re

#We look for emails that contains "Process 1" in the body
match_key = "Process 1".upper()

#Matching the subjects:
for email in emails:
	subject = email.Body
	match = re.search(fr"\b{match_key}\b", paragraph)

	if match:
		#Code
```

If the if statement is **True** we code below the conditional could be the following. 

```python
#Code
if any_statement:

#Looking for any attachments
	for attachment in email.Attachments:
		attachment.SaveAsFile(f"{directory}//{attachment.FileName}")
		#Print the basic info of each emails:
		 print(f" * Subject: {email.Subject}",
				  f"- Received on: {str(email.ReceivedTime).split()[0]}",
				  f"- Attachment: {attachment.FileName}")
else:
	continue
```

#### **From a folder:**
This part is fundamental , because you can create folders to store emails by applying some rules: https://support.microsoft.com/en-us/office/set-up-rules-in-outlook-75ab719a-2ce8-49a7-a214-6d62b67cbd41

**Folder created in the main section**
There are a lot of ways to get a folder located in the main section this is one of them.

```python
#Import Libraries
import win32com.client as win32

outlook = win32.Dispatch('Outlook.Application')

# access the into the folder
email_account = "example@outlook.com"
folder_name = "Folder Example"
folder = outlook.GetNamespace("MAPI").Folders(email_account).Folders(folder_name)

# Get the emails from the folder
emails = folder.Items
```

**Folder created as subfolder in your inbox section**


```python
#Import Libraries
import win32com.client as win32

outlook = win32.Dispatch('Outlook.Application')

# Access the Account
namespace = outlook.GetNamespace("MAPI")

for account in namespace.Accounts:
  if account.DisplayName == email_address:
    break
    
print(f"\nSigning in with: {account}")

#Outlook parameters
email_account = "example@outlook.com"
folder_name = "Folder Example"
    
#Get the folder:
inbox = account.DeliveryStore.GetDefaultFolder(6)
folder = inbox.Folders[folder_name]
```

## Send Emails

You can specify recipients, set the subject and body of the email, attach files, and send the email directly through Outlook. 

```python
import win32com.client as win32

outlook = win32.Dispatch('Outlook.Application')
mail = outlook.CreateItem(0)

recipients = ["example_1@emial.com","example_2@email.com"]
 
for recipient in recipients:
    mail.Recipients.Add(recipient)

mail.Subject = 'Hello from Python using pypiwin32'
mail.Body = """
If one of your principal taskes is sending emails.
stop doing that repetitive job and Join to Python Community
You can automate that process using this great tool.
             """
attachment_path = r'path/to/attachment.txt'
mail.Attachments.Add(attachment_path)
mail.Send()
```

_More functionalities will coming soon, stay tuned. Feel free to leave a comment or a contribution to improve this project_

