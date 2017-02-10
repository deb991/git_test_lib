I thought I'd add something on navigating through folders too - this is all derived from the Microsoft documentation above, but might be helpful to have here, particularly if you're trying to go anywhere in the Outlook folder structure except the inbox.

You can navigate through the folders collection using folders - note in this case, there's no GetDefaultFolder after the GetNamespace (otherwise you'll likely end up with the inbox).

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace('MAPI')
folder = outlook.Folders[1]
The number is the index of the folder you want to access. To find out how many sub-folders are in there:

folder.Count
If there more sub-folders you can use another Folders to go deeper:

folder.Folders[2]
Folders returns a list of sub-folders, so to get the names of all the folders in the current directory, you can use a quick loop.

for i in range(folder.Count):
    print (folder[i].Name)
Each of the sub-folders has a .Items method to get a list of the emails.