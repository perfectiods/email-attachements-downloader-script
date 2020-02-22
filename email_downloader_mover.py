#####
#
# Step 1. Create directory for yesterday's files and move them into it
# Step 2. Connect to MS Outlook, scan letters in sub folder, and if find letter with predefined theme,
# download all attachments to folder /get_path
#
####


import win32com.client, os, shutil,time, datetime
# directory with files
dir  = "C:\\Users\\Joe\\files”
today = datetime.date.today()
todaystr = today.isoformat()

# Step 1. Create directory for yesterday's files and move them into it
dir_name = todaystr
print (f'dir_name: {dir_name}')
# create path to this directory by adding to dir's name today date
dir_path = dir + dir_name
print (f'dir_path: {dir_path}')
# scan directory in cycle
For file in os.listdir(dir):
    #check if such dir exists and create it if no
    If not os.path.exists(dir_path):
         Os.makedrs(dir_path)
    #add filename to path
    If os.path.exists(dir_path):
        file_path = dir + file
        Print (f'file_path: {file_path}')
         # move only files, eliminate dirs
         For fname in os.listdir(dir):
              Path = os.path.join(dir, fname)
              Print(f'path :{path}')
               If os.path.isdir(path):
                     Continue
                # move files into created folder.
                Shutil.move(path, dir_path) #(path+file_name, path where)


# Step 2. Connect to MS Outlook, scan letters in sub folder, and if find letter with predefined theme,
# download all attachments to folder /get_path
Outlook = win32.com.client.Dispatch(”Outlook.Application”).GetNamespace(”MAPI”)
#inbox = outlook.GetDefaultFolder(6) where 6 refers to the index of filder (inbox)
My_folder = outlook.Folders ['test@test.com'].Folders['inbox'].Folders['subfolder']
messages = my_folder.Items
message = messages.GetFitst ()
subject = message.Subject

# define directory where put attachments
get_path = "C:\\Users\\Joe\\files”

For m in messages:
   If m.Subject == ”Theme name”:
      Print (message)
      Attachments = message.attachements
      Num_attach = len(([x for x in attachments]))
      For x in range(1,num_attach + 1):
           Attachment = attachments.Item(x)
           Attachment.SaveAsFile(os.path.join(get_path, attachment.FileName))
      Print(attachment)
     message = messages.GetNext()
   Else:
      message = messages.GetNext()
Print('ok')