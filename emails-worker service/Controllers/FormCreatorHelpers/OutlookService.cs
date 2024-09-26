using System;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

public class OutlookService : IEmailService, IDisposable
{
    private Application _outlookApp; // Store the application instance
    private NameSpace _outlookNamespace;
    private MAPIFolder _inboxFolder;
    private bool _disposed = false; // Track whether the object has been disposed

    public OutlookService()
    {
        // Initialize the Outlook application and namespace
        _outlookApp = new Application();
        _outlookNamespace = _outlookApp.GetNamespace("MAPI");
        _inboxFolder = _outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
    }

    // Get the Inbox folder
    public MAPIFolder GetInbox()
    {
        return _inboxFolder;
    }

    private void GetFolders(MAPIFolder folder)
    {
        if (folder.Folders.Count == 0)
        {
            Console.WriteLine(folder.FullFolderPath);
        }
        else
        {
            foreach (MAPIFolder subFolder in folder.Folders)
            {
                GetFolders(subFolder);
            }
        }
    }

        // Get a subfolder within the Inbox by name
        public MAPIFolder GetSubfolder(string folderName)
    {
        try
        {
            return _inboxFolder.Parent.Folders[folderName];
        }
        catch (System.Exception ex)
        {
            Console.WriteLine($"Error accessing subfolder '{folderName}': {ex.Message}");
            return null;
        }
    }

    // Get a specific folder by name from the root of the namespace or within the Inbox
    public MAPIFolder GetFolder(string folderName)
    {
        try
        {
            return _outlookNamespace.Folders[folderName];
        }
        catch (System.Exception ex)
        {
            Console.WriteLine($"Failed to retrieve folder: {folderName}. Error: {ex.Message}");
            return null;
        }
    }

    // Get a mail item by its EntryID
    public MailItem GetMailItemById(string entryId)
    {
        try
        {
            return _outlookNamespace.GetItemFromID(entryId) as MailItem;
        }
        catch (System.Exception ex)
        {
            Console.WriteLine($"Failed to retrieve mail item with EntryID {entryId}: {ex.Message}");
            return null;
        }
    }

    // Move a mail item to a designated folder
    public void MoveItemToFolder(MailItem mailItem, MAPIFolder targetFolder)
    {
        try
        {
            mailItem.Move(targetFolder);
        }
        catch (System.Exception ex)
        {
            Console.WriteLine($"Failed to move email '{mailItem.Subject}' to folder '{targetFolder.Name}'. Error: {ex.Message}");
        }
    }

    // Process all emails in a specified folder
   

    // Dispose method to release COM objects
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                // Managed resources clean-up (if any)
            }

            // Release the COM objects
            if (_outlookApp != null)
            {
                Marshal.ReleaseComObject(_outlookApp);
                _outlookApp = null;
            }

            if (_outlookNamespace != null)
            {
                Marshal.ReleaseComObject(_outlookNamespace);
                _outlookNamespace = null;
            }

            if (_inboxFolder != null)
            {
                Marshal.ReleaseComObject(_inboxFolder);
                _inboxFolder = null;
            }

            _disposed = true;
        }
    }

    // Finalizer in case Dispose was not called explicitly
    ~OutlookService()
    {
        Dispose(false);
    }
}
