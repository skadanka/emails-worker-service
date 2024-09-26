using Microsoft.Office.Interop.Outlook;

public interface IEmailService
{
    // Retrieves the Inbox folder
    MAPIFolder GetInbox();

    // Retrieves a subfolder within the Inbox by name
    MAPIFolder GetSubfolder(string folderName);

    // Retrieves a specific folder by name from the root of the namespace or within the Inbox
    MAPIFolder GetFolder(string folderName);

    // Retrieves a mail item by its EntryID
    MailItem GetMailItemById(string entryId);

    // Moves a mail item to a designated folder
    void MoveItemToFolder(MailItem mailItem, MAPIFolder targetFolder);


}
