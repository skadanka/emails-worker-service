using Microsoft.Office.Interop.Outlook;

public interface IMailItemWrapper
{
    string HTMLBody { get; }
    string EntryID { get; }
    DateTime CreationTime { get; }
    Attachments Attachments { get; }
}

public class MailItemWrapper : IMailItemWrapper
{
    private readonly MailItem _mailItem;

    public MailItemWrapper(MailItem mailItem)
    {
        _mailItem = mailItem;
    }

    public string HTMLBody => _mailItem.HTMLBody;
    public string EntryID => _mailItem.EntryID;
    public DateTime CreationTime => _mailItem.CreationTime;
    public Attachments Attachments => _mailItem.Attachments;

    public string Subject => _mailItem.Subject;
}
