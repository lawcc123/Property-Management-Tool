using System.IO;
using System.Windows;
using System.Windows.Controls;
using PropertyManager.App.Operations;

namespace PropertyManager.App.Security;

public partial class CredentialManagerWindow
{
    private void AddEmailSettings()
    {
        var tab=(TabItem)ServiceTabs.Items[2];
        var scroll=(ScrollViewer)tab.Content;
        var panel=(StackPanel)((Border)scroll.Content).Child;
        panel.Children.Add(new TextBlock{Text="Email delivery settings",FontSize=18,FontWeight=FontWeights.SemiBold,Margin=new Thickness(0,28,0,8)});
        panel.Children.Add(new TextBlock{Text="The sender account is shared by reminders and Personalized Email. Saving settings does not send an email. The password stays in Windows Credential Manager.",TextWrapping=TextWrapping.Wrap,Style=(Style)FindResource("Caption")});
        var config=OperationsSettings.Load();
        TextBox Field(string label,string value)
        {
            panel.Children.Add(new TextBlock{Text=label,Margin=new Thickness(0,14,0,6)});
            var input=new TextBox{Text=value,Padding=new Thickness(9)};panel.Children.Add(input);return input;
        }
        var server=Field("SMTP server (STARTTLS)",config.SmtpServer);
        var port=Field("Port",config.SmtpPort.ToString());
        panel.Children.Add(new TextBlock{Text="Use your email provider's SMTP settings. Port 587 is commonly used for STARTTLS. SSL-only port 465 is not supported.",TextWrapping=TextWrapping.Wrap,Style=(Style)FindResource("Caption"),Margin=new Thickness(0,6,0,0)});
        var from=Field("From email address",config.From);
        panel.Children.Add(new TextBlock{Text="Reminder recipients and subject",FontSize=16,FontWeight=FontWeights.SemiBold,Margin=new Thickness(0,24,0,0)});
        panel.Children.Add(new TextBlock{Text="These fields apply only to reminders. Personalized Email uses its own recipients and message.",TextWrapping=TextWrapping.Wrap,Style=(Style)FindResource("Caption")});
        var to=Field("Reminder To (separate email addresses with semicolons)",string.Join("; ",config.Recipients));
        var cc=Field("Cc (optional)",string.Join("; ",config.Cc));
        var subject=Field("Subject",config.Subject);
        var save=new Button{Content="Save email settings",HorizontalAlignment=HorizontalAlignment.Left,Margin=new Thickness(0,18,0,0)};
        var status=new TextBlock{TextWrapping=TextWrapping.Wrap,Style=(Style)FindResource("Caption"),Margin=new Thickness(0,10,0,0)};
        save.Click+=(_,_)=>
        {
            try
            {
                if(!int.TryParse(port.Text,out int number))throw new InvalidDataException("Enter a numeric SMTP port (1-65535).");
                // Reload so saving email fields preserves workbook and schedule settings.
                var current=OperationsSettings.Load();current.SmtpServer=server.Text.Trim();current.SmtpPort=number;current.From=from.Text.Trim();current.Recipients=OperationsSettings.Split(to.Text);current.Cc=OperationsSettings.Split(cc.Text);current.Subject=subject.Text.Trim();
                current.SaveEmailSettings(accountOnly:true);status.Text="Email account and reminder settings saved. Rebuild affected previews before sending. No email was sent.";
            }
            catch(Exception ex){MessageBox.Show(this,ex.Message,"Email sending settings",MessageBoxButton.OK,MessageBoxImage.Warning);}
        };
        panel.Children.Add(save);panel.Children.Add(status);
    }
}
