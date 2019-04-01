#r "Microsoft.SharePoint.Client.dll"  
#r "Microsoft.SharePoint.Client.Runtime.dll"  

using System;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.TenantAdministration;  
using Microsoft.Online.SharePoint.TenantManagement;

using System.Security.Authentication;

// For more information about this template visit http://aka.ms/azurebots-csharp-basic
[Serializable]
public class EchoDialog : IDialog<object>
{
    protected int count = 1;

    public Task StartAsync(IDialogContext context)
    {
        try
        {
            context.Wait(MessageReceivedAsync);
        }
        catch (OperationCanceledException error)
        {
            return Task.FromCanceled(error.CancellationToken);
        }
        catch (Exception error)
        {
            return Task.FromException(error);
        }

        return Task.CompletedTask;
    }

    public virtual async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> argument)
    {
        var message = await argument;
        if (message.Text == "reset")
        {
            PromptDialog.Confirm(
                context,
                AfterResetAsync,
                "desea ingresar otro incidente?",
                "no entiendo lo que dices!",
                promptStyle: PromptStyle.Auto);
        }
        else
        {
         if (this.count==1)
       {
         this.count++;
             // SHAREPOINT
         
     
await context.PostAsync($"Su mensaje: {message.Text}, ha sido trasladado, pronto nos comunicaremos con  usted.");
         
     using ClientContext ctx= new ClientContext("https://alcsa.sharepoint.com//sites///soportealcsa"); 
          
         //---------------------------------   
          //INSERTAR EN  LISTA DE CHERPOINT

                string login = "jsum@alcsa.com.gt"; //give your username here  
                string password = "alcsa1234"; //give your password  
                var securePassword = new SecureString();
                foreach (char c in password)
                {
                    securePassword.AppendChar(c);
                }

                string siteUrl = "https://alcsa.sharepoint.com/sites/soportealcsa";
                ClientContext clientContext = new ClientContext(siteUrl);

                Client.ListmyList = clientContext.Web.Lists.GetByTitle("Prueba Clavos");
                ListItem CreationInformationitemInfo = newListItemCreationInformation();
                ListItem myItem = myList.AddItem(itemInfo);
                myItem["Title"] = "Prueba: " + this.count;
                myItem["El clavo de los clavos"] = message.Text;

                myItem.Update();
                var onlineCredentials = new SharePointOnlineCredentials(login, securePassword);
                clientContext.Credentials = onlineCredentials;
                clientContext.ExecuteQuery();
         
         //-----------------------
         
         
         
        }
         else if (this.count > 1)
         {
         this.count++;
         await context.PostAsync($"Si desea agregar otro incidente escriba reset");
         }
        else
      //  {
     //      this.count++;
   //    await context.PostAsync($"Su mensaje: {message.Text}, ha sido trasladado, pronto nos comunicaremos con  usted.");
         
      //   }
            context.Wait(MessageReceivedAsync);
        }
    }

    public async Task AfterResetAsync(IDialogContext context, IAwaitable<bool> argument)
    {
        var confirm = await argument;
        if (confirm)
        {
            this.count = 1;
            await context.PostAsync("Puede crear una nueva solicitud");
        }
        else
        {
            await context.PostAsync("No se realizo el reinicio");
        }
        context.Wait(MessageReceivedAsync);
    }
}
