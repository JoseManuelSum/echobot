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
       
         if (message.Text == "Reset")
        {
            PromptDialog.Confirm(
                context,
                AfterResetAsync,
                "desea ingresar otro incidente?",
                "no entiendo lo que dices!",
promptStyle: PromptStyle.Auto);
         else
        {
       
        if (this.count==1)
       {
         this.count++;
             // SHAREPOINT
         
     
await context.PostAsync($"Su mensaje: {message.Text}, ha sido trasladado, pronto nos comunicaremos con  usted.");
         
         //   ClientContext ctx= new ClientContext("https://alcsa.sharepoint.com/sites/soportealcsa"); 
           // List announcementsList = ctx.Web.Lists.GetByTitle("Prueba Clavos"); 
            // We are just creating a regular list item, so we don't need to 
            // set any properties. If we wanted to create a new folder, for 
            // example, we would have to set properties such as 
            // UnderlyingObjectType to FileSystemObjectType.Folder. 
          //  ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation(); 
         //   ListItem newItem = announcementsList.AddItem(itemCreateInfo); 
         //   newItem["Title"] = "My New Item"; 
         //   newItem["El clavo de los clavos"] = message.Text; 
        //    newItem.Update(); 
//    ctx.ExecuteQuery();    
        
         
         //-----------------------
         
         
         
     }
         else if (this.count > 1)
         {
         this.count++;
         await context.PostAsync($"Si desea agregar otro incidente escriba reset");
         }
       // else
       // {
        //   this.count++;
      // await context.PostAsync($"Su mensaje: {message.Text}, ha sido trasladado, pronto nos comunicaremos con  usted.");
         
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
