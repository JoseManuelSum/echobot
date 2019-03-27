using System;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Microsoft.SharePoint.Client;

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
                "Are you sure you want to reset the count?",
                "Didn't get that!",
                promptStyle: PromptStyle.Auto);
        }
        else
        {
         if (this.count==2)
         {
         this.count++;
             // SHAREPOINT
         await context.PostAsync($"Su mensaje: {message.Text}, ha sido trasladado, pronto nos comunicaremos con  usted.");
  
            ClientContext contextSP = new ClientContext("https://alcsa.sharepoint.com/sites/soportealcsa"); 

            // Assume that the web has a list named "Announcements". 
            List announcementsList = contextSP.Web.Lists.GetByTitle("Prueba Clavos"); 
            // We are just creating a regular list item, so we don't need to 
            // set any properties. If we wanted to create a new folder, for 
            // example, we would have to set properties such as 
            // UnderlyingObjectType to FileSystemObjectType.Folder. 
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation(); 
            ListItem newItem = announcementsList.AddItem(itemCreateInfo); 
            newItem["Title"] = "My New Item"; 
            newItem["El clavo de los clavos"] = message.Text; 
            newItem.Update(); 

            contextSP.ExecuteQuery();    
         
         
         //-----------------------
         
         
         
         }
         else if (this.count > 2)
         {
         this.count++;
         await context.PostAsync($"Si desea agregar otro incidente escriba RESET");
         }
         else
         {
           this.count++;
         await context.PostAsync($"Por favor escriba su requerimiento en un solo mensaje.");
         
         }
            context.Wait(MessageReceivedAsync);
        }
    }

    public async Task AfterResetAsync(IDialogContext context, IAwaitable<bool> argument)
    {
        var confirm = await argument;
        if (confirm)
        {
            this.count = 1;
            await context.PostAsync("Reset count.");
        }
        else
        {
            await context.PostAsync("Did not reset count.");
        }
        context.Wait(MessageReceivedAsync);
    }
}
