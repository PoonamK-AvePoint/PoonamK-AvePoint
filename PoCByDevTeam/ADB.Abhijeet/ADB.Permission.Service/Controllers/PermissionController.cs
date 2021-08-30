using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using ADB.Permission.Service.Models;


namespace ADB.Permission.Service.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class PermissionController : ControllerBase
    {   
        [HttpPost()]
        public string Post([FromBody] Parameters parameters)
        {
           string success = "";
           using(ClientContext context = CommonUtility.GetClientContextWithAccessToken(parameters.SiteUrl))
           {
                Web web = context.Web;
                
                context.Load(web);
                context.ExecuteQuery();
                Console.WriteLine(web.Title);
                success = web.Title;             
               
                List list= web.Lists.GetByTitle(parameters.ListName); 
                
                context.Load(list); 
                    context.Load(list.RootFolder,s=> s.ServerRelativeUrl);                  
                context.ExecuteQuery();
                var listrelativeUrl= list.RootFolder.ServerRelativeUrl;
                ListItem item = null;
                if(parameters.IsFile == true)
                {   
                        var documentUrl= $"{listrelativeUrl}/{parameters.ID}";
                        var file = web.GetFileByServerRelativeUrl(documentUrl);
                        item= file.ListItemAllFields;
                        context.Load(file);
                        context.Load(item);
                        context.ExecuteQuery();                        
                }
                else if(parameters.IsFile == false)
                {
                        item= list.GetItemById(parameters.ID);
                }                                 
                context.Load(item, i => i.HasUniqueRoleAssignments);          
                context.ExecuteQuery(); 
                if (item.HasUniqueRoleAssignments)
                {
                    item.ResetRoleInheritance();            
                    context.ExecuteQuery();
                }
                item.BreakRoleInheritance(false, false); 
                context.ExecuteQuery();  

                RoleAssignmentCollection roleAssignments= null;        
                roleAssignments = item.RoleAssignments; 
                
        
                RoleDefinitionBindingCollection roleDefCol = null;
                roleDefCol = new RoleDefinitionBindingCollection(context);
                roleDefCol.RemoveAll();

                foreach(UserAccess userAccessItem in parameters.accessList){
                    User user_group= null;
                    user_group = web.SiteUsers.GetByEmail(userAccessItem.User); 
                    switch (userAccessItem.PermissionLevel){
                        case "Read": 
                        roleDefCol.Add(web.RoleDefinitions.GetByType(RoleType.Reader));
                        break;
                        case "Edit": 
                        roleDefCol.Add(web.RoleDefinitions.GetByType(RoleType.Editor));
                        break;
                        case "Admin": 
                        roleDefCol.Add(web.RoleDefinitions.GetByType(RoleType.Administrator));
                        break;
                        case "View": 
                        roleDefCol.Add(web.RoleDefinitions.GetByType(RoleType.Guest));
                        break;
                        case "Contribute": 
                        roleDefCol.Add(web.RoleDefinitions.GetByType(RoleType.Contributor));
                        break;
                        case "Review": 
                        roleDefCol.Add(web.RoleDefinitions.GetByType(RoleType.Reviewer));
                        break;                           
                    }
                    roleAssignments.Add(user_group, roleDefCol);
                    context.Load(roleAssignments);
                }
                item.Update();                    
                context.ExecuteQuery();
                }
            return success;
        }
    }

    
}