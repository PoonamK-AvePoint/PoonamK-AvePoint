using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using ADB.LogError.Service.Models;


namespace ADB.LogError.Service.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class LogError : ControllerBase
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
                string listName = "LogError"; 
                bool listExist= web.ListExists(listName);
                
                if(!listExist)
                {
                    ListCreationInformation creationInfo = new ListCreationInformation(); 
                    creationInfo.Title = listName;  
                    creationInfo.Description = "Log Error list";  
                    creationInfo.TemplateType = (int) ListTemplateType.CustomGrid;
                    List logList = context.Web.Lists.Add(creationInfo);                     
                    context.Load(logList);               
                    context.ExecuteQuery(); 
                    Field msgField = logList.Fields.AddFieldAsXml("<Field DisplayName=\'Message\' Type=\'Text\' />", true, AddFieldOptions.DefaultValue);
                    Field desscField = logList.Fields.AddFieldAsXml("<Field DisplayName=\'Description\' Type=\'Note\' NumLines=\'99\' Name=\'Description\'/>",true,AddFieldOptions.DefaultValue);
                    msgField.Update();
                    desscField.Update();
                    context.ExecuteQuery(); 
                }
                List list= web.Lists.GetByTitle(listName);
                context.Load(list); 
                context.ExecuteQuery();
                ListItemCreationInformation oListItemCreationInformation = new ListItemCreationInformation();
                ListItem oItem = list.AddItem(oListItemCreationInformation);

                foreach(string key in parameters.ErrorLogDetails.Keys)
                {
                        oItem[key] = parameters.ErrorLogDetails[key].ToString();                              
                }
                oItem.Update();             
                context.ExecuteQuery();
           }
            return success;
        }
    }

    
}