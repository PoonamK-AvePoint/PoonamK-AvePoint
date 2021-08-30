using System.Collections.Generic;

namespace ADB.Permission.Service.Models
{
    public class Parameters
    {
        public string SiteUrl {get;set;}
       
        public string ListName {get;set;}
       
        public string ID {get;set;}

        public bool IsFile{get;set;}

       public List<UserAccess> accessList {get;set;}

    }
}