using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;

namespace ADB.CopyDocument.Service.Controllers
{
    [Route("api/[controller]")]
    public class FieldTypesController : Controller
    {
        private readonly ILogger<FieldTypesController> _logger;

        public FieldTypesController(ILogger<FieldTypesController> logger)
        {
            _logger = logger;
        }

        [HttpGet("{listName}")]
        public List<string> Get(string listName)
        {
            List<string> data = new List<string>();
            ClientContext context = CommonUtility.GetClientContextWithAccessToken("https://v2smartsolutions.sharepoint.com/sites/ADB");
            List documentsList = context.Web.GetListByName(listName);
            FieldCollection fields = documentsList.Fields;
            context.Load(documentsList);
            context.Load(fields);
            context.ExecuteQuery();

            foreach (Field field in fields)
            {
                context.Load(field);
                context.ExecuteQuery();
                data.Add($"FieldName:{field.InternalName}   |  FieldType:{field.TypeAsString}");
            }

            return data;
        }

        public IActionResult Index()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View("Error!");
        }
    }
}