using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using ADB.CopyDocument.Service.Models;

namespace ADB.CopyDocument.Service.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class CopyController : ControllerBase
    {
        private string[] ExcludeFields = new string[] {"ContentTypeId",
"_ModerationComments","FileLeafRef","Modified_x0020_By","Created_x0020_By","File_x0020_Type","HTML_x0020_File_x0020_Type",
"_SourceUrl","_SharedFileIndex","ComplianceAssetId","TemplateUrl","xd_ProgID","xd_Signature","_ShortcutUrl",
"_ShortcutSiteId","_ShortcutWebId","_ShortcutUniqueId","_ExtendedDescription","MediaServiceMetadata","MediaServiceFastMetadata",
"MediaServiceAutoTags","MediaServiceOCR","MediaServiceGenerationTime","MediaServiceEventHashCode","TaxCatchAll",
"ContentType","ID","_HasCopyDestinations","_CopySource","_ModerationStatus","FileRef","FileDirRef","Last_x0020_Modified",
"Created_x0020_Date","File_x0020_Size","FSObjType","SortBehavior","PermMask","PrincipalCount","CheckedOutUserId","IsCheckedoutToLocal",
"CheckoutUser","UniqueId","SyncClientId","ProgId","ScopeId","VirusStatus","CheckedOutTitle","_CheckinComment","LinkCheckedOutTitle",
"_EditMenuTableStart","_EditMenuTableStart2","_EditMenuTableEnd","LinkFilenameNoMenu","LinkFilename","LinkFilename2","DocIcon","ServerUrl",
"EncodedAbsUrl","BaseName","FileSizeDisplay","MetaInfo","_Level","_IsCurrentVersion","ItemChildCount","FolderChildCount","Restricted","OriginatorId",
"NoExecute","ContentVersion","_ComplianceFlags","_ComplianceTag","_ComplianceTagWrittenTime","_ComplianceTagUserId","_IsRecord",
"BSN","_ListSchemaVersion","_Dirty","_Parsable","_StubFile","_HasEncryptedContent","AccessPolicy","_VirusStatus","_VirusVendorID",
"_VirusInfo","_CommentFlags","_CommentCount","_LikeCount","_RmsTemplateId","_IpLabelId","_DisplayName","_IpLabelAssignmentMethod",
"A2ODMountCount","_ExpirationDate","AppAuthor","AppEditor","SMTotalSize","SMLastModifiedDate","SMTotalFileStreamSize","SMTotalFileCount",
"SelectTitle","SelectFilename","Edit","owshiddenversion","_UIVersion","_UIVersionString","InstanceID","Order","GUID","WorkflowVersion",
"WorkflowInstanceID","ParentVersionString","ParentLeafName","DocConcurrencyNumber","ParentUniqueId","StreamHash","Combine","RepairDocument",
"Created","Author","Modified","Editor","TaxCatchAllLabel", "b9798e09b2df41948d41147a2261e268",
"maebf37f863444c89c7b26c1d06434c8"};
        private readonly ILogger<CopyController> _logger;

        public CopyController(ILogger<CopyController> logger)
        {
            _logger = logger;
        }


        [HttpPost()]
        public bool Post([FromBody] Parameters parameters)
        {
            bool success = false;
            string fileName = string.Empty;
            ClientContext sourceContext = null;
            ClientContext destinationContext = null;
            string strNewUrl = string.Empty;
            Dictionary<string, object> fieldValuesForCopy = new Dictionary<string, object>();

            try
            {
                using (sourceContext = CommonUtility.GetClientContextWithAccessToken(parameters.SourceSiteUrl))
                {
                    File sourceFile = sourceContext.Web.GetFileByUrl(parameters.SourceFileUrl);


                    var data = sourceFile.OpenBinaryStream();
                    sourceContext.Load(sourceFile);
                    sourceContext.Load(sourceFile, s => s.Name, s => s.ListId);
                    sourceContext.ExecuteQuery();
                    if (sourceFile == null)
                        throw new Exception("*File Not Found !!!!");

                    strNewUrl = $"{parameters.DestinationFolder}/{(string.IsNullOrEmpty(parameters.DestinationFileName) ? sourceFile.Name : parameters.DestinationFileName)}";

                    if (parameters.SourceSiteUrl == parameters.DestinationSiteUrl)
                    {
                        if (parameters.IsMove)
                        {
                            sourceFile.MoveTo(strNewUrl, MoveOperations.Overwrite);
                        }
                        else
                        {
                            sourceFile.CopyTo(strNewUrl, true);
                        }
                        sourceContext.ExecuteQuery();
                        success = true;
                    }
                    else
                    {
                        System.IO.Stream sourceFileStream = data.Value;
                        using (destinationContext = CommonUtility.GetClientContextWithAccessToken(parameters.DestinationSiteUrl))
                        {
                            List destinationList = destinationContext.Web.GetListByTitle(parameters.DestinationLibrary);
                            destinationContext.Load(destinationList);
                            destinationContext.ExecuteQuery();

                            Folder destinationFolder = null;
                            if (parameters.DestinationFolder == "/" || string.IsNullOrEmpty(parameters.DestinationFolder))
                                destinationFolder = destinationList.RootFolder;
                            else
                                destinationFolder = destinationContext.Web.GetFolderByServerRelativeUrl($"{parameters.DestinationSiteUrl}/{parameters.DestinationLibrary}/{parameters.DestinationFolder}");
                            destinationContext.Load(destinationFolder, d => d.ServerRelativeUrl);
                            destinationContext.ExecuteQuery();

                            FieldCollection destinationFields = destinationList.Fields;
                            destinationContext.Load(destinationList);
                            destinationContext.Load(destinationFields);
                            destinationContext.ExecuteQuery();

                            #region Perform Copy Via Upload
                            fileName = string.IsNullOrEmpty(parameters.DestinationFileName) ? $"{destinationFolder.ServerRelativeUrl}/{sourceFile.Name}" : $"{destinationFolder.ServerRelativeUrl}/{parameters.DestinationFileName}";

                            File destinationFile = null;
                            destinationFile = UploadFileSlicePerSlice(destinationContext, destinationFolder, sourceFileStream, fileName, sourceFile.Length, 3);
                            if (destinationFile == null)
                            {
                                throw new ApplicationException("Could not upload the file.");
                            }
                            ListItem destinationListItem = destinationFile.ListItemAllFields;
                            destinationContext.Load(destinationFile, d => d.ListId);
                            destinationContext.Load(destinationListItem);
                            destinationContext.ExecuteQuery();
                            #endregion

                            #region Validate Source Fields and Destination Feilds Match
                            List sourceList = sourceContext.Web.GetListById(sourceFile.ListId);
                            ListItem sourceListItem = sourceFile.ListItemAllFields;
                            FieldCollection sourceFields = sourceList.Fields;
                            sourceContext.Load(sourceList);
                            sourceContext.Load(sourceListItem);
                            sourceContext.Load(sourceFields);
                            sourceContext.ExecuteQuery();
                            foreach (Field currentSourceField in sourceFields)
                            {
                                if (this.ExcludeFields.Contains(currentSourceField.InternalName))
                                {
                                    continue;
                                }
                                sourceContext.Load(currentSourceField);
                                sourceContext.ExecuteQuery();

                                if (destinationFields.FirstOrDefault(k => k.InternalName == currentSourceField.InternalName) != null)
                                {
                                    Field currentField = destinationFields.GetByInternalNameOrTitle(currentSourceField.InternalName);
                                    destinationContext.Load(currentField);
                                    destinationContext.ExecuteQuery();
                                    if (currentField.TypeAsString == "User")
                                    {
                                        FieldUserValue sourceUserValue = sourceListItem[currentSourceField.InternalName] as FieldUserValue;
                                        User user = destinationContext.Web.EnsureUser(sourceUserValue.Email);
                                        destinationContext.Load(user, u => u.Id);
                                        destinationContext.ExecuteQuery();
                                        FieldUserValue userValue = new FieldUserValue() { LookupId = user.Id };
                                        fieldValuesForCopy.Add(currentSourceField.InternalName, userValue);
                                    }
                                    else if (currentField.TypeAsString == "UserMulti")
                                    {
                                        List<FieldUserValue> users = new List<FieldUserValue>();
                                        FieldUserValue[] sourceUserValues = sourceListItem[currentSourceField.InternalName] as FieldUserValue[];

                                        foreach (FieldUserValue sourceFieldValue in sourceUserValues)
                                        {
                                            User user = destinationContext.Web.EnsureUser(sourceFieldValue.Email);
                                            destinationContext.Load(user, u => u.Id);
                                            destinationContext.ExecuteQuery();
                                            FieldUserValue userValue = new FieldUserValue() { LookupId = user.Id };
                                            users.Add(userValue);
                                        }
                                        fieldValuesForCopy.Add(currentSourceField.InternalName, users);
                                    }
                                    // SINCE OTHER FIELDS ARE COPIED AS IS ON UPLOAD, NO NEED TO COPY IT SEPARATELY
                                    /*
                                    else if (currentField.TypeAsString == "TaxonomyFieldType")
                                    {
                                        TaxonomyField taxField = destinationContext.CastTo<TaxonomyField>(currentField);
                                        TaxonomyFieldValue termValue = sourceListItem[currentSourceField.InternalName] as TaxonomyFieldValue;
                                        // taxField.SetFieldValueByValue(destinationListItem, termValue);
                                        fieldValuesForCopy.Add(currentSourceField.InternalName, $"{termValue.Label}|{termValue.TermGuid}");
                                    }
                                    else if (currentField.TypeAsString == "TaxonomyFieldTypeMulti")
                                    {
                                        TaxonomyField taxField = destinationContext.CastTo<TaxonomyField>(currentField);
                                        TaxonomyFieldValueCollection termValue = sourceListItem[currentSourceField.InternalName] as TaxonomyFieldValueCollection;
                                        // taxField.SetFieldValueByValueCollection(destinationListItem, termValue);
                                        fieldValuesForCopy.Add(currentSourceField.InternalName, string.Join(";", termValue.Select(k => $"{k.Label}|{k.TermGuid}")));
                                    }
                                    else
                                    {
                                        fieldValuesForCopy.Add(currentSourceField.InternalName, sourceListItem[currentSourceField.InternalName]);
                                    }
                                    */
                                }
                            }
                            #endregion

                            #region Create Dictionary of Fields and Values for updating the destination item metadata
                            if (parameters.MetadataForDestinationFile != null)
                            {
                                foreach (string key in parameters.MetadataForDestinationFile.Keys)
                                {
                                    if (destinationFields.First(k => k.InternalName == key) != null)
                                    {
                                        Field currentField = destinationFields.GetByInternalNameOrTitle(key);
                                        destinationContext.Load(currentField);
                                        destinationContext.ExecuteQuery();
                                        if (currentField.TypeAsString == "User")
                                        {
                                            User user = destinationContext.Web.EnsureUser(parameters.MetadataForDestinationFile[key]);
                                            destinationContext.Load(user, u => u.Id);
                                            destinationContext.ExecuteQuery();
                                            FieldUserValue userValue = new FieldUserValue() { LookupId = user.Id };
                                            if (fieldValuesForCopy.ContainsKey(key))
                                            {
                                                fieldValuesForCopy[key] = userValue;
                                            }
                                            else
                                            {
                                                fieldValuesForCopy.Add(key, userValue);
                                            }
                                        }
                                        else if (currentField.TypeAsString == "UserMulti")
                                        {
                                            List<FieldUserValue> users = new List<FieldUserValue>();
                                            foreach (string uv in parameters.MetadataForDestinationFile[key].ToString().Split(';', StringSplitOptions.RemoveEmptyEntries))
                                            {
                                                User user = destinationContext.Web.EnsureUser(uv);
                                                destinationContext.Load(user, u => u.Id);
                                                destinationContext.ExecuteQuery();
                                                FieldUserValue userValue = new FieldUserValue() { LookupId = user.Id };
                                                users.Add(userValue);
                                            }
                                            if (fieldValuesForCopy.ContainsKey(key))
                                            {
                                                fieldValuesForCopy[key] = users;
                                            }
                                            else
                                            {
                                                fieldValuesForCopy.Add(key, users);
                                            }
                                        }
                                        else if (currentField.TypeAsString == "TaxonomyFieldType")
                                        {
                                            TaxonomyField taxField = destinationContext.CastTo<TaxonomyField>(currentField);
                                            TaxonomyFieldValue termValue = new TaxonomyFieldValue();
                                            string[] term = parameters.MetadataForDestinationFile[key].ToString().Split('|');
                                            termValue.Label = term[0];
                                            termValue.TermGuid = term[1];
                                            termValue.WssId = -1;
                                            taxField.SetFieldValueByValue(destinationListItem, termValue);
                                            if (fieldValuesForCopy.ContainsKey(key))
                                            {
                                                fieldValuesForCopy[key] = termValue;
                                            }
                                            else
                                            {
                                                fieldValuesForCopy.Add(key, termValue);
                                            }
                                        }
                                        else if (currentField.TypeAsString == "TaxonomyFieldTypeMulti")
                                        {
                                            TaxonomyField taxField = destinationContext.CastTo<TaxonomyField>(currentField);
                                            TaxonomyFieldValueCollection termValue = new TaxonomyFieldValueCollection(destinationContext, parameters.MetadataForDestinationFile[key].ToString(), taxField);
                                            taxField.SetFieldValueByValueCollection(destinationListItem, termValue);
                                            if (fieldValuesForCopy.ContainsKey(key))
                                            {
                                                fieldValuesForCopy[key] = termValue;
                                            }
                                            else
                                            {
                                                fieldValuesForCopy.Add(key, termValue);
                                            }
                                        }
                                        else
                                        {
                                            if (fieldValuesForCopy.ContainsKey(key))
                                            {
                                                fieldValuesForCopy[key] = parameters.MetadataForDestinationFile[key].ToString();
                                            }
                                            else
                                            {
                                                fieldValuesForCopy.Add(key, parameters.MetadataForDestinationFile[key].ToString());
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region Update Item
                            foreach (string key in fieldValuesForCopy.Keys)
                            {
                                destinationListItem[key] = fieldValuesForCopy[key];
                            }
                            #endregion


                            destinationListItem.Update();
                            destinationContext.ExecuteQuery();

                            success = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(new EventId(1), ex, ex.Message, null);
                success = false;
            }

            return success;
        }

        private File UploadFileSlicePerSlice(ClientContext ctx, Folder destinationFolder, System.IO.Stream sourceFileStream, string fileName, long fileSize, int fileChunkSizeInMB = 3)
        {
            // Each sliced upload requires a unique ID.
            Guid uploadId = Guid.NewGuid();

            // Get the name of the file.
            string uniqueFileName = System.IO.Path.GetFileName(fileName);

            // File object.
            File uploadFile = null;

            // Calculate block size in bytes.
            int blockSize = fileChunkSizeInMB * 1024 * 1024;

            if (fileSize <= blockSize)
            {

                // Use regular approach.
                // using (System.IO.StreamWriter fs = new System.IO.StreamWriter(sourceFileStream))
                {
                    FileCreationInformation fileCreationInformation = new FileCreationInformation();
                    fileCreationInformation.ContentStream = sourceFileStream;
                    fileCreationInformation.Url = fileName;
                    fileCreationInformation.Overwrite = true;
                    uploadFile = destinationFolder.Files.Add(fileCreationInformation);
                    ctx.Load(uploadFile);
                    ctx.ExecuteQuery();
                    // Return the file object for the uploaded file.
                    return uploadFile;
                }
            }
            else
            {
                // Use large file upload approach.
                ClientResult<long> bytesUploaded = null;

                try
                {
                    using (System.IO.BinaryReader br = new System.IO.BinaryReader(sourceFileStream))
                    {
                        byte[] buffer = new byte[blockSize];
                        Byte[] lastBuffer = null;
                        long fileoffset = 0;
                        long totalBytesRead = 0;
                        int bytesRead;
                        bool first = true;
                        bool last = false;

                        // Read data from file system in blocks. 
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            totalBytesRead = totalBytesRead + bytesRead;

                            // You've reached the end of the file.
                            if (totalBytesRead == fileSize)
                            {
                                last = true;
                                // Copy to a new buffer that has the correct size.
                                lastBuffer = new byte[bytesRead];
                                Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                            }

                            if (first)
                            {
                                using (System.IO.MemoryStream contentStream = new System.IO.MemoryStream())
                                {
                                    // Add an empty file.
                                    FileCreationInformation fileInfo = new FileCreationInformation();
                                    fileInfo.ContentStream = contentStream;
                                    fileInfo.Url = uniqueFileName;
                                    fileInfo.Overwrite = true;
                                    uploadFile = destinationFolder.Files.Add(fileInfo);

                                    // Start upload by uploading the first slice. 
                                    using (System.IO.MemoryStream s = new System.IO.MemoryStream(buffer))
                                    {
                                        // Call the start upload method on the first slice.
                                        bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                        ctx.ExecuteQuery();
                                        // fileoffset is the pointer where the next slice will be added.
                                        fileoffset = bytesUploaded.Value;
                                    }

                                    // You can only start the upload once.
                                    first = false;
                                }
                            }
                            else
                            {
                                // Get a reference to your file.
                                uploadFile = ctx.Web.GetFileByServerRelativeUrl(fileName);

                                if (last)
                                {
                                    // Is this the last slice of data?
                                    using (System.IO.MemoryStream s = new System.IO.MemoryStream(lastBuffer))
                                    {
                                        // End sliced upload by calling FinishUpload.
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();

                                        // Return the file object for the uploaded file.
                                        // return uploadFile; // Commented to return after complete upload is done.
                                    }
                                }
                                else
                                {
                                    using (System.IO.MemoryStream s = new System.IO.MemoryStream(buffer))
                                    {
                                        // Continue sliced upload.
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();
                                        // Update fileoffset for the next slice.
                                        fileoffset = bytesUploaded.Value;
                                    }
                                }
                            }

                        } // while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)

                        uploadFile = ctx.Web.GetFileByServerRelativeUrl(fileName);
                    }
                }
                finally
                {

                }
            }

            return uploadFile;
        }

    }
}
