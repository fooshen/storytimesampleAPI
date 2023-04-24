using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Extensions;
using System;
using System.Diagnostics;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Linq;
using System.IO;
using System.Text;

namespace RTStoryTime.RTStoryTimeSlides
{
    /// <summary>
    /// Plugin development guide: https://docs.microsoft.com/powerapps/developer/common-data-service/plug-ins
    /// Best practices and guidance: https://docs.microsoft.com/powerapps/developer/common-data-service/best-practices/business-logic/
    /// </summary>
    public class MergeSlides : PluginBase
    {
        public MergeSlides(string unsecureConfiguration, string secureConfiguration) : base(typeof(MergeSlides))
        {
            // TODO: Implement your custom configuration handling
            // https://docs.microsoft.com/powerapps/developer/common-data-service/register-plug-in#set-configuration-data
        }

        // Entry point for custom business logic execution
        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var context = localPluginContext.PluginExecutionContext;
            ITracingService tracingSvc = localPluginContext.TracingService;

            tracingSvc.Trace("MergeSlides running");
            // Check for the entity on which the plugin would be registered
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
            {
                
                var entityRef = (EntityReference)context.InputParameters["Target"];
          
                // Check for entity name on which this plugin would be registered
                if (entityRef.LogicalName == "rio_deck")
                {
                    var deckEntity = localPluginContext.CurrentUserService.Retrieve(entityRef.LogicalName, entityRef.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("rio_templates"));
                    tracingSvc.Trace($"Target: {deckEntity.LogicalName}");

                    string templateJSON = deckEntity.Attributes["rio_templates"] as String;
                    tracingSvc.Trace($"Templates value: { templateJSON}");

                    List<Template> templates = JsonConvert.DeserializeObject<List<Template>>(templateJSON);
                    if(templates.Count == 2)
                    {
                        tracingSvc.Trace($"Source ID: {templates[0].Id}, Name: {templates[0].Name}");

                        byte[] sourceFileContent = GetStoryTimeTemplateFile(localPluginContext.CurrentUserService, templates[0].Id);
                        var sourceDoc = PresentationDocument.Open(new MemoryStream(sourceFileContent), false); 
                        var sourcePresentationPart = sourceDoc.PresentationPart;
                        var sourcePresentation = sourcePresentationPart.Presentation;

                        tracingSvc.Trace($"Destination File ID: {templates[1].Id}, Name: {templates[1].Name}");

                        
                        byte[] outputFileContent = GetStoryTimeTemplateFile(localPluginContext.CurrentUserService, templates[1].Id);
                        MemoryStream outputFile = new MemoryStream();
                        outputFile.Write(outputFileContent, 0, outputFileContent.Length);
                        tracingSvc.Trace("Acquired destination file");

                        var destDoc = PresentationDocument.Open(outputFile, true);
                        tracingSvc.Trace("Presentation doc created");
                        var destPresentationPart = destDoc.PresentationPart;
                        var destPresentation = destPresentationPart.Presentation;
                        
                        int countSlidesInSourcePresentation = sourcePresentation.SlideIdList.Count();
                        tracingSvc.Trace($"Source slides count: { countSlidesInSourcePresentation}");

                        if (countSlidesInSourcePresentation > 0)
                        {
                            for (int slideIndex = 1; slideIndex <= countSlidesInSourcePresentation; slideIndex++)
                            {
                                tracingSvc.Trace($"Merge slides: {slideIndex}");
                                MergePresentationSlides(ref destDoc, ref destPresentationPart, ref destPresentation, ref sourcePresentationPart, slideIndex);
                            }
                        }

                        sourceDoc.Close();
                        destDoc.Close();                        

                        outputFile.Seek(0, SeekOrigin.Begin);
;
                        tracingSvc.Trace($"Merge completes. Writing to Table");
                        this.UploadFile(localPluginContext.CurrentUserService, entityRef, "rio_mergeddeck", outputFile);
                        tracingSvc.Trace("Slides merged");
                    }
                }
            }
        }

        internal byte[] GetStoryTimeTemplateFile(IOrganizationService service, Guid templateId)
        {
            Entity templateEntity = service.Retrieve("rio_template", templateId, new Microsoft.Xrm.Sdk.Query.ColumnSet("rio_template"));

            //checks if file is present
            if (templateEntity.Contains("rio_template"))
            {
                //get the file
                InitializeFileBlocksDownloadRequest fileRequest = new InitializeFileBlocksDownloadRequest();
                fileRequest.Target = new EntityReference("rio_template", templateId);
                fileRequest.FileAttributeName = "rio_template";
                InitializeFileBlocksDownloadResponse fileResponse = (InitializeFileBlocksDownloadResponse)service.Execute(fileRequest);
                DownloadBlockRequest fileDownloadRequest = new DownloadBlockRequest();
                fileDownloadRequest.FileContinuationToken = fileResponse.FileContinuationToken;
                DownloadBlockResponse fileDownloadResponse = (DownloadBlockResponse)service.Execute(fileDownloadRequest);

                return fileDownloadResponse.Data;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Uploads a file or image column value
        /// </summary>
        /// <param name="service">The service</param>
        /// <param name="entityReference">A reference to the record with the file or image column</param>
        /// <param name="fileAttributeName">The name of the file or image column</param>
        /// <param name="fileInfo">Information about the file or image to upload.</param>
        /// <returns></returns>
        internal Guid UploadFile(IOrganizationService service, EntityReference entityReference, string fileAttributeName,MemoryStream file)
        {
            string fileName = $"StoryTime {entityReference.Id.ToString()}.pptx";
            // Initialize the upload
            InitializeFileBlocksUploadRequest initializeFileBlocksUploadRequest = new InitializeFileBlocksUploadRequest()
            {
                Target = entityReference,
                FileAttributeName = fileAttributeName,
                FileName = fileName
            };

            var initializeFileBlocksUploadResponse = (InitializeFileBlocksUploadResponse)service.Execute(initializeFileBlocksUploadRequest);
            string fileContinuationToken = initializeFileBlocksUploadResponse.FileContinuationToken;

            // Capture blockids while uploading for chunking uploads (> 4MB)
            List<string> blockIds = new List<string>();
            int blockSize = 4 * 1024 * 1024; // 4 MB
            byte[] buffer = new byte[blockSize];
            int bytesRead = 0;

            long fileSize = file.Length;

            // The number of iterations that will be required:
            // int blocksCount = (int)Math.Ceiling(fileSize / (float)blockSize);
            int blockNumber = 0;
                        
            // While there is unread data from the file
            while ((bytesRead = file.Read(buffer, 0, buffer.Length)) > 0)
            {
                // The file or final block may be smaller than 4MB
                if (bytesRead < buffer.Length) Array.Resize(ref buffer, bytesRead);

                blockNumber++;

                string blockId = Convert.ToBase64String(Encoding.UTF8.GetBytes(Guid.NewGuid().ToString()));
                blockIds.Add(blockId);

                // Prepare the request
                UploadBlockRequest uploadBlockRequest = new UploadBlockRequest()
                {
                    BlockData = buffer,
                    BlockId = blockId,
                    FileContinuationToken = fileContinuationToken,
                };

                // Send the request
                service.Execute(uploadBlockRequest);
            }
            file.Close();
            file.Dispose();
            

            // Commit the upload
            CommitFileBlocksUploadRequest commitFileBlocksUploadRequest = new CommitFileBlocksUploadRequest()
            {
                BlockList = blockIds.ToArray(),
                FileContinuationToken = fileContinuationToken,
                FileName = fileName,
                MimeType = "application/vnd.openxmlformats-officedocument.presentationml.presentation" //powerpoint Mime type
            };

            var commitFileBlocksUploadResponse = (CommitFileBlocksUploadResponse)service.Execute(commitFileBlocksUploadRequest);

            return commitFileBlocksUploadResponse.FileId;

        }

        private void MergePresentationSlides(
            ref PresentationDocument destDoc, ref PresentationPart destPresentationPart, 
            ref Presentation destPresentation, ref PresentationPart sourcePresentationPart, int copiedSlideIndex)
        {
            SlideId copiedSlideId = sourcePresentationPart.Presentation.SlideIdList.ChildElements[--copiedSlideIndex] as SlideId;
            SlidePart copiedSlidePart = sourcePresentationPart.GetPartById(copiedSlideId.RelationshipId) as SlidePart;

            SlidePart addedSlidePart = destPresentationPart.AddPart<SlidePart>(copiedSlidePart);

            NotesSlidePart noticePart = addedSlidePart.GetPartsOfType<NotesSlidePart>().FirstOrDefault();
            if (noticePart != null) addedSlidePart.DeletePart(noticePart);

            SlideMasterPart addedSlideMasterPart = destPresentationPart.AddPart(addedSlidePart.SlideLayoutPart.SlideMasterPart);

            // Create new slide ID
            SlideId slideId = new SlideId
            {
                Id = CreateId(destPresentation.SlideIdList),
                RelationshipId = destDoc.PresentationPart.GetIdOfPart(addedSlidePart)
            };
            destPresentation.SlideIdList.Append(slideId);

            // Create new master slide ID
            uint masterId = CreateId(destPresentation.SlideMasterIdList);
            SlideMasterId slideMaterId = new SlideMasterId
            {
                Id = masterId,
                RelationshipId = destDoc.PresentationPart.GetIdOfPart(addedSlideMasterPart)
            };
            destDoc.PresentationPart.Presentation.SlideMasterIdList.Append(slideMaterId);

            destDoc.PresentationPart.Presentation.Save();

            // Make sure that all slide layouts have unique ids.
            foreach (SlideMasterPart slideMasterPart in destDoc.PresentationPart.SlideMasterParts)
            {
                foreach (SlideLayoutId slideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList)
                {
                    masterId++;
                    slideLayoutId.Id = masterId;
                }

                slideMasterPart.SlideMaster.Save();
            }

            destDoc.PresentationPart.Presentation.Save();
        }

        private uint CreateId(SlideMasterIdList slideMasterIdList)
        {
            uint currentId = 0;
            foreach (SlideMasterId masterId in slideMasterIdList)
            {
                if (masterId.Id > currentId)
                {
                    currentId = masterId.Id;
                }
            }
            return ++currentId;
        }
        private uint CreateId(SlideIdList slideIdList)
        {
            uint currentId = 0;
            foreach (SlideId slideId in slideIdList)
            {
                if (slideId.Id > currentId)
                {
                    currentId = slideId.Id;
                }
            }
            return ++currentId;
        }
    }
    internal class Template
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
    }
}
