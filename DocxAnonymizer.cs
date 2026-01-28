using Microsoft.AspNetCore.Http;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Net;

namespace RCL.PL;

public class DocxAnonymizer
{
    private readonly ILogger<DocxAnonymizer> _logger;

    public DocxAnonymizer(ILogger<DocxAnonymizer> logger)
    {
        _logger = logger;
    }

    [Function("DocxAnonymizer")]
    public async Task<HttpResponseData> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req)
    {
        _logger.LogInformation("Anonimizacja DOCX rozpoczęta");
        
        try
        {
            var inputStream = new MemoryStream();
            await req.Body.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var outputStream = new MemoryStream();
            AnonymizeDocx(inputStream, outputStream);

            outputStream.Position = 0;
            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            response.Body = outputStream;
            
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Błąd: {ex.Message}");
            return req.CreateResponse(HttpStatusCode.BadRequest);
        }
    }

    private static void AnonymizeDocx(Stream inputStream, Stream outputStream)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(inputStream, true))
        {
            var coreProps = doc.PackageProperties;
            coreProps.Creator = null;
            coreProps.LastModifiedBy = null;
            coreProps.LastPrinted = null;
            coreProps.Created = null;
            coreProps.Modified = null;

            var settingsPart = doc.MainDocumentPart.DocumentSettingsPart ?? 
                            doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
            if (settingsPart.Settings == null) 
                settingsPart.Settings = new Settings();
            
            if (settingsPart.Settings.GetFirstChild<RemovePersonalInformation>() == null)
            {
                settingsPart.Settings.Append(new RemovePersonalInformation());
            }

            doc.MainDocumentPart.DeleteParts(doc.MainDocumentPart.CustomXmlParts);
            doc.Save();
            inputStream.Position = 0;
            inputStream.CopyTo(outputStream);
        }
    }
}