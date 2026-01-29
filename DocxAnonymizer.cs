using Microsoft.AspNetCore.Http;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Net;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml;

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
            // 1. Odczytujemy CAŁE body jako BinaryData - to rozwiązuje problem pustych strumieni
            var data = await BinaryData.FromStreamAsync(req.Body);
            byte[] inputBytes = data.ToArray();

            if (inputBytes.Length == 0)
            {
                _logger.LogWarning("Otrzymano puste body.");
                return req.CreateResponse(HttpStatusCode.BadRequest);
            }

            _logger.LogInformation($"Odebrano plik: {inputBytes.Length} bajtów");

            // 2. Przetwarzanie
            byte[] outputBytes = this.AnonymizeDocx(inputBytes);

            // 3. Budowanie odpowiedzi
            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            
            // Ważne: Zapisujemy bajty bezpośrednio
            await response.WriteBytesAsync(outputBytes);
            
            _logger.LogInformation($"Wysłano plik: {outputBytes.Length} bajtów");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Błąd krytyczny: {ex.Message} \n {ex.StackTrace}");
            return req.CreateResponse(HttpStatusCode.InternalServerError);
        }
    }

    private byte[] AnonymizeDocx(byte[] inputBytes)
    {
        using var ms = new MemoryStream();
        ms.Write(inputBytes, 0, inputBytes.Length);
        ms.Position = 0;

        // Otwieramy dokument w strumieniu
        using (WordprocessingDocument doc = WordprocessingDocument.Open(ms, true))
        {
            // Czyszczenie właściwości
            _logger.LogInformation("Czyszczenie właściwości dokumentu...");
            var cp = doc.PackageProperties;
            if (cp.LastPrinted != null)
            {
                _logger.LogInformation($"Usuwanie daty ostatniego wydruku: {cp.LastPrinted}");
                cp.LastPrinted = null;
            }
            if (cp.Creator != null && cp.Creator != "")
            {
                _logger.LogInformation($"Usuwanie Creator: {cp.Creator}");
                cp.Creator = null;
            }
            if (cp.LastModifiedBy != null && cp.LastModifiedBy != "")
            {
                _logger.LogInformation($"Usuwanie LastModifiedBy: {cp.LastModifiedBy}");
                cp.LastModifiedBy = null;
            }
            if (cp.Title != null && cp.Title != "")
            {
                _logger.LogInformation($"Usuwanie Title: {cp.Title}");
                cp.Title = null;
            }
            if (cp.Subject != null && cp.Subject != "")
            {
                _logger.LogInformation($"Usuwanie Subject: {cp.Subject}");
                cp.Subject = null;
            }
            if (cp.Description != null)
            {
                _logger.LogInformation($"Usuwanie Description: {cp.Description}");
                cp.Description = null;
            }
            if (cp.Keywords != null && cp.Keywords != "")
            {
                _logger.LogInformation($"Usuwanie Keywords: {cp.Keywords}");
                cp.Keywords = null;
            }
            if (cp.Created != null)
            {
                _logger.LogInformation($"Zmiana daty Created: {cp.Created} na 1970-01-01");
                cp.Created = new DateTime(1970, 1, 1);
            }
            if (cp.Modified != null)
            {
                _logger.LogInformation($"Zmiana daty Modified: {cp.Modified} na 1970-01-01");
                cp.Modified = new DateTime(1970, 1, 1);
            }
            if (cp.Revision != null)
            {
                _logger.LogInformation($"Usuwanie Revision: {cp.Revision}");
                cp.Revision = null;
            }

            // EXTENDED PROPERTIES (docProps/app.xml)
            var extProperties = doc.ExtendedFilePropertiesPart ?? doc.AddNewPart<ExtendedFilePropertiesPart>();
            var extPropsRoot = extProperties.Properties ?? new Properties();

            if (extPropsRoot.Template != null)
            {
                _logger.LogInformation($"Czyszczenie Template: {extPropsRoot.Template.Text}");
                extPropsRoot.Template = null;
            }
            if (extPropsRoot.TotalTime != null)
            {
                _logger.LogInformation($"Ustawianie TotalTime: {extPropsRoot.TotalTime.Text} na wartość 0");
                extPropsRoot.TotalTime = new TotalTime { Text = "0" };
            }
            if (extPropsRoot.Company != null)
            {
                _logger.LogInformation($"Czyszczenie Company: {extPropsRoot.Company.Text}");
                extPropsRoot.Company = new Company { Text = "" };
            }
            if (extPropsRoot.Manager != null)
            {
                _logger.LogInformation($"Czyszczenie Manager: {extPropsRoot.Manager.Text}");
                extPropsRoot.Manager = new Manager { Text = "" };
            }
            
            _logger.LogInformation("Zapis Extended Properties");
            extProperties.Properties?.Save(); // zapis root-a partu

            // SETTINGS
            var settingsPart = doc.MainDocumentPart?.DocumentSettingsPart ?? doc.MainDocumentPart?.AddNewPart<DocumentSettingsPart>();
            
            if (settingsPart != null)
            {
                if (settingsPart.Settings == null) 
                {
                    _logger.LogInformation("Brak Settings w DocumentSettingsPart, tworzenie nowego.");
                    settingsPart.Settings = new Settings();
                }
                
                // Flaga anonimizacji
                if (settingsPart.Settings.GetFirstChild<RemovePersonalInformation>() == null)
                {
                    _logger.LogInformation("Dodawanie RemovePersonalInformation do ustawień dokumentu");
                    settingsPart.Settings.Append(new RemovePersonalInformation());
                }
                
                // Zamiana szablonu na Normal.dotm
                AttachedTemplate? attachedTemplateElement = settingsPart.Settings.GetFirstChild<AttachedTemplate>();
                if (attachedTemplateElement != null)
                {
                    _logger.LogInformation($"Usuwanie AttachedTemplate: {attachedTemplateElement.InnerText}");
                    attachedTemplateElement.Remove();
                }
            }

            // CUSTOM XML - USUNIĘCIE
            _logger.LogInformation("Usuwanie Custom XML Parts");
            doc.MainDocumentPart?.DeleteParts(doc.MainDocumentPart.CustomXmlParts);

            ProcessRevisions(doc);
            ProcessComments(doc);
            doc.Save();
        }

        // Zwracamy tablicę bajtów po zamknięciu WordprocessingDocument (Dispose jest kluczowe!)
        return ms.ToArray();
    }

    private void ProcessRevisions(WordprocessingDocument doc)
    {
        _logger.LogInformation("Anonimizacja autorów w śledzeniu zmian (Body, Nagłówki, Stopki)...");

        if (doc.MainDocumentPart == null) return;

        var partsToProcess = new List<OpenXmlPart>();
        partsToProcess.Add(doc.MainDocumentPart);
        partsToProcess.AddRange(doc.MainDocumentPart.HeaderParts);
        partsToProcess.AddRange(doc.MainDocumentPart.FooterParts);

        // Przestrzeń nazw dla WordprocessingML
        string wmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        foreach (var part in partsToProcess)
        {
            var root = part.RootElement;
            if (root == null) continue;

            // Pobieramy tylko te elementy, które faktycznie mają zdefiniowany atrybut 'author'
            // Używamy GetAttributes(), co jest bezpieczne i nie rzuca wyjątku przy braku uprawnień elementu do atrybutu
            var elementsWithAuthor = root.Descendants().Where(e => 
                e.GetAttributes().Any(a => a.LocalName == "author" && a.NamespaceUri == wmlNamespace));

            foreach (var el in elementsWithAuthor)
            {
                // Pobieramy istniejący atrybut, żeby zachować jego prefix (zazwyczaj 'w')
                var existingAttr = el.GetAttributes().First(a => a.LocalName == "author" && a.NamespaceUri == wmlNamespace);
                
                // Ustawiamy nową wartość
                _logger.LogInformation($"Anonimizacja elementu {el.LocalName}, ustawianie author na 'Author'");
                el.SetAttribute(new OpenXmlAttribute(
                    existingAttr.Prefix, 
                    "author", 
                    wmlNamespace, 
                    "Author"));
            }
        }
    }

    private void ProcessComments(WordprocessingDocument doc)
    {
        _logger.LogInformation("Anonimizacja autorów komentarzy i śledzenia zmian...");

        if (doc.MainDocumentPart == null) return;

        var commentsPart = doc.MainDocumentPart.WordprocessingCommentsPart;
        if (commentsPart?.Comments != null)
        {
            foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
            {
                _logger.LogInformation($"Anonimizacja komentarza ID: {comment.Id}, Autor: {comment.Author}, Inicjały: {comment.Initials}");
                if (comment.Author != null) comment.Author = "Author";
                if (comment.Initials != null) comment.Initials = "A";
            }
            commentsPart.Comments.Save();
        }
    }
}