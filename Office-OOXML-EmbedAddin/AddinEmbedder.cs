// Copyright(c) Microsoft.All rights reserved.Licensed under the MIT license.See full license at the root of the repo.

// This file provides the business logic. It makes calls into the OOXML helper.

using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;

namespace Office_OOXML_EmbedAddin
{
    public static class AddinEmbedder
    {
        // Embeds the add-in into a file of the specified type.
        public static void EmbedAddin(string fileType, MemoryStream memoryStream, string snippetID)
        {
            // Each Office file type has its own *Document class in the OOXML SDK.
            switch (fileType)
            {
                case "Excel":
                    using (var spreadsheet = SpreadsheetDocument.Open(memoryStream, true))
                    {
                        spreadsheet.DeletePart(spreadsheet.WebExTaskpanesPart);
                        var webExTaskpanesPart = spreadsheet.AddWebExTaskpanesPart();
                        OOXMLHelper.CreateWebExTaskpanesPart(webExTaskpanesPart, snippetID);
                    }
                    break;
                case "Word":
                    using (var document = WordprocessingDocument.Open(memoryStream, true))
                    {
                        var webExTaskpanesPart = document.AddWebExTaskpanesPart();
                        OOXMLHelper.CreateWebExTaskpanesPart(webExTaskpanesPart, snippetID);
                    }
                    break;
                case "PowerPoint":
                    using (var slidedeck = PresentationDocument.Open(memoryStream, true))
                    {
                        var webExTaskpanesPart = slidedeck.AddWebExTaskpanesPart();
                        OOXMLHelper.CreateWebExTaskpanesPart(webExTaskpanesPart, snippetID);
                    }
                    break;
                default:
                    throw new Exception("Invalid File Type");
            }
        }
    }
}