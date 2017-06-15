<%--Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the root of the repo.--%>
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Home.aspx.cs" Inherits="Office_OOXML_EmbedAddin.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <h1>Embed Script Lab in an Office File</h1>
            <p>1. Browse to any Excel (.xlsx), Word (.docx), PowerPoint (.pptx) file <b>that does not already have an add-in set to automatically open with the file</b>, and then click <b>Upload</b></p>
            <asp:FileUpload ID="FileUploadControl" runat="server" />
            <p />
            <asp:Button ID="btnUpload" runat="server" Text="Upload" OnClick="btnUpload_Click" />
            <asp:Label ID="lblUploadOutcome" runat="server" Text=""></asp:Label>
            <p />
            <hr /> 
            <p>2. Press <b>Embed Script Lab</b> to configure the document to open Script Lab automatically when the document opens. Optionally, you may enter the gist ID of a snippet you want to import into Script Lab.</p>
            <asp:Label ID="lblSnippetID" runat="server" Text="Label">Snippet ID: </asp:Label>
            <asp:TextBox ID="tbSnippetID" runat="server"></asp:TextBox>
            <asp:Label ID="lblSnippetOutcome" runat="server" Text=""></asp:Label>
            <p />
            <asp:Button ID="btnEmbed" runat="server" Text="Embed Script Lab" OnClick="btnEmbed_Click" />
            <asp:Label ID="lblEmbedOutcome" runat="server" Text=""></asp:Label>
            <p />
            <hr /> 
            <p>3. Click <b>Download</b>. Your browser will give you the options of opening the file, saving it, or saving it under an alternate name.</p>
            <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" />
            <asp:Label ID="lblMissingFile" runat="server" Text=""></asp:Label>
            <p />
            <p>When you open the file for the first time, the taskpane opens automatically and you are prompted to trust Script Lab. In the future, whenever you open the file, Script Lab opens automatically.</p>
        </div>
    </form>
</body>
</html>
