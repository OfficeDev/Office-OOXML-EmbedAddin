// Copyright(c) Microsoft.All rights reserved.Licensed under the MIT license.See full license at the root of the repo.

//This file provides the button event handlers and basic UI manipulation.

using System;
using System.IO;
using System.Text.RegularExpressions;

namespace Office_OOXML_EmbedAddin
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (FileUploadControl.HasFile)
                try
                {
                    // The OOXML SDK has different classes for the different Office
                    // file types, so the code needs to know which classes to use.
                    // So cache the Office file type; for example "Excel".
                    var file = new FileInfo(FileUploadControl.FileName);
                    Session["FileType"] = GetFileType(file);

                    if (Session["FileType"] != null)
                    {
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            memoryStream.Write(FileUploadControl.FileBytes, 0, FileUploadControl.FileBytes.Length);

                            // Validate the file. Throws exception if the file is not valid.
                            OOXMLHelper.ValidateOfficeFile(Session["FileType"].ToString(), memoryStream);

                            lblUploadOutcome.Text = "File uploaded successfully";
                            Session["ByteArray"] = FileUploadControl.FileBytes;
                            Session["UploadedFileName"] = FileUploadControl.FileName;
                        }
                    }
                    else
                    {
                        lblUploadOutcome.Text = "You must choose an Office file with an extension of .xlsx, .docx, or .pptx.";
                        Session["ByteArray"] = null;
                        Session["UploadedFileName"] = null;
                    }
                }
                catch (Exception ex)
                {
                    lblUploadOutcome.Text = "ERROR: " + ex.Message.ToString();
                    Session["ByteArray"] = null;
                    Session["UploadedFileName"] = null;
                }
            else
            {
                lblUploadOutcome.Text = "You have not specified a file.";
                Session["ByteArray"] = null;
                Session["UploadedFileName"] = null;
            }
        }


        protected void btnEmbed_Click(object sender, EventArgs e)
        {
            string snippetID;

            var validGist = new Regex("^[a-zA-Z0-9]+$");
            if (validGist.IsMatch(tbSnippetID.Text))
            {
                snippetID = tbSnippetID.Text;
            }
            else
            {
                lblSnippetOutcome.Text = "A Gist ID has only letters and numerals. Default snippet will be used.";
                // Set a default snippet.
                snippetID = "c7ba602e8a107cfe5dd6c42ca41deac1";
            }

            // Get the file.
            byte[] byteArray = (byte[])(Session["ByteArray"]);

            if (byteArray != null)
            {
                try
                {
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        memoryStream.Write(byteArray, 0, byteArray.Length);
                        AddinEmbedder.EmbedAddin(Session["FileType"].ToString(), memoryStream, snippetID);

                        // Save the modified file.
                        Session["ByteArray"] = memoryStream.ToArray();

                        lblEmbedOutcome.Text = "The file has been configured to automatically open Script Lab";
                    }
                }
                catch (Exception ex)
                {
                    lblEmbedOutcome.Text = "ERROR: " + ex.Message.ToString();                    
                }
            }
            else
            {
                lblEmbedOutcome.Text = "Sorry, we seem to have lost that file. Please upload it again.";
            }
        }

        protected void btnDownload_Click(object sender, EventArgs e)
        {
            // Get the stored file so it can be downloaded.
            byte[] byteArray = (byte[])(Session["ByteArray"]);

            if (byteArray != null)
            {
                Response.Clear();
                Response.ContentType = "application/octet-stream";
                string fileName = (string)(Session["UploadedFileName"]);
                Response.AddHeader("Content-Disposition",
                    String.Format("attachment; filename={0}", fileName));
                Response.BinaryWrite(byteArray);
                Response.Flush();
                Response.End();
            }
            else
            {
                lblMissingFile.Text = "Sorry, we seem to have lost that file. Please upload it again.";
            }
        }

        private string GetFileType(FileInfo file)
        {
            switch (file.Extension.ToLower())
            {
                case ".xlsx":
                    return "Excel";
                case ".docx":
                    return "Word";
                case ".pptx":
                    return "PowerPoint";
                default:
                    return null;
            }
        }

    }

}