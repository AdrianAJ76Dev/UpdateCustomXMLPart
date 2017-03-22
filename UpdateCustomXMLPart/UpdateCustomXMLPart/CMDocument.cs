﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Open XML SDK
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using System.IO;

namespace UpdateCustomXMLPart
{
    class CMDocument
    {
        //private const string strFULLPATH_SSL_SAMPLE = @"C:\Users\ajones\Documents\Automation\Documents\Sole Source Letter Custom XML Part Sample.docx";
        private const string strFULLPATH_SSL_SAMPLE = @"C:\Users\ajones\Documents\Automation\Documents\Sole Source Letter Custom XML Part Sample v2.docx";
        private const string strFULLPATH_SSL_XML_SAMPLE = @"C:\Users\ajones\Documents\Automation\Documents\XML\SSL XML Sample 2.xml";

        public void UpdateSSLContactInfo()
        {
            // Open a document with a Custom XML Part
            using (WordprocessingDocument SSLDoc = WordprocessingDocument.Open(strFULLPATH_SSL_SAMPLE, true))
            {
                MainDocumentPart SSLMain = SSLDoc.MainDocumentPart;
                Console.WriteLine("Custom XML Parts Count => {0}", SSLMain.CustomXmlParts.Count());

                foreach (CustomXmlPart currCXP in SSLMain.CustomXmlParts)
                {
                    Console.WriteLine("Uri: {0}", currCXP.Uri.ToString());
                    Console.WriteLine("Child Part (CustomXmlProperties) Uri: {0}", currCXP.CustomXmlPropertiesPart.Uri.ToString());
                    Console.WriteLine("Child Part (Relationship) {0}", currCXP.RelationshipType);
                    Console.WriteLine("DataStoreItem: {0}", currCXP.CustomXmlPropertiesPart.DataStoreItem.ItemId);
                    Console.WriteLine();

                    /* Get the stream of the XML File sample 
                     * THIS WORKS!!!!
                     */
                    //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                    StreamReader strdXML = new StreamReader(strFULLPATH_SSL_XML_SAMPLE);
                    currCXP.FeedData(strdXML.BaseStream);
                    //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


                    /* This just points out that the CustomXmlPropertiesPart exists I can get this without cycling
                     * through the custom xml part's parts
                    foreach (IdPartPair ipp in currCXP.Parts)
                    {
                        Console.WriteLine("Part 1 Found {0}", ipp.OpenXmlPart.Uri.ToString());
                    }
                    */
                }
                Console.ReadLine();
                SSLDoc.MainDocumentPart.Document.Save();


                /* 
                 * Find Custom XML Part bound to the Content Controls
                 * Read New Custom XML Part replacing the current Custom XML Part (for now will be in a special directory)
                 * Use FeedData of Old Custom XML Part to replace with New Custom XML Part GetStream
                 */
            }

        }

    }
}
