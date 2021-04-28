using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace VSML
{
    class VsmlWriter
    {
        private XmlDocument XmlDocument { get; set; }
        private List<Style> AvailableStyles = new List<Style>();
        private readonly List<String> attributeNodes = new List<string>
            {
                "#text",
                "weight",
                "bold",
                "italic",
                "underlined",
                "strikethrough",
                "colour",
                "align",
                "list_id"
            };

        public VsmlWriter(XmlDocument document)
        {
            this.XmlDocument = document;
            InitialiseStyles();
        }

        /// <summary>
        /// Adds all of the predefined styles to the list of styles 
        /// </summary>
        private void InitialiseStyles()
        {
            int h1Size = 24;
            int h2Size = 16;
            int h3Size = 14;
            int h4Size = 13;
            int h5Size = 12;
            int h6Size = 11;

            // Heading 1
            StyleName vsmlHeading1StyleName = new StyleName() { Val = "vsml Heading 1" };
            Style vsmlHeading1Style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "vsmlHeading1",
                CustomStyle = true,
                Default = false
            };
            vsmlHeading1Style.Append(vsmlHeading1StyleName);
            vsmlHeading1Style.Append(new BasedOn() { Val = "Normal" });
            vsmlHeading1Style.Append(new NextParagraphStyle() { Val = "Normal" });
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            FontSize fontSize1 = new FontSize() { Val = (h1Size * 2).ToString() };
            styleRunProperties1.Append(new Bold());
            styleRunProperties1.Append(fontSize1);
            vsmlHeading1Style.Append(styleRunProperties1);
            AvailableStyles.Add(vsmlHeading1Style);

            // Heading 2
            StyleName vsmlHeading2StyleName = new StyleName() { Val = "vsml Heading 2" };
            Style vsmlHeading2Style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "vsmlHeading2",
                CustomStyle = true,
                Default = false
            };
            vsmlHeading2Style.Append(vsmlHeading2StyleName);
            vsmlHeading2Style.Append(new BasedOn() { Val = "Normal" });
            vsmlHeading2Style.Append(new NextParagraphStyle() { Val = "Normal" });
            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            FontSize fontSize2 = new FontSize() { Val = (h2Size * 2).ToString() };
            styleRunProperties2.Append(new Bold());
            styleRunProperties2.Append(fontSize2);
            vsmlHeading2Style.Append(styleRunProperties2);
            AvailableStyles.Add(vsmlHeading2Style);

            // Heading 3
            StyleName vsmlHeading3StyleName = new StyleName() { Val = "vsml Heading 3" };
            Style vsmlHeading3Style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "vsmlHeading3",
                CustomStyle = true,
                Default = false
            };
            vsmlHeading3Style.Append(vsmlHeading3StyleName);
            vsmlHeading3Style.Append(new BasedOn() { Val = "Normal" });
            vsmlHeading3Style.Append(new NextParagraphStyle() { Val = "Normal" });
            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            FontSize fontSize3 = new FontSize() { Val = (h3Size * 2).ToString() };
            styleRunProperties3.Append(new Bold());
            styleRunProperties3.Append(fontSize3);
            vsmlHeading3Style.Append(styleRunProperties3);
            AvailableStyles.Add(vsmlHeading3Style);

            // Heading 4
            StyleName vsmlHeading4StyleName = new StyleName() { Val = "vsml Heading 4" };
            Style vsmlHeading4Style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "vsmlHeading4",
                CustomStyle = true,
                Default = false
            };
            vsmlHeading4Style.Append(vsmlHeading4StyleName);
            vsmlHeading4Style.Append(new BasedOn() { Val = "Normal" });
            vsmlHeading4Style.Append(new NextParagraphStyle() { Val = "Normal" });
            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            FontSize fontSize4 = new FontSize() { Val = (h4Size * 2).ToString() };
            styleRunProperties4.Append(new Bold());
            styleRunProperties4.Append(fontSize4);
            vsmlHeading4Style.Append(styleRunProperties4);
            AvailableStyles.Add(vsmlHeading4Style);

            // Heading 5
            StyleName vsmlHeading5StyleName = new StyleName() { Val = "vsml Heading 5" };
            Style vsmlHeading5Style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "vsmlHeading5",
                CustomStyle = true,
                Default = false
            };
            vsmlHeading5Style.Append(vsmlHeading5StyleName);
            vsmlHeading5Style.Append(new BasedOn() { Val = "Normal" });
            vsmlHeading5Style.Append(new NextParagraphStyle() { Val = "Normal" });
            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            FontSize fontSize5 = new FontSize() { Val = (h5Size * 2).ToString() };
            styleRunProperties5.Append(new Bold());
            styleRunProperties5.Append(new Italic());
            styleRunProperties5.Append(fontSize5);
            vsmlHeading5Style.Append(styleRunProperties5);
            AvailableStyles.Add(vsmlHeading5Style);

            // Heading 6
            StyleName vsmlHeading6StyleName = new StyleName() { Val = "vsml Heading 6" };
            Style vsmlHeading6Style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "vsmlHeading6",
                CustomStyle = true,
                Default = false
            };
            vsmlHeading6Style.Append(vsmlHeading6StyleName);
            vsmlHeading6Style.Append(new BasedOn() { Val = "Normal" });
            vsmlHeading6Style.Append(new NextParagraphStyle() { Val = "Normal" });
            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            FontSize fontSize6 = new FontSize() { Val = (h6Size * 2).ToString() };
            styleRunProperties6.Append(new Bold());
            styleRunProperties5.Append(new Italic());
            styleRunProperties6.Append(fontSize6);
            vsmlHeading6Style.Append(styleRunProperties6);
            AvailableStyles.Add(vsmlHeading6Style);
        }

        /// <summary>
        /// Creates a .docx file from the XML stored in the XmlDocument attribute
        /// </summary>
        /// <param name="fileName">The location of the file to be created</param>
        /// <returns>
        /// Tuple, item1 is boolean, true if successfully created Document false otherwise.
        /// item2 is string, containing error message if there is any.
        /// </returns>
        public Tuple<bool, string> ToDocx(string fileName)
        {
            // Remove unnesseccary XML to simplfy
            XmlDocument doc = RemoveExtraXml(XmlDocument);
            try
            {
                // Create a Document by supplying the filepath. 
                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
                {
                    // Add a main document part. 
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    // Add a styles part
                    StyleDefinitionsPart part = wordDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                    Styles styles = new Styles();
                    styles.Save(part);

                    // Add all styles to the document
                    foreach (Style style in AvailableStyles)
                        part.Styles.AppendChild(style.CloneNode(true));

                   
                    // Add numbering element for bullet lists 
                    NumberingDefinitionsPart numberingPart = 
                        mainPart.AddNewPart<NumberingDefinitionsPart>("someUniqueIdHere");

                    int numberOfLists = XmlDocument.GetElementsByTagName("list").Count;

                    // Add numbering element for bullet lists 
                    Numbering element = new Numbering();
                    element.AppendChild(
                        new AbstractNum(
                            new Level(
                                new NumberingFormat() { Val = NumberFormatValues.Bullet }, //type of char to mark list
                                new LevelText() { Val = "·" },
                                new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel },
                                new StartNumberingValue() { Val = 1 }
                            )
                            { LevelIndex = 0 }
                        )
                        { AbstractNumberId = 1 });

                    for (int i = 0; i < numberOfLists; i++)
                    {
                        element.AppendChild(
                       new AbstractNum(
                           new Level(
                               new NumberingFormat() { Val = NumberFormatValues.Decimal }, //type of char to mark list
                               new LevelText() { Val = "%1." },
                               new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel },
                               new StartNumberingValue() { Val = 1 }
                           )
                           { LevelIndex = 0 }
                       )
                       { AbstractNumberId = i + 2 });
                    }
                    element.AppendChild(
                        new NumberingInstance(
                            new AbstractNumId() { Val = 1 })
                        { NumberID = 1 });
                    for (int i = 0; i < numberOfLists; i++)
                    {
                        element.AppendChild(
                        new NumberingInstance(
                            new AbstractNumId() { Val = 2 + i })
                        { NumberID = 2 + i });
                    }

                    element.Save(numberingPart);

                    // Create the word document using a tree traversal
                    mainPart.Document = (Document)RecursivelyCreateWordDoc(doc.ChildNodes.Item(0), wordDocument);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("hey?");
                return new Tuple<bool, string>(false, e.Message);
            }
            return new Tuple<bool, string>(true, "");
        }

        /// <summary>
        /// Takes the xml input and removes any nodes that are not relevant to a .docx context
        /// Nodes are removed in place and their child nodes added to the parent of the removed
        /// node
        /// </summary>
        /// <param name="doc">The xml Document to remove the nodes from</param>
        /// <returns>The supplied xml Document with the nodes removed</returns>
        private XmlDocument RemoveExtraXml(XmlDocument doc)
        {
            List<XmlNode> nodesToKeep = RemoveUnneededXmlNodes(doc.ChildNodes.Item(0), new List<XmlNode>());

            // Remake the XmlDocument 
            // Remove all the old children
            doc.ChildNodes.Item(0).RemoveAll();

            foreach (XmlNode child in doc.ChildNodes)
                Console.WriteLine(child.Name);

            // Re-add namespace attribute
            ((XmlElement)doc.ChildNodes.Item(0)).SetAttribute("xmlns", "vsml");

            // Re-add the neccessary children
            foreach (XmlNode n in nodesToKeep)
                doc.ChildNodes.Item(0).AppendChild(n.Clone());

            return doc;
        }

        /// <summary>
        /// Removes the nodes form the xml Document by using an in order treee traversal
        /// </summary>
        /// <param name="root">The current root node of the traversal</param>
        /// <param name="nodesToAdd">A list representing the nodes that are to
        /// be added back to the Document when others have been removed/param>
        /// <returns>
        /// A list of XmlNodes representing the nodes that the Document should
        /// be comprised of 
        /// </returns>
        private List<XmlNode> RemoveUnneededXmlNodes(XmlNode root,
            List<XmlNode> nodesToAdd)
        {
            List<string> nodesToKeep = new List<string> {
                "paragraph",
                "heading",
                "o_list_item",
                "u_list_item",
                "page_break", };

            // If the nodes is meant to be kept
            if (nodesToKeep.Contains(root.Name))
                // Keep it
                nodesToAdd.Add(root.Clone());

            if (root.ChildNodes.Count > 0)
                for (int i = 0; i < root.ChildNodes.Count; i++)
                    nodesToAdd = RemoveUnneededXmlNodes(root.ChildNodes.Item(i), nodesToAdd);

            return nodesToAdd;
        }

        /// <summary>
        /// Carries out a traversal of the XmlDocument tree, and at each node adds the 
        /// suitable word docuemnt equivalent to the word Document
        /// </summary>
        /// <param name="root">The current root of the traversal</param>
        /// <param name="document">The word docuemnt that the elements are added to</param>
        /// <returns>A word Document with the root nodes information added to it</returns>
        private OpenXmlElement RecursivelyCreateWordDoc(XmlNode root,
            WordprocessingDocument document)
        {
            OpenXmlElement thisElement = CreateOpenXmlElement(root, document);

            // If the current node has any children
            if (root.ChildNodes.Count > 0)
            {
                // Add the children to the node
                foreach (XmlNode ele in root.ChildNodes)
                {
                    // Explore each node that is not an attribute node
                    if (!this.attributeNodes.Contains(ele.Name))
                        thisElement.Append(RecursivelyCreateWordDoc(ele, document));
                }
            }
            else
                return thisElement;
            return thisElement;
        }

        /// <summary>
        /// Create an OpenXml element from a given XmlNode
        /// </summary>
        /// <param name="root">The XmlNode to create an OpenXmlElement of</param>
        /// <param name="document">The Document that the element will be a part of</param>
        /// <returns>Element representing the supplied root element</returns>
        private OpenXmlElement CreateOpenXmlElement(XmlNode root,
            WordprocessingDocument document)
        {
            // Look at this node, create the correct part of word doc
            String type = root.Name;
            Console.WriteLine("matthew" + type);

            OpenXmlElement thisElement;

            switch (type)
            {
                case "doc":
                    thisElement = new Document();
                    break;
                case "paragraph":
                    thisElement = new Paragraph();

                    // Add attributes
                    ParagraphProperties pPr = thisElement.PrependChild(new ParagraphProperties());
                    foreach (XmlNode ele in root.ChildNodes)
                        switch (ele.Name)
                        {
                            case "align":
                                switch (ele.InnerText)
                                {
                                    case "left":
                                        pPr.Append(new Justification { Val = JustificationValues.Left });
                                        break;
                                    case "right":
                                        pPr.Append(new Justification { Val = JustificationValues.Right });
                                        break;
                                    case "centre":
                                        pPr.Append(new Justification { Val = JustificationValues.Center });
                                        break;
                                    case "middle":
                                        pPr.Append(new Justification { Val = JustificationValues.Both });
                                        break;
                                }
                                break;
                        }
                    break;
                case "page_break":
                    thisElement = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
                    break;
                case "run":
                    thisElement = new Run();

                    // Add attributes
                    RunProperties rPr = thisElement.PrependChild(new RunProperties());
                    foreach (XmlNode ele in root.ChildNodes)
                    {
                        switch (ele.Name)
                        {
                            case ("#text"):
                                Text text = new Text()
                                {
                                    // Ensures spaces are not lost
                                    Space = SpaceProcessingModeValues.Preserve,
                                    Text = ele.InnerText
                                };
                                thisElement.AppendChild(text);
                                break;
                            case ("bold"):
                                rPr.Append(new Bold { Val = OnOffValue.FromBoolean(true) });
                                break;
                            case ("italic"):
                                rPr.Append(new Italic { Val = OnOffValue.FromBoolean(true) });
                                break;
                            case ("underlined"):
                                rPr.Append(new Underline { Val = UnderlineValues.Single });
                                break;
                            case ("strikethrough"):
                                rPr.Append(new Strike { Val = OnOffValue.FromBoolean(true) });
                                break;
                            case ("colour"):
                                rPr.Append(new Color { Val = ele.InnerText });
                                break;
                        }
                    }
                    break;
                case "heading":
                    thisElement = new Paragraph();

                    // Add attributes
                    int weight = 1;     //default heading size
                    ParagraphProperties hpPr = new ParagraphProperties();
                    thisElement.AppendChild(hpPr);
                    foreach (XmlNode ele in root.ChildNodes)
                    {
                        switch (ele.Name)
                        {
                            case "weight":
                                weight = int.Parse(ele.InnerText);
                                break;
                            case "align":
                                switch (ele.InnerText)
                                {
                                    case "left":
                                        hpPr.Append(new Justification { Val = JustificationValues.Left });
                                        break;
                                    case "right":
                                        hpPr.Append(new Justification { Val = JustificationValues.Right });
                                        break;
                                    case "centre":
                                        hpPr.Append(new Justification { Val = JustificationValues.Center });
                                        break;
                                    case "middle":
                                        hpPr.Append(new Justification { Val = JustificationValues.Both });
                                        break;
                                }
                                break;
                        }
                    }
                    // Add the relevant heading style to the paragraph    
                    AddStyleToParagraph(document, "vsmlHeading" + weight, "vsmlHeading" + weight, thisElement);
                    break;

                case "u_list_item":
                    thisElement = new Paragraph();

                    // Add attrbiutes
                    ParagraphProperties uListpPr = thisElement.PrependChild(new ParagraphProperties());
                    uListpPr.PrependChild(new NumberingProperties(
                        new NumberingLevelReference() { Val = 0 },
                        new NumberingId() { Val = 1 }));

                    foreach (XmlNode child in root.ChildNodes)
                        switch (child.Name)
                        {
                            case "align":
                                switch (child.InnerText)
                                {
                                    case "left":
                                        uListpPr.Append(new Justification { Val = JustificationValues.Left });
                                        break;
                                    case "right":
                                        uListpPr.Append(new Justification { Val = JustificationValues.Right });
                                        break;
                                    case "centre":
                                        uListpPr.Append(new Justification { Val = JustificationValues.Center });
                                        break;
                                    case "middle":
                                        uListpPr.Append(new Justification { Val = JustificationValues.Both });
                                        break;
                                }
                                break;
                        }
                    break;
                case "o_list_item":
                    thisElement = new Paragraph();
                    // Get the list id
                    int listId = 0;
                    foreach (XmlNode child in root.ChildNodes)
                        if (child.Name.Equals("list_id"))
                            listId = int.Parse(child.InnerText);
                    // Create a new numbering part for this id
                    //AddNewNumbering(document, listId);

                    // Add the new numbering part to the properties
                    ParagraphProperties oListpPr = thisElement.PrependChild(new ParagraphProperties());
                    oListpPr.PrependChild(new NumberingProperties(
                        new NumberingLevelReference() { Val = 0 },
                        new NumberingId() { Val = listId }));

                    foreach (XmlNode child in root.ChildNodes)
                    {
                        switch (child.Name)
                        {
                            case "align":
                                switch (child.InnerText)
                                {
                                    case "left":
                                        oListpPr.Append(new Justification { Val = JustificationValues.Left });
                                        break;
                                    case "right":
                                        oListpPr.Append(new Justification { Val = JustificationValues.Right });
                                        break;
                                    case "centre":
                                        oListpPr.Append(new Justification { Val = JustificationValues.Center });
                                        break;
                                    case "middle":
                                        oListpPr.Append(new Justification { Val = JustificationValues.Both });
                                        break;
                                }
                                break;
                        }
                    }
                    break;
                default:
                    Console.WriteLine("DEFAULT ERROR: " + type);
                    thisElement = new Paragraph(new Text("default_" + type));
                    break;
            }
            return thisElement;
        }

        /// <summary>
        /// Adds a given style to a given paragraph
        /// </summary>
        /// <param name="doc">The Document the paragraph is in</param>
        /// <param name="styleid">the styleid that is to be added </param>
        /// <param name="stylename">the name of the style to be added</param>
        /// <param name="p">The paragrapg the style is to be added to</param>
        private void AddStyleToParagraph(WordprocessingDocument doc, string styleid,
            string stylename, OpenXmlElement p)
        {
            // If the paragraph has no ParagraphProperties object, create one.
            if (p.Elements<ParagraphProperties>().Count() == 0)
                p.PrependChild<ParagraphProperties>(new ParagraphProperties());

            // Get the paragraph properties element of the paragraph.
            ParagraphProperties pPr = p.Elements<ParagraphProperties>().First();

            // Get the Styles part for this Document.
            StyleDefinitionsPart part = doc.MainDocumentPart.StyleDefinitionsPart;

            // If a ParagraphStyleId object doesn't exist, create one.
            if (pPr.ParagraphStyleId == null)
                pPr.ParagraphStyleId = new ParagraphStyleId();

            // Set the style of the paragraph.
            pPr.ParagraphStyleId.Val = styleid;
        }

        /// <summary>
        /// Adds a new numering part to the Document
        /// </summary>
        /// <param name="doc">The docuemnt the numbering part is to be added to</param>
        /// <param name="numberingId">The id of the new numbering part</param>
        private void AddNewNumbering(WordprocessingDocument doc, int numberingId)
        {
            Console.WriteLine("numberingId" + numberingId);
            NumberingDefinitionsPart numberingPart = doc.MainDocumentPart.NumberingDefinitionsPart;
            Numbering numbering = numberingPart.Numbering;
            Console.WriteLine(numbering.ChildElements.Count);

          

            numbering.AddChild(
               new AbstractNum(
                 new Level(
                   new NumberingFormat() { Val = NumberFormatValues.Bullet },
                   new LevelText() { Val = "·" }
                 )
                 { LevelIndex = 0 }
               )
               { AbstractNumberId = numberingId });
                            numbering.AddChild(
 new NumberingInstance(
                  new AbstractNumId() { Val = numberingId }
                )
                { NumberID = numberingId });

            
            numbering.Save(numberingPart);
            //NumberingDefinitionsPart numberingPart = doc.MainDocumentPart.NumberingDefinitionsPart;

            //// Add numbering element for bullet lists 
            //Numbering numbering = numberingPart.Numbering;
            //numbering.AppendChild(new AbstractNum(
            //            new Level(
            //                new NumberingFormat() { Val = NumberFormatValues.Decimal }, //type of char to mark list
            //                new LevelText() { Val = "%1peen." },
            //                new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel },
            //                new StartNumberingValue() { Val = 1 }
            //            )
            //            { LevelIndex = 0 }
            //        )
            //{ AbstractNumberId = numberingId });

            //numbering.AppendChild(


            //    new NumberingInstance(
            //    new AbstractNumId() { Val = numberingId })
            //    { NumberID = numberingId });

            //numbering.Save(numberingPart);
        }

        /// <summary>
        /// Creates a html file from the string input 
        /// </summary>
        /// <param name="input">The string to create the file from</param>
        /// <param name="title">The title of the document</param>
        /// <returns></returns>
        public string ToHtml()
        {
            // Return the html
            return InorderTraversalToCreateHtml(this.XmlDocument.ChildNodes.Item(0), "", "title");
        }

        /// <summary>
        /// Creates the html document by traversing the xml document 
        /// </summary>
        /// <param name="root">The current root node of the traversal</param>
        /// <param name="htmlString">The string representing the html document</param>
        /// <param name="title">The title of the html document</param>
        /// <returns></returns>
        private string InorderTraversalToCreateHtml(XmlNode root, string htmlString, string title)
        {
            List<string> explorableNodes = new List<string> {
                "doc",
                "section",
                "list",
                "paragraph",
                "heading",
                "o_list_item",
                "u_list_item",
                "run",
                "page_break"
            };

            Tuple<string, string> htmlTags = GetHtmlText(root, title);
            htmlString += htmlTags.Item1;
            string closingString = htmlTags.Item2;

            // Traversal - to add child tags
            if (root.ChildNodes.Count > 0)
                for (int i = 0; i < root.ChildNodes.Count; i++)
                    htmlString = InorderTraversalToCreateHtml(root.ChildNodes.Item(i), htmlString, title);

            // add closing tag
            htmlString += closingString;
            return htmlString;
        }

        /// <summary>
        /// Determins the relevant html tags to add to a html document for the 
        /// supplied xmlnode
        /// </summary>
        /// <param name="root">The xml element to create the html tags for</param>
        /// <returns>
        /// Tuple, with the first item being the opening tag and 
        /// the second being the closing tag
        /// </returns>
        private Tuple<string, string> GetHtmlText(XmlNode root, string title)
        {
            string htmlString = "";
            string closingString = "";
            switch (root.Name)
            {
                case "doc":
                    htmlString += "<!DOCTYPE html><html><head><title>" + title + "</title></head><body>";
                    closingString = "</body></html>";
                    break;
                case "section":
                    htmlString += "<section>";
                    closingString = "</section>";
                    break;
                case "list":
                    string type = "";
                    //find out list type
                    foreach (XmlNode child in root.ChildNodes)
                    {
                        if (child.Name.Equals("list_type"))
                            type = child.InnerText;
                        if (type.Equals("unordered"))
                            type = "u";
                        if (type.Equals("ordered"))
                            type = "o";
                        break;
                    }
                    htmlString += "<" + type + "l>";
                    closingString += "</" + type + "l>";
                    break;
                case "paragraph":
                    htmlString += "<p>";
                    closingString = "</p>";
                    break;
                case "heading":
                    string weight = "";
                    foreach (XmlNode child in root.ChildNodes)
                    {
                        if (child.Name.Equals("weight"))
                            weight = child.InnerText;
                    }
                    htmlString += "<h" + weight + ">";
                    closingString = "</h>";
                    break;
                case "o_list_item":
                case "u_list_item":
                    htmlString += "<li>";
                    closingString = "</li>";
                    break;
                case "run":
                    string text = "";
                    foreach (XmlNode child in root.ChildNodes)
                    {
                        switch (child.Name)
                        {
                            case "bold":
                                htmlString += "<b>";
                                closingString += "</b>";
                                break;
                            case "italic":
                                htmlString += "<i>";
                                closingString += "</i>";
                                break;
                            case "underlined":
                                htmlString += "<u>";
                                closingString += "</u>";
                                break;
                            case "strikethrough":
                                htmlString += "<del>";
                                closingString += "</del>";
                                break;
                            case "#text":
                                text = child.InnerText;
                                break;
                        }
                        htmlString += text;
                    }
                    break;

                case "page_break":
                    htmlString += "<br>";
                    break;
            }
            return new Tuple<string, string>(htmlString, closingString);
        }
    }
}

