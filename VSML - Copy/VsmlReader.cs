using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace VSML
{
    class VsmlReader
    {

        public XmlDocument Document { get; set; }
        private readonly String schemaLocation = @"/vsmlschema.xsd";
        private readonly String headingPattern = @"(?<!\¬)\#[1-6]";
        private readonly String sectionStartPattern = @"^\/section((( header)|( footer))?)(( align=)((left|right|centre|middle)))?\/";
        private readonly String sectionEndPattern = @"^\\section\\";
        private readonly String orderedListPattern = @"^[0-9]+\.";
        private readonly String colourStartPattern = @"(?<!\¬)\/col=#[a-f|A-F|0-9]{6}\/";
        private readonly String colourEndPattern = @"(?<!\¬)\\col\\";
        private readonly String escapedColourStartPattern = @"\¬\/col=#[a-f|A-F|0-9]{6}\/";
        private readonly String escapedColourEndPattern = @"\¬\\col\\";
        private readonly Dictionary<string, char> values = new Dictionary<string, char>
        {
            { "bold",           '*' },
            { "italic",         '~' },
            { "underlined",     '_' },
            { "strikethrough",  '%' },
            { "superscript",    '^' },
            { "subscript",      '>' },
            { "heading",        '#' },
            { "unordered_list", '-' },
            { "ordered_list", '+' },
        };
        private readonly Dictionary<string, char> invalidMidLineValues = new Dictionary<string, char>
        {
            { "bold",           '*' },
            { "italic",         '~' },
            { "underlined",     '_' },
            { "strikethrough",  '%' },
            { "superscript",    '^' },
            { "subscript",      '>' },
        };


        /// <summary>
        /// Initialises a new VSML object to parse vsml files and provide outputs to other file types
        /// </summary>
        public VsmlReader()
        {
            this.Document = new XmlDocument();
        }


        /// <summary>
        /// Takes a vsml string input and creates a corresponding xml file
        /// The Xml is saved to the VSML object's attributes
        /// </summary>
        /// <param name="input">The vsml string to be converted</param>
        /// <param name="fileName">The location that the XML file will be stored in</param>
        /// <returns>
        /// Tuple, item 1 - wether the Document is valid or not, item 2 - the line number and 
        /// error message of any errors in the xml file. (not the input string)
        /// </returns>
        public Tuple<bool, Tuple<int, string>> VsmlToXml(String input)
        {
            this.Document.RemoveAll();
            XmlDocument document = Document;

            //Create the root node and add the namespace attribute to it
            XmlNode root = document.CreateElement("doc");
            ((XmlElement)root).SetAttribute("xmlns", "vsml");

            //Get an arary of all the lines in the input string
            String[] lines = input.Split(new[] { "\n" }, StringSplitOptions.None);

            Stack<XmlNode> parentNodes = new Stack<XmlNode>();
            String sectionAlign = "left";

            foreach (string aLine in lines)
            {
                String line = aLine;    //Make a copy so it can be edited
                line = line.Replace("\n", "").Replace("\r", "");    //remove extra chars 

                // If an empty line skip this item
                if (aLine.Length == 0)
                    line=" ";

                //Used if a section is just started/ended
                bool sectionStart = false;
                bool sectionEnd = false;

                bool hasText = true;    //represents if an element stores text
                XmlElement nextNode = null;

                // Get type of the line
                String type = this.DetermineType(line);

                // Depending on the type of the line, create the relevant XmlElement for that line
                switch (type)
                {
                    case "section_start":
                        {
                            sectionStart = true;
                            nextNode = document.CreateElement("section");
                            hasText = false;

                            // Get the attrbiutes from the tag if they exist
                            String str = line.Replace("/", "").Replace("\\", "");   // remove tag characters
                                                                                    // Valid directions for alignment
                            List<String> directions = new List<String>(new String[] { "left", "right", "middle", "centre" });
                            // Split the section tag on spaces (will split attributes)
                            List<String> attributes = new List<String>(str.Split(' '));

                            // For each attribute 
                            foreach (string att in attributes)
                            {
                                //Split it into attibute name, and value
                                List<String> attribute2 = new List<String>(att.Split('='));

                                bool align = false;
                                int alignInt = -1;

                                for (int i = 0; i < attribute2.Count; i++)
                                {
                                    // Search the list of attributes for align attribute
                                    // once found mark align as true
                                    // once found = true, the next attribute will be the value

                                    if (attribute2[i].Equals("align"))
                                    {
                                        align = true;
                                        alignInt = i;   //gets the index of the align attribute
                                    }
                                    if (align & i != alignInt & directions.Contains(attribute2[i]))
                                    {
                                        sectionAlign = attribute2[i];
                                        XmlElement alignment = document.CreateElement("align");
                                        alignment.InnerText = sectionAlign;
                                        nextNode.AppendChild(alignment);
                                        break;
                                    }
                                }
                            }
                            break;
                        }
                    case "section_end":
                        {
                            sectionEnd = true;
                            hasText = false;
                            sectionAlign = "left";  //reset in case there was an alignment attribute
                            break;
                        }
                    case "page_break":
                        {
                            nextNode = document.CreateElement("page_break");
                            hasText = false;
                            break;
                        }
                    case "unordered_list_item":
                        {
                            nextNode = document.CreateElement("u_list_item");
                            line = line.Remove(0, 1);
                            if (line.Length == 0)
                                line += "*";
                            // Check if there is a element on the stack that this could inherit alignment from 
                            if (parentNodes.Count > 0)
                                foreach (XmlNode ele in parentNodes.Peek().ChildNodes)
                                    if (ele.Name.Equals("align"))
                                    {
                                        XmlElement alignment = document.CreateElement("align");
                                        alignment.InnerText = ele.InnerText;
                                        nextNode.AppendChild(alignment);
                                    }
                            break;
                        }
                    case "ordered_list_item":
                        {
                            nextNode = document.CreateElement("o_list_item");

                            //Remove list tag from start of the line
                            line = Regex.Replace(line, orderedListPattern, "");
                            if (line.Length == 0)
                                line += "*";
                            //line = "+" + line;  //prepend ordered list modifier symbol - allows empty line
                            //Check if there is a element on the stack that this could inherit alignment from 
                            if (parentNodes.Count > 0)
                                foreach (XmlNode ele in parentNodes.Peek().ChildNodes)
                                    if (ele.Name.Equals("align"))
                                    {
                                        XmlElement alignment = document.CreateElement("align");
                                        alignment.InnerText = ele.InnerText;
                                        nextNode.AppendChild(alignment);
                                    }
                            break;
                        }
                    case "heading":
                        {
                            nextNode = document.CreateElement("heading");
                            XmlElement weight = document.CreateElement("weight");
                            // Add the weight attribute 
                            weight.InnerText = line.ElementAt(1).ToString();
                            nextNode.AppendChild(weight);
                            line = line.Substring(2);
                            // Check if there is a element on the stack that this could inherit alignment from 
                            if (parentNodes.Count > 0)
                                foreach (XmlNode ele in parentNodes.Peek().ChildNodes)
                                    if (ele.Name.Equals("align"))
                                    {
                                        XmlElement alignment = document.CreateElement("align");
                                        alignment.InnerText = ele.InnerText;
                                        nextNode.AppendChild(alignment);
                                    }
                            if (line.Length == 0)
                                hasText = false;
                            break; 
                        }
                    case "paragraph":
                        {
                            nextNode = document.CreateElement("paragraph");

                            // Check if there is a element on the stack that this could inherit alignment from 
                            if (parentNodes.Count > 0)
                                foreach (XmlNode ele in parentNodes.Peek().ChildNodes)
                                    if (ele.Name.Equals("align"))
                                    {
                                        XmlElement alignment = document.CreateElement("align");
                                        alignment.InnerText = ele.InnerText;
                                        nextNode.AppendChild(alignment);
                                    }
                            break;
                        }
                }
                // Add any text to the element if neccessary
                if (hasText)
                {
                    // Get the details for the run and use them to create a list of runs 
                    // Add each run form the list to the element
                    List<String>[] runDetails = CreateStringList(line);
                    foreach (XmlNode ele in ReturnTextElement(line, runDetails, document))
                        nextNode.AppendChild(ele);
                }

                // If just making a section - add it to the parent nodes.
                if (sectionStart)
                {
                    parentNodes.Push(nextNode);
                    continue;
                }
                // If the section is being ended - there will be nothing to add
                if (sectionEnd)
                {
                    XmlNode toAdd;
                    if (parentNodes.Count > 0)
                        toAdd = parentNodes.Pop();
                    else
                        toAdd = null;

                    if (parentNodes.Count > 0)
                        // Append newNode to this parent
                        parentNodes.Peek().AppendChild(toAdd);
                    // Else add it to the root
                    else
                        if (toAdd != null)
                        root.AppendChild(toAdd);
                    continue;
                }
                // If there is a parent node
                if (parentNodes.Count > 0)
                    // Append newNode to this parent
                    parentNodes.Peek().AppendChild(nextNode);
                // Else add it to the root
                else
                    root.AppendChild(nextNode);
            }

            // If there is any extra nodes that have been added, then add them
            // (if this is used the Document will be invalid)
            while (parentNodes.Count > 0)
                root.AppendChild(parentNodes.Pop());

            // Add the newly made Document to the Document class attribute
            this.Document.AppendChild(root);
            this.Document = AddListTags(this.Document);

            // Save this doc to a file (not used as output)
            this.Document.Save(Path.Combine(Environment.CurrentDirectory + "latestDoc.xml"));

            //Verify and return the validity of the newly made xml file
            XmlVerifyer very =
                new XmlVerifyer(Path.Combine(Environment.CurrentDirectory + schemaLocation),
                Path.Combine(Environment.CurrentDirectory + "latestDoc.xml"));
            var result = very.VerifyDocument();

            return result;

        }

        /// <summary>
        /// Adds the neccessary list tags to the xml Document 
        /// </summary>
        /// <param name="doc">The xml Document to add list tags too</param>
        /// <returns>
        /// The same xml Document, but with extra list nodes added 
        /// around each set of list items
        /// </returns>
        private XmlDocument AddListTags(XmlDocument doc)
        {
            // Take the xml Document as a line of text and 
            // using regex insert the list tags
            string text = doc.OuterXml;

            text = Regex.Replace(text, @"<u_list_item \/>",
                "<u_list_item></u_list_item>");
            text = Regex.Replace(text, @"<o_list_item \/>",
                "<o_list_item></o_list_item>");
            text = Regex.Replace(text, @"(?<!(<\/o_list_item>))<o_list_item>",
                "<list><list_type>ordered</list_type><o_list_item>");
            text = Regex.Replace(text, @"(?<!(<\/u_list_item>))<u_list_item>",
                "<list><list_type>unordered</list_type><u_list_item>");
            text = Regex.Replace(text, @"<\/o_list_item>(?!<o_list_item>)",
                "</o_list_item></list>");
            text = Regex.Replace(text, @"<\/u_list_item>(?!<u_list_item>)",
                "</u_list_item></list>");

            //Recreate the Document as an XmlDocument object, 
            // add listIds to the Document
            doc.LoadXml(text);
            NumberLists(doc);
            return doc;
        }

        /// <summary>
        /// Adds a unique listId to each list in the Document
        /// </summary>
        /// <param name="doc">The Document to number the lists in</param>
        private void NumberLists(XmlDocument doc)
        {
            // Get a list of all the XmlElements that are named list
            XmlNodeList allNodes = doc.GetElementsByTagName("list");
            Console.WriteLine("This many lists: " + allNodes.Count);
            // For every list, add a listId to the list item in the list
            for (int i = 0; i < allNodes.Count; i++)
            {
                XmlNode currentList = allNodes.Item(i);
                foreach (XmlNode child in currentList.ChildNodes)
                {
                    if (child.Name.Equals("o_list_item"))
                    {
                        XmlNode listId = doc.CreateElement("list_id", "vsml");
                        listId.PrependChild(doc.CreateTextNode((i + 2).ToString())); // Add two so starts from 2 - list id 1 is bullet points 0 not used
                        child.PrependChild(listId);
                    }
                    if (child.Name.Equals("u_list_item"))
                    {
                        XmlNode listId = doc.CreateElement("list_id", "vsml");
                        listId.PrependChild(doc.CreateTextNode("1")); // Add two so starts from 2 - list id 1 is bullet points 0 not used
                        child.PrependChild(listId);
                    }
                }
            }
        }

        /// <summary>
        /// Takes a string input and creates an array of strings for each character
        /// in the given input. The array stores a string code representing the effects 
        /// that are applied to the character at the corresponding index of the string input
        /// A second string[] is returned in the list, representing the colour at each index
        /// A second array stores the colour for each character
        /// </summary>
        /// <param name="line">The line to return the characteristics for</param>
        /// <returns>
        /// List of string arrays, each array represents a different set of effects on the text
        /// The first array corresponds to the effects on the text e.g. bold, italic etc.
        /// The second corresponds to the colour of the text
        /// A List is used so further attributes of the text could be added in the future
        /// </returns>
        private List<String>[] CreateStringList(String line)
        {
            List<String> colours = new List<String>();
            Regex start = new Regex(colourStartPattern);
            Regex end = new Regex(colourEndPattern);
            bool foundAll = false;

            // Insert black as a default colour
            for (int i = 0; i < line.Length; i++)
                colours.Insert(i, "000000");

            int colourStartingIndex = 0;
            while (!foundAll)
            {
                int colourRunStart = -1;
                int colourRunEnd = -1;

                // If there is a match for a start and end tag
                Match colourStart = start.Match(line, colourStartingIndex);
                Match colourEnd = end.Match(line, colourStartingIndex);
                if (colourStart.Success & colourEnd.Success)
                {
                    // Get the start and end indices of the colour tags
                    colourRunEnd = colourEnd.Index;
                    colourRunStart = colourStart.Index;
                    // For the chars between these indices, update the colour to the 
                    // specified colour
                    for (int i = colourRunStart; i <= colourRunEnd; i++)
                        colours.Insert(i, colourStart.Value.Substring(6, 6));
                    // Update the start location
                    colourStartingIndex = colourRunEnd;
                }
                // If there is no starting tag - or there is no space for an ending tag left
                // - or there is no ending tag
                if (!start.IsMatch(line, colourStartingIndex) |
                    line.Length - colourRunEnd < 12 | !colourEnd.Success)
                    foundAll = true;
            }

            List<String> modifiers = new List<String>();
            // Depending on the value of these, the char will be determined to be 
            // bold or italic etc
            bool bold = false;
            bool italic = false;
            bool underlined = false;
            bool strikethrough = false;

            // Check for single character modifiers (* or - etc)
            // For every letter 
            for (int i = 0; i < line.Length; i++)
            {
                // assume the char is escaped
                bool escaped = true;
                // cannot escape first char
                if (i == 0)
                    escaped = false;
                //if the char behind is not a slash cannot be escaped
                else if (line.ElementAt(i - 1) != '¬')
                    escaped = false;

                // if a modififier character found
                // toggle the status of the variable 
                if (!escaped)
                {
                    char currnetChar = line.ElementAt(i);
                    if (currnetChar == values["bold"])
                        bold = !bold;
                    if (currnetChar == values["italic"])
                        italic = !italic;
                    if (currnetChar == values["underlined"])
                        underlined = !underlined;
                    if (currnetChar == values["strikethrough"])
                        strikethrough = !strikethrough;
                }

                // Add to the list the relevant values depending on the effects 
                // applied to the character
                List<char> toInsert = new List<char>();
                if (bold)
                    toInsert.Add(values["bold"]);
                if (italic)
                    toInsert.Add(values["italic"]);
                if (underlined)
                    toInsert.Add(values["underlined"]);
                if (strikethrough)
                    toInsert.Add(values["strikethrough"]);

                toInsert.Sort();
                modifiers.Insert(i, String.Join("", toInsert));
            }
            return new List<String>[] { modifiers, colours };
        }

        /// <summary>
        /// Determines the type of the given line
        /// </summary>
        /// <param name="line">The line to determine the type of</param>
        /// <returns>The name of the type of line</returns>
        private String DetermineType(String line)
        {
            // Use simple equivalence or regular expressions to determine the type
            if (line.Length == 0)
                return null;
            
            if (line.Equals("/break/"))
                return "page_break";

            Regex rx = new Regex(headingPattern);
            if (rx.IsMatch(line))
                return "heading";

            if (line.ElementAt(0) == values["unordered_list"])
                return "unordered_list_item";

            rx = new Regex(sectionStartPattern);
            if (rx.IsMatch(line))
                return "section_start";

            rx = new Regex(sectionEndPattern);
            if (rx.IsMatch(line))
                return "section_end";

            rx = new Regex(orderedListPattern);
            if (rx.IsMatch(line))
                return "ordered_list_item";

            return "paragraph";
        }

        /// <summary>
        /// Takes a line, and the list of arrays representing the attributes of each
        /// character in the supplied line, and returns the equivalent list of run 
        /// that correspond to the input
        /// </summary>
        /// <param name="line">The line to make the xml nodes from</param>
        /// <param name="stringValues">the attributes of each character in the input</param>
        /// <param name="doc">The doc that the nodes will be a part of</param>
        /// <returns></returns>
        private List<XmlNode> ReturnTextElement(string line,
            List<string>[] stringValues, XmlDocument doc)
        {
            List<XmlNode> runs = new List<XmlNode>();

            // Remove extra chars from the line
            var removedChars = RemoveModifierCharacters(line, stringValues);
            line = removedChars.Item1;
            stringValues = removedChars.Item2;

            // Find every consecutive series of characters within the line, 
            // that have the same effects applied to them

            //start and end points of a run of matching characters
            int start = 0;
            int end = 0;

            for (int i = 0; i < line.Length; i++)
            {
                bool buildRun = false;
                // If end of the line
                if (i == line.Length - 1)
                    buildRun = true;
                else
                {
                    // If the colour, and all effects of the next character in 
                    // the string are the same, increment the end of the run
                    if ((stringValues[0][i].Equals(stringValues[0][i + 1]))
                        & stringValues[1][i].Equals(stringValues[1][i + 1]))
                        end += 1;
                    // If not then a run needs to be made
                    else
                        buildRun = true;
                }

                if (buildRun)
                {
                    XmlNode run = doc.CreateElement("run");

                    // Add the effects to the the run node
                    foreach (KeyValuePair<string, char> entry in values)
                        if (stringValues[0][i].Contains(entry.Value))
                        {
                            XmlNode att = doc.CreateElement(entry.Key);
                            run.AppendChild(att);
                        }
                    // Add the colour to the run
                    XmlNode colour = doc.CreateElement("colour");
                    XmlText colourCode = doc.CreateTextNode("#" + stringValues[1][i]);
                    colour.AppendChild(colourCode);
                    run.AppendChild(colour);

                    // Add the text to the run
                    String text = line.Substring(start, end - start + 1);
                    XmlText textPart = doc.CreateTextNode(text);
                    run.AppendChild(textPart);

                    // Add the run to the return list, and update start and end
                    // points for the next run
                    runs.Add(run);
                    start = i + 1;
                    end = start;
                }
            }
            return runs;
        }

        private Tuple<string, List<string>[]> RemoveModifierCharacters(string line,
            List<string>[] stringValues)
        {
            List<int> indicesToRemove = new List<int>(); //indices of positions to be removed
            Console.WriteLine(">"+line+"<");
            //check for the 0th char - cannot be escaped
            char charAtI = line.ElementAt(0);

            //Remove modifier if first char of string
            if (invalidMidLineValues.Values.Contains(charAtI))
                indicesToRemove.Add(0);

            for (int i = 1; i < line.Length; i++)
            {
                // If the ith character is a modifier character (*, ~, etc) and 
                // not preceded escape character add index to be removed
                charAtI = line.ElementAt(i);
                if (invalidMidLineValues.Values.Contains(charAtI) & line.ElementAt(i - 1) != '¬')
                    indicesToRemove.Add(i);
                if (invalidMidLineValues.Values.Contains(charAtI) & line.ElementAt(i - 1) == '¬')
                    indicesToRemove.Add(i - 1);
                //if escaped list item
                if (charAtI == '-' & line.ElementAt(i - 1) == '¬')
                    indicesToRemove.Add(i - 1);
                Match match = Regex.Match(line, @"^[0-9]+\.");
                if (match.Index == 0 & line.ElementAt(i - 1) == '¬')
                    indicesToRemove.Add(i - 1);
            }

            Regex escapedStartColourMatcher = new Regex(escapedColourStartPattern);
            Regex escapedEndColourMatcher = new Regex(escapedColourEndPattern);

            foreach (Match match in escapedStartColourMatcher.Matches(line))
                indicesToRemove.Add(match.Index);
            foreach (Match match in escapedEndColourMatcher.Matches(line))
                indicesToRemove.Add(match.Index);

            // Remove only valid colour tags
            Regex colourStartRx = new Regex(colourStartPattern);
            Regex colourEndRx = new Regex(colourEndPattern);
            int colourStartingIndex = 0;
            bool foundAll = false;
            while (!foundAll)
            {
                int colourRunStart = -1;
                int colourRunEnd = -1;

                // If there is a match for start and end tags
                Match colourStart = colourStartRx.Match(line, colourStartingIndex);
                Match colourEnd = colourEndRx.Match(line, colourStartingIndex);
                if (colourStart.Success & colourEnd.Success)
                {
                    // Get the index of the start of the start and end tags
                    colourRunEnd = colourEnd.Index;
                    colourRunStart = colourStart.Index;

                    // Add the indicies of all characters in the tags
                    // 12 and 4 length of tags, so count forwards that many chars to remove all
                    for (int i = 0; i <= 12; i++)
                        indicesToRemove.Add(colourRunStart + i);

                    for (int i = 0; i <= 4; i++)
                        indicesToRemove.Add(colourRunEnd + i);
                    colourStartingIndex = colourRunEnd;
                }
                // If no start tag, or no end tag, or no space for a start tag
                if (!colourStartRx.IsMatch(line, colourStartingIndex) |
                    line.Length - colourRunEnd < 12 | !colourEnd.Success)
                    foundAll = true;
            }
            // Remove the duplicates from the list
            List<int> nonDuplicates = new List<int>();
            foreach (int index in indicesToRemove)
                if (!nonDuplicates.Contains(index))
                    nonDuplicates.Add(index);
            indicesToRemove = nonDuplicates;

            // Sort the list - so indices are ascending order
            indicesToRemove.Sort();

            // Remove the list - so when removing items, 
            // you dont alter the index of other items
            indicesToRemove.Reverse();

            // Remove the items at the index in all relevant lists / lines
            foreach (int index in indicesToRemove)
            {
                line = line.Remove(index, 1);
                stringValues[0].RemoveAt(index);
                stringValues[1].RemoveAt(index);
            }
            return new Tuple<string, List<string>[]>(line, stringValues);
        }
    }
}
