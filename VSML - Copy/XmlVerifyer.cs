using System;
using System.Xml;
using System.Xml.Schema;

namespace VSML
{
    class XmlVerifyer
    {
        private string XsdLocation { get; }
        private string XmlLocation { get; }
        private bool ValidDocument;
        private Tuple<int, string> error;

        public XmlVerifyer(string xsdLocation, string xmlLocation)
        {
            this.XsdLocation = xsdLocation;
            this.XmlLocation = xmlLocation;
            this.ValidDocument = false;
        }

        public Tuple<bool, Tuple<int, String>> VerifyDocument()
        {
            this.ValidDocument = true;

            XmlReaderSettings settings = new XmlReaderSettings();
            settings.Schemas.Add("vsml", XsdLocation);
            settings.DtdProcessing = DtdProcessing.Parse;
            settings.CloseInput = true;
            settings.ValidationEventHandler += new ValidationEventHandler(ValidationEventHandler);
            settings.ValidationType = ValidationType.Schema;
            settings.ValidationFlags =
                XmlSchemaValidationFlags.ReportValidationWarnings |
                XmlSchemaValidationFlags.ProcessIdentityConstraints |
                XmlSchemaValidationFlags.ProcessInlineSchema |
                XmlSchemaValidationFlags.ProcessSchemaLocation;
            XmlReader reader = XmlReader.Create(XmlLocation, settings);
            XmlDocument document = new XmlDocument();
            document.Load(reader);
            reader.Dispose();
            return new Tuple<bool, Tuple<int, string>>(ValidDocument, error);
        }

        private void ValidationEventHandler(object sender, ValidationEventArgs e)
        {
            Console.WriteLine(sender.GetType());
            this.ValidDocument = false;
            this.error = new Tuple<int, string>(e.Exception.LineNumber, e.Message);
        }
    }
}
