using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using T1.Classes;
using System.Xml;
using System.Runtime.InteropServices;
using System.Collections;
using System.Globalization;
using System.Windows.Forms;

namespace T1.B1.Helpers
{
    public class Documents
    {
        static private Documents objDocumentsHelpers = null;

        private Documents()
        {
            
        }

        

        static public XmlDocument getBaseAmountsFromDocument(SAPbobsCOM.Documents oDocument)
        {

            /*
            Definicion del algoritmo.

            1. Busco las retenciones por articulos configuradas
                1.1 Si no hay definidas por artículos asumimos que todas aplican a todos
            2. Busco las retenciones que aplican al socio de negocio
                2.1 Si no hay retenciones definidas asumimos que fueron eliminadas
            3. Busco todas las retenciones
            4. Recorro las lineas del documento
                4.1 Construyo detalle de las lineas
                4.2 Construyo el detalle de impuestos
                4.3 Costruyo el detalle de retenciones
            5. Recorro los gastos adicionales
                5.1 Construyo el detalle de los gastos
                5.2 Construyo el detalle de los impuestos
                5.3 Construyo el detalle de retenciones
            6. Comparo con valores creados, si hay diferencia se convierte en cálculo manual de retencion.
            */




            XmlDocument oResult = null;
            XmlNode oLinesNode = null;
            XmlNode oTaxesNode = null;
            XmlNode oWithHoldingNode = null;
            XmlNode oExpensesNode = null;
            XmlNode oDocumentsNode = null;
            
            try
            {
                ///Get all BPs WithHolding Taxes Configured
                B1.Helpers.BusinessPartners.cacheHashWTCodesForBP(oDocument.CardCode, "B1.Helpers.Documents.CacheBP");
                ///Since Item and WT definitions are already cached form the beggining no need to search anymore
                

                ///Create XML Declartion
                XmlDocument doc = new XmlDocument();
                XmlNode docNode = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
                doc.AppendChild(docNode);

                ///Create Document Parent Node
                oDocumentsNode = doc.CreateElement("Document");
                XmlAttribute docEntryAtt = doc.CreateAttribute("DocEntry");
                docEntryAtt.Value = oDocument.DocEntry.ToString();
                oDocumentsNode.Attributes.Append(docEntryAtt);
                XmlAttribute docTypeAtt = doc.CreateAttribute("DocType");
                docTypeAtt.Value = oDocument.DocObjectCodeEx;
                oDocumentsNode.Attributes.Append(docTypeAtt);
                doc.AppendChild(oDocumentsNode);

                ///Build Line Information
                oLinesNode = doc.CreateElement("Lines");
                for(int i=1; i <= oDocument.Lines.Count; i++)
                {
                    oDocument.Lines.SetCurrentLine(i);

                    XmlNode xnlineNode = doc.CreateElement("Line");
                    XmlAttribute lineNume = doc.CreateAttribute("LineNum");
                    lineNume.Value = oDocument.Lines.LineNum.ToString();
                    xnlineNode.Attributes.Append(lineNume);

                    oLinesNode.AppendChild(xnlineNode);

                }
                doc.AppendChild(oLinesNode);

                /*


                XmlNode productNode = doc.CreateElement("product");
                XmlAttribute productAttribute = doc.CreateAttribute("id");
                productAttribute.Value = "01";
                productNode.Attributes.Append(productAttribute);
                productsNode.AppendChild(productNode);

                XmlNode nameNode = doc.CreateElement("Name");
                nameNode.AppendChild(doc.CreateTextNode("Java"));
                productNode.AppendChild(nameNode);
                XmlNode priceNode = doc.CreateElement("Price");
                priceNode.AppendChild(doc.CreateTextNode("Free"));
                productNode.AppendChild(priceNode);

                // Create and add another product node.
                productNode = doc.CreateElement("product");
                productAttribute = doc.CreateAttribute("id");
                productAttribute.Value = "02";
                productNode.Attributes.Append(productAttribute);
                productsNode.AppendChild(productNode);
                nameNode = doc.CreateElement("Name");
                nameNode.AppendChild(doc.CreateTextNode("C#"));
                productNode.AppendChild(nameNode);
                priceNode = doc.CreateElement("Price");
                priceNode.AppendChild(doc.CreateTextNode("Free"));
                productNode.AppendChild(priceNode);

    */

                return doc;

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);

            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);

            }

            return oResult;
        }


    }
}
