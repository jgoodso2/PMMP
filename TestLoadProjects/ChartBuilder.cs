//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using DocumentFormat.OpenXml.Packaging;
//using System.IO;
//using System.Xml;
//using Ionic.Zip;

//namespace TestLoadProjects
//{
//    class ChartBuilder
//    {
//        public void EditGraph(PresentationPart pPart, DocumentFormat.OpenXml.Packaging.ChartPart[] charts)
//        {
//            String check = grab(pPart);
//            ChartPart chart = null;

//            for (int i = 0; i < charts.Count(); i++) //loop th
//            {
//                chart = charts[i];
//                if (check.ToLower().Equals(chart.Uri.ToString().ToLower()))
//                    break;
//            }
//            //chart now contains the chart-cache you are looking to edit.
//        }

//        private String grab(PresentationPart pPart)
//        {
//            //isolate chart
//            #region grab
//            var rids = RelsRID;
//            DocumentFormat.OpenXml.Wordprocessing.Drawing elem = null;
//            DocumentFormat.OpenXml.Drawing.Charts.ChartReference gd = null;
//            try
//            {
//                elem = pPart.NextSibling().Elements<Drawing>().ElementAt(0); //første forsøg på at finde vores graf
//            }
//            catch (Exception)
//            { //forsøg nummer 2
//                OpenXmlElement testE = pPart.NextSibling();
//                while (testE.LocalName.ToLower() != "drawing")
//                {
//                    testE = testE.NextSibling();
//                    for (int i = 0; i < testE.Elements().Count(); i++)
//                        if (testE.ElementAt(i).LocalName.ToLower() == "drawing") testE = testE.ElementAt(i);
//                }
//                elem = (DocumentFormat.OpenXml.Wordprocessing.Drawing)testE;
//            }
//            try
//            { //first try at grabbing graph data
//                gd = (DocumentFormat.OpenXml.Drawing.Charts.ChartReference)elem.Inline.Graphic.GraphicData.ElementAt(0);
//            }
//            catch (Exception)
//            { //second possible route
//                gd = (DocumentFormat.OpenXml.Drawing.Charts.ChartReference)elem.Anchor.Elements<Graphic>().First().Elements<GraphicData>().First().ElementAt(0);
//            }
//            var id = gd.Id;
//            String matchname = "/word/" + rids[id.ToString()]; //create filepath
//            #endregion
//            return matchname;
//        }

//        private Dictionary<String, String> RelsRIDToFile(string file)
//        {
//            String rels;
//            using (MemoryStream memory = new MemoryStream())
//            {
//                using (ZipFile zip = ZipFile.Read(file))
//                {
//                    ZipEntry e = zip["ppt/_rels/presentation.xml.rels"];
//                    e.Extract(memory);
//                }
//                using (StreamReader reader = new StreamReader(memory))
//                {
//                    memory.Seek(0, SeekOrigin.Begin);
//                    rels = reader.ReadToEnd();
//                }
//            }
//            XmlDataDocument xml = new XmlDataDocument();
//            xml.LoadXml(rels);
//            XmlNodeList xmlnode;
//            xmlnode = xml.GetElementsByTagName("Relationship");
//            Dictionary<String, String> result = new Dictionary<string, string>();
//            for (int i = 0; i < xmlnode.Count; i++)
//            {
//                var node = xmlnode[i];
//                var atts = node.Attributes;
//                String id = "";
//                String target = "";
//                for (int ii = 0; ii < atts.Count; ii++)
//                {
//                    var att = atts[ii];
//                    if (att.Name.ToLower() == "id") id = att.Value;
//                    if (att.Name.ToLower() == "target") target = att.Value;
//                }
//                result[id] = target;
//            }
//            return result;
//        }
//    }
//}
