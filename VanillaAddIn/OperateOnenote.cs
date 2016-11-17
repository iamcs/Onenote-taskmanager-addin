using Microsoft.Office.Interop.OneNote;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Application = Microsoft.Office.Interop.OneNote.Application;
using System.IO;
using System.Collections;
using System.Threading;

namespace MyApplication.VanillaAddIn
{

    class OperateOnenote
    {
        protected Application OneNoteApplication
        { get; set; }
        private const string NS = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        public static XmlNamespaceManager onenotenamespace;
        public static XmlDocument hierxml;
        public OperateOnenote(Application onenoteapplication)
        {
            this.OneNoteApplication = onenoteapplication;
            hierxml = gethier();
            onenotenamespace = GetNSManager(hierxml.NameTable);
            hierxml = removerecyclebin(hierxml);
        }

        /// <summary>
        /// Returns the namespace manager from the passed xml name table.
        /// </summary>
        /// <param name="nameTable">Name table of the xml document.</param>
        /// <returns>Returns the namespace manager.</returns>
        public XmlNamespaceManager GetNSManager(XmlNameTable nameTable)
        {
            var nsManager = new XmlNamespaceManager(nameTable);
            try
            {
                nsManager.AddNamespace("one", NS);
            }
            catch (Exception e)
            {
                throw new ApplicationException("Error in GetNSManager: " + e.Message, e);
            }
            return nsManager;
        }

        //public 

        public XmlDocument GetPageContent(string pageId)
        {
            var doc = new XmlDocument();
            try
            {
                string pageInfo;
                OneNoteApplication.GetPageContent(pageId, out pageInfo, PageInfo.piAll);
                doc.LoadXml(pageInfo);
            }
            catch (Exception e)
            {
                throw new ApplicationException("Error in GetPageContent: " + e.Message + pageId.ToString(), e);
            }
            return doc;
        }

        internal void XmltoTreeView(TreeView treeView1, XmlDocument dom)
        {
            treeView1.Nodes.Clear();
            treeView1.Nodes.Add(new TreeNode(dom.DocumentElement.Name));
            TreeNode tNode = new TreeNode();
            tNode = treeView1.Nodes[0];

            // SECTION 3. Populate the TreeView with the DOM nodes.
            AddNode(dom.DocumentElement, tNode);
            treeView1.ExpandAll();
        }

        private void AddNode(XmlNode inXmlNode, TreeNode inTreeNode)
        {
            XmlNode xNode;
            TreeNode tNode;
            XmlNodeList nodeList;
            int i;

            // Loop through the XML nodes until the leaf is reached.
            // Add the nodes to the TreeView during the looping process.
            if (inXmlNode.HasChildNodes)
            {
                nodeList = inXmlNode.ChildNodes;
                for (i = 0; i <= nodeList.Count - 1; i++)
                {
                    xNode = inXmlNode.ChildNodes[i];
                    inTreeNode.Nodes.Add(new TreeNode(xNode.Name));
                    tNode = inTreeNode.Nodes[i];
                    AddNode(xNode, tNode);
                }
            }
            else
            {
                // Here you need to pull the data from the XmlNode based on the
                // type of node, whether attribute values are required, and so forth.
                inTreeNode.Text = (inXmlNode.OuterXml).Trim();
            }
        }

        public void AddPageContent(string pageId, string content, int yPos = 80, int width = 520)//add content to the end of the 1st outline,not a independent line
        {
            var doc = GetPageContent(pageId);
            var outlinedoc = GetOutlineContent(pageId);

            string newPageContent = doc.InnerXml.Replace(outlinedoc, outlinedoc + "\n" + content);
            OneNoteApplication.UpdatePageContent(newPageContent, DateTime.MinValue);
        }

        public string GetOutlineContent(string pageId)//get 1st line of the 1st outline 
        {
            var doc = GetPageContent(pageId);
            string INpage = doc.SelectSingleNode("*/one:Outline//one:T", onenotenamespace).InnerText;
            return INpage;
        }

        public XmlDocument gethier()//get xml from top to pages
        {
            XmlDocument xm = new XmlDocument();
            string xml;
            this.OneNoteApplication.GetHierarchy(null, HierarchyScope.hsPages, out xml);
            xm.LoadXml(xml);
            return xm;
        }
        public string getsectionid(string sectionname)//get section's id by name
        {
            string id;
            id = hierxml.SelectSingleNode("//one:Section[@name='" + sectionname + "']", onenotenamespace).Attributes.GetNamedItem("ID").Value;
            return id;
        }
        public string getpageid(string pagename)//get page's id by name
        {
            string id = null;
            XmlNode pagenode = hierxml.SelectSingleNode("//one:Page[@name='" + pagename + "']", onenotenamespace);
            if (pagenode != null)
            {
                id = pagenode.Attributes.GetNamedItem("ID").Value;
            }            
            return id;
        }
        public string get1stoutlineid(string pageid)//get page's first outline id use pageid
        {
            string id;
            id = hierxml.SelectSingleNode("//one:Outline", onenotenamespace).Attributes.GetNamedItem("ID").Value;
            return id;
        }
        public string getoeid(XmlNode oenode)//get id of specific oe node
        {
            string id;
            id = oenode.Attributes.GetNamedItem("objectID").Value;
            return id;
        }

        public XmlNode createoeline(string content,XmlDocument xml)//create an oe element, has some cdata content, no tag
        {
            XmlElement oenode = xml.CreateElement("one", "OE", NS);
            XmlElement tnode = xml.CreateElement("one", "T", NS);
            XmlCDataSection cdata = xml.CreateCDataSection(content);

            tnode.AppendChild(cdata);
            oenode.AppendChild(tnode);
            return oenode;
        }
        /// <summary>
        /// create a tagdef for pagexml
        /// pagexml:without tagdef -> pagexml:withtagdef
        /// </summary>
        /// <param name="xml"></param>
        public void addtagdef(XmlDocument xml)
        {
            XmlElement tagdef = xml.CreateElement("one", "TagDef", NS);//<one:TagDef index="0" name="待办事项" type="0" symbol="3" />
            tagdef.SetAttribute("index", "0");
            tagdef.SetAttribute("name", "待办事项");
            tagdef.SetAttribute("type", "0");
            tagdef.SetAttribute("symbol", "3");
            XmlNode pagenode = xml.SelectSingleNode("//one:Page", onenotenamespace);
            pagenode.PrependChild(tagdef);
        }
        /// <summary>
        /// create a oeline with tag
        /// pagexml , content -> page with a new tag line, the oe tag line node
        /// </summary>
        /// <param name="content"></param>
        /// <param name="xml"></param>
        /// <returns></returns>
        public XmlNode createtagline(string content, XmlDocument xml)//create an oe element, has some cdata content, has tag
        {

            try
            {
                XmlNode tagdefnode = xml.SelectSingleNode("//one:TagDef[@index=\"0\"]", onenotenamespace);
                if (tagdefnode == null)//[@name = \"个人\"]
                {
                    addtagdef(xml);
                }
            }
            catch (Exception e)
            {               
                throw e;
            }
            XmlElement oenode = xml.CreateElement("one", "OE", NS);
            XmlElement tnode = xml.CreateElement("one", "T", NS);
            XmlElement tagnode = xml.CreateElement("one", "Tag", NS);
            tagnode.SetAttribute("index", "0");
            XmlCDataSection cdata = xml.CreateCDataSection(content);

            tnode.AppendChild(cdata);
            oenode.AppendChild(tnode);
            oenode.PrependChild(tagnode);
            return oenode;
        }
        public void addlinkline(XmlNode targetnode, ref XmlDocument pagexml, string targetpageid)//add targetnode to the last line of pagexml
        {
            string link;
            OneNoteApplication.GetHyperlinkToObject(targetpageid, getoeid(targetnode), out link);
            string text = targetnode.SelectSingleNode("one:T", onenotenamespace).InnerText;
            AddPageline(pagexml, "<a href = \"" + link + "\"> " + text + " </ a >",linetype.oeline);
        }
        public enum linetype
        {
            tagline,
            oeline,
        }
        public XmlNode creatoutline(XmlDocument xml)
        {
            var pagenode = xml.SelectSingleNode("//one:Page", onenotenamespace);
            XmlElement outlinenode = xml.CreateElement("one", "Outline", NS);
            XmlElement positionnode = xml.CreateElement("one", "Position", NS);
            positionnode.SetAttribute("x", "36");
            positionnode.SetAttribute("y", "100");
            positionnode.SetAttribute("z", "0");
            XmlElement sizenode = xml.CreateElement("one", "Size", NS);
            sizenode.SetAttribute("width", "80");
            sizenode.SetAttribute("height", "15");
            XmlElement oecnode = xml.CreateElement("one", "OEChildren", NS);

            pagenode.AppendChild(outlinenode);
            outlinenode.AppendChild(positionnode);
            outlinenode.AppendChild(sizenode);
            outlinenode.AppendChild(oecnode);
            return oecnode;
        }
        public XmlNode AddPageline(XmlDocument pagexml, string content ,linetype type)//add line to the end of the 1st outline,a independent line,in the first class
        {
            var outlinenode = pagexml.SelectSingleNode("//one:OEChildren", onenotenamespace);
            if (outlinenode == null)
            {
                outlinenode = creatoutline(pagexml);
            }
            XmlNode oenode= null;
            switch (type)
            {
                case linetype.oeline:
                    oenode = createoeline(content, pagexml);
                    break;
                case linetype.tagline:
                    oenode = createtagline(content, pagexml);
                    break;
                default:
                    break;
            }
            outlinenode.AppendChild(oenode);
            return oenode;
            //OneNoteApplication.UpdatePageContent(doc.InnerXml, DateTime.MinValue);
        }
        public void creattable(int column,int row, ArrayList[] content, XmlNode fathernode,XmlDocument xml)//construct a xml table
        {
            /***if ((content.GetLength(0) != row)||(content.GetLength(1)!= column))
            {
                MessageBox.Show("data can't match with table!");
                return; }***/
            XmlElement oetop = xml.CreateElement("one", "OE", NS);
            fathernode.AppendChild(oetop);
            XmlElement table = xml.CreateElement("one", "Table", NS);
            table.SetAttribute("bordersVisible", "true");//bordersVisible="true"
            oetop.AppendChild(table);
            XmlElement columns = xml.CreateElement("one", "Columns", NS);
            
            for (int n = column;n>=1;n--)
            {
                XmlElement columnnode = xml.CreateElement("one", "Column", NS);
                columnnode.SetAttribute("index",(n-1).ToString());
                columnnode.SetAttribute("width", "37");
                columns.PrependChild(columnnode);
            }
            for (int n = row; n >= 1; n--)//xmlnode[row][column]
            {
                XmlElement rownode = xml.CreateElement("one", "Row", NS);

                for (int m = column; m >= 1; m--)//week = m day = m+7*(n-1)+1 - startday
                {
                    //initialize
                    int i = 0;
                    DateTime now = DateTime.Now;
                    int cellnum = m + 7 * (n - 1);//count of cell
                    XmlElement cellnode = xml.CreateElement("one", "Cell", NS);

                    //start from what day
                    int startday = (int)now.AddDays(1 - now.Day).DayOfWeek;
                    if (startday == 0)
                    { startday = 7; }

                    //month day of cell
                    int cellday = m + 7 * (n - 1) + 1 - startday;                                       

                    //set color for the cell of today, else do nothing
                    //shadingColor="#B2A1C7"
                    if (now.Day == cellday)
                    {
                        cellnode.SetAttribute("shadingColor", "#B2A1C7");
                    }


                    XmlElement oecnode = xml.CreateElement("one", "OEChildren", NS);
                    
                   
                    
                    if (cellnum- startday + 1 <= DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month) & cellnum >= startday)
                    {
                        oecnode.AppendChild(createoeline(cellday.ToString() + "号-周" + m.ToString(), xml));
                        foreach (XmlNode x in content[0])//day
                        {
                            XmlNode node = x.Clone();
                            oecnode.AppendChild(node);
                        }
                        i = 0;
                        foreach (int x in content[2])//week
                        {
                            if (x == m)
                            {
                                XmlNode node = ((XmlNode)(content[1])[i]).Clone();
                                oecnode.AppendChild(node);
                            }
                            i++;
                        }
                        i = 0;
                        foreach (int x in content[4])//month
                        {
                            
                            if (x == cellnum - startday+1)
                            {
                                oecnode.AppendChild(((XmlNode)(content[3])[i]));
                            }
                            i++;
                        }
                        i = 0;
                        foreach (int x in content[6])//year
                        {
                            
                            if (x == DateTime.Now.Month & n == 1 & m == startday)
                            {
                                oecnode.AppendChild(((XmlNode)(content[5])[i]));
                            }
                            i++;
                        }
                        i = 0;
                        foreach (DateTime x in content[8])//specific
                        {
                            
                            if (x.Month == DateTime.Now.Month & x.Day == cellnum - startday + 1)
                            {
                                oecnode.AppendChild((XmlNode)(content[7])[i]);
                            }
                            i++;
                        }
                    }
                    else
                    {
                        oecnode.AppendChild(createoeline("", xml));
                    }
                    
                    
                    cellnode.PrependChild(oecnode);
                    rownode.PrependChild(cellnode);
                }

                table.PrependChild(rownode);
            }
            table.PrependChild(columns);
        }
        /// <summary>
        /// remove recycle bin sectiongroup
        /// </summary>
        /// <param name="xml"></param>
        public XmlDocument removerecyclebin(XmlDocument xml)
        {
            XmlNodeList sectiongroup = xml.SelectNodes("//one:SectionGroup[@name = \"OneNote_RecycleBin\"]", onenotenamespace);
            foreach (XmlNode y in sectiongroup)
            {
                y.ParentNode.RemoveChild(y);
            }
            return xml;
        }
        /// <summary>
        /// remove font format from string
        /// </summary>
        public string removeformat(string line)
        {
            line = line.Replace("en-US", "\"en-US\"");
            line = line.Replace("zh-CN", "\"zh-CN\"");
            line = "<?xml version=\"1.0\"?><a>" + line + "</a>";
            XmlDocument temp = new XmlDocument();
            temp.LoadXml(line);
            line = temp.InnerText;
            return line;
        }
        /// <summary>
        /// get unfinshitag nodes from page's xml
        /// </summary>
        public XmlNodeList getunfinishtags(XmlNode pagexml)
        {
            string pageid = pagexml.Attributes.GetNamedItem("ID").Value;
            var inpagexml = GetPageContent(pageid);
            XmlNodeList completedtags = inpagexml.SelectNodes("//one:Tag[@completed = \"true\"]", onenotenamespace);
            foreach (XmlNode z in completedtags)
            {
                z.ParentNode.RemoveChild(z);
            }
            XmlNodeList unfinishtag = inpagexml.SelectNodes("//one:Tag", onenotenamespace);
            return unfinishtag;
        }
        /// <summary>
        /// get xmlnode list of time task
        /// </summary>
        /// <returns></returns>
        public ArrayList[] getoenodelist(XmlDocument xml)
        {
            ArrayList day = new ArrayList();
            //ArrayList daytime = new ArrayList();
            ArrayList week = new ArrayList();
            ArrayList weektime = new ArrayList();
            ArrayList month = new ArrayList();
            ArrayList monthtime = new ArrayList();
            ArrayList year = new ArrayList();
            ArrayList yeartime = new ArrayList();
            ArrayList specifc = new ArrayList();
            ArrayList specifctime = new ArrayList();
            string cdata = "";
            ArrayList[] alltask;
            XmlNodeList pagenodes = hierxml.SelectNodes("//one:Page", onenotenamespace);
            foreach (XmlNode m in pagenodes)
            {
                XmlNodeList unfinishtags = getunfinishtags(m);
                foreach (XmlNode n in unfinishtags)
                {
                    string timetask = n.ParentNode.SelectSingleNode("one:T", onenotenamespace).InnerText;
                    XmlDocument timedoc = new XmlDocument();

                    //parse time
                    if (timetask.Contains("$"))
                    {
                        string link;
                        OneNoteApplication.GetHyperlinkToObject(m.Attributes.GetNamedItem("ID").Value, n.ParentNode.Attributes.GetNamedItem("objectID").Value, out link);
                        if (timetask.IndexOf("</span>") != -1)
                        {
                            timetask = removeformat(timetask);
                        }
                        
                        string[] times = timetask.Split(new char[] {'$'});
                        string time = times[1];
                        string thing = times[0];
                        if (thing.Length >= 4)
                        { thing = thing.Substring(0, 4); }
                        
                        string[] timepiece = time.Split(new char[] {' '});
                            switch (timepiece[0])
                            {
                                case "DAY":
                                cdata = "<a href=\"" + link + "\"> <span style = 'background:#92D050' > " + thing + " </span> </a>";
                                day.Add(createoeline(cdata,xml));
                                    break;
                                case "WEEK":
                                cdata = "<a href=\"" + link + "\"> <span style = 'background:#00B0F0' >" + thing + " </span> </a>";
                                week.Add(createoeline(cdata, xml));
                                weektime.Add(Convert.ToInt32( timepiece[1]));
                                    break;
                                case "MONTH":
                                cdata = "<a href=\"" + link + "\"> <span style = 'background:#FFC000' >" + thing  + " </span> </a>";
                                month.Add(createoeline(cdata, xml));
                                monthtime.Add(Convert.ToInt32(timepiece[1]));
                                break;
                                case "YEAR":
                                cdata = "<a href=\"" + link + "\"> <span style = 'background:#FF99CC' >" + thing  + " </span> </a>";
                                year.Add(createoeline(cdata, xml));
                                yeartime.Add(Convert.ToInt32(timepiece[1]));
                                break;
                                default:
                                cdata = "<a href=\"" + link + "\"> " + thing  + "  </a>";
                                specifc.Add(createoeline(cdata, xml));
                                specifctime.Add(DateTime.Parse(timepiece[0]));
                                break;
                            }        
                    }
                }
            }

            alltask = new ArrayList[9]{ day, week,weektime, month,monthtime, year,yeartime, specifc,specifctime};
            return alltask;
        }
        public void creattimetasktable(XmlDocument xml)
        {
            //var xml = GetPageContent(getpageid("OVERVIEW"));
            ArrayList[] oenodelist = getoenodelist(xml);
            XmlNode oec = xml.SelectSingleNode("//one:OEChildren", onenotenamespace);
            creattable(7, 5, oenodelist,oec, xml);
            //this.OneNoteApplication.UpdatePageContent(xml.InnerXml, DateTime.MinValue);
        }
        
        //create a treeed todo section-page  based on a existed blank todo page
        public void createtodopage() 
        {
            //get id of todo page
            var todoid = getpageid("OVERVIEW");
            if (todoid == null)
            {
                IntPtr myWindowHandle = new IntPtr((long)this.OneNoteApplication.Windows.CurrentWindow.WindowHandle);
                NativeWindow nativeWindow = new NativeWindow();
                nativeWindow.AssignHandle(myWindowHandle);
                MessageBox.Show(nativeWindow, "没有找到名为“OVERVIEW”的页面，请选择一个分区新建此页面！");
                return;
            }
            
            //get page xml of page
            var targetwholexml = GetPageContent(todoid);

            XmlNode oec = targetwholexml.SelectSingleNode("//one:OEChildren", onenotenamespace);
            oec.RemoveAll();

            creattimetasktable(targetwholexml);

            addtagdef(targetwholexml);

                      
            //loop through notebook to get all task
            XmlNode sectionnode;
            sectionnode = targetwholexml.SelectSingleNode("//one:OEChildren", onenotenamespace);
            XmlNode notebooknode = hierxml.SelectSingleNode("//one:Notebook[@name = \"个人\"]", onenotenamespace);
            XmlNodeList sectionnodes = notebooknode.SelectNodes(".//one:Section", onenotenamespace);
            foreach (XmlNode m in sectionnodes)
            {
                sectionnode = AddPageline(targetwholexml, m.Attributes.GetNamedItem("name").Value,linetype.oeline);
                XmlNodeList pagenodes = m.SelectNodes(".//one:Page", onenotenamespace);
                foreach (XmlNode x in pagenodes)
                {
                    string pageid = x.Attributes.GetNamedItem("ID").Value;
                    var inpagexml = GetPageContent(pageid);
                    string tagdef0Index = "\"\"";
                    string tagdef1Index = "\"\"";
                    //get index of tag, type 0 -> task, type 1 -> suspended
                    XmlNode tagdef0 = inpagexml.SelectSingleNode("/one:Page/one:TagDef[@type = \"0\"]", onenotenamespace);
                    if(tagdef0 != null)
                    {
                        tagdef0Index = tagdef0.Attributes.GetNamedItem("index").Value;
                    }
                    
                    XmlNode tagdef1 = inpagexml.SelectSingleNode("/one:Page/one:TagDef[@type = \"1\"]", onenotenamespace);
                    if (tagdef1 != null)
                    {
                        tagdef1Index = tagdef1.Attributes.GetNamedItem("index").Value;
                    }

                    
                    //if page have tagdef then contine
                    if (tagdef0 != null || tagdef1!= null)
                    {
                        //get all task tags
                        XmlNodeList tasktags = inpagexml.SelectNodes("//one:Tag[@index = " + tagdef0Index + "]", onenotenamespace);
                        //get number of unfinished task and finished task
                        int CountOfUnfinished = 0;
                        int CountOffinished = 0;
                        foreach (XmlNode n in tasktags)
                        {
                            string CompletedStatus = n.Attributes.GetNamedItem("completed").Value;
                            if (CompletedStatus == "true")
                            {
                                CountOffinished++;
                            }
                            else
                            {
                                CountOfUnfinished++;
                            }
                            
                        }
                        //get all suspended tags
                        XmlNodeList suspendedtags = inpagexml.SelectNodes("//one:Tag[@index = " + tagdef1Index + "]", onenotenamespace);
                        int CountOfSuspended = suspendedtags.Count;
                        /***XmlNodeList completedtags = inpagexml.SelectNodes("//one:Tag[@completed = \"true\"]", onenotenamespace);
                        int? numcompleted = completedtags.Count;
                        foreach (XmlNode z in completedtags)
                        {
                            z.ParentNode.RemoveChild(z);
                        }
                        int? numincompleted = inpagexml.SelectNodes("//one:Tag", onenotenamespace).Count;***/
                        if (CountOffinished != 0 || CountOfUnfinished != 0 || CountOfSuspended!=0)
                        {
                            //get link of page
                            string link;
                            OneNoteApplication.GetHyperlinkToObject(pageid, "", out link);

                            XmlElement line0 = targetwholexml.CreateElement("one", "OEChildren", NS);
                            XmlElement line1 = targetwholexml.CreateElement("one", "OE", NS);
                            XmlElement line2 = targetwholexml.CreateElement("one", "T", NS);
                            XmlCDataSection cdata = targetwholexml.CreateCDataSection(addvirtualline(CountOffinished, CountOfUnfinished, CountOfSuspended, x.Attributes.GetNamedItem("name").Value, link));

                            sectionnode.AppendChild(line0);
                            line0.AppendChild(line1);
                            line1.AppendChild(line2);
                            line2.AppendChild(cdata);
                        }
                    }
                   
                }
            }
            removeblank(ref targetwholexml);
                      
            OneNoteApplication.UpdatePageContent(targetwholexml.InnerXml, DateTime.MinValue);
        }

        //creat vitrual line for cdata
        public string addvirtualline(int? completed, int? incompleted, int? suspended, string pagename,string link)
        {
            string completedstr = completed.ToString();
            string incompletedstr = incompleted.ToString();
            string suspendedstr = suspended.ToString();
            string line;
            //create completedstr
            if (completed != 0)
            {
                while (completed >= 1)
                {
                    completedstr = completedstr + " ";
                    completed--;
                }
                completedstr = "<span style = 'background:lime;mso-highlight:lime' >" + completedstr + "</ span >";//< spanstyle = 'background:lime;mso-highlight:lime' > 12 </ span >
            }
            else
            { completedstr = ""; }
            //create incompletedstr
            if (incompleted != 0)
            {
                while (incompleted >= 1)
                {
                    incompletedstr = incompletedstr + " ";
                    incompleted--;
                }
                incompletedstr = "<span style = 'background:red;mso-highlight:red' >" + incompletedstr + "</ span >";//<span style = 'background:red;mso-highlight:red' > 24 </ span >
            }
            else
            {
                incompletedstr = "";
            }
            //create suspendedstr
            if (suspended != 0)
            {
                while (suspended >= 1)
                {
                    suspendedstr = suspendedstr + " ";
                    suspended--;
                }
                suspendedstr = "<span style = 'background:yellow;mso-highlight:yellow' >" + suspendedstr + "</ span >";//<span style = 'background:red;mso-highlight:red' > 24 </ span >
            }
            else
            {
                suspendedstr = "";
            }
            //create out line
            line = "<a href=\"" + link + "\"> " + pagename + " </span> </a>" + completedstr + suspendedstr + incompletedstr;
            return line;
        }

        //go through page to get a todolist xml
        public void recursion(XmlNode node, XmlNode oec, ref XmlDocument targetpagexml)
        {
            if (node.HasChildNodes == true)
            {
                XmlNodeList childlist = node.ChildNodes;
                //if (!(node.Name == "one:OE" & node.SelectSingleNode("//one:Tag", GetNSManager(targetpagexml.NameTable)) == null))         

                foreach (XmlNode child in childlist)
                {
                    if (child.Name == "one:OEChildren")
                    {
                        XmlElement tabchildren = targetpagexml.CreateElement("one", "OEChildren", NS);// MessageBox.Show(targetpagexml.InnerXml);
                        if (oec.HasChildNodes == true & oec.Name == "one:OE")
                        {
                            oec.AppendChild(tabchildren);
                            oec = tabchildren;
                        }
                    }
                   

                    if (child.Name == "one:Tag" & oec.Name == "one:OEChildren")//[@completed = \"true\"] & (child.Attributes.GetNamedItem("completed").Value == "true") 
                    {
                        
                        if (child.Attributes.GetNamedItem("completed").Value == "true")
                        { }
                        else if (true)
                        {
                            
                            XmlElement tagchildren = targetpagexml.CreateElement("one", "Tag", NS);
                            //tagchildren.SetAttribute("","");completed="false" disabled="false"
                            //tagchildren.SetAttribute("completed", child.Attributes.GetNamedItem("completed").Value);
                            tagchildren.SetAttribute("disabled", "false");
                            tagchildren.SetAttribute("index", "0");
                            XmlElement oechildren = targetpagexml.CreateElement("one", "OE", NS);
                            //oechildren.SetAttribute("style", "font-family:'Microsoft YaHei';font-size:11.0pt");
                            //oechildren.SetAttribute("alignment", "left");
                            XmlElement tchildren = targetpagexml.CreateElement("one", "T", NS);
                            XmlCDataSection cdata = targetpagexml.CreateCDataSection(child.ParentNode.SelectSingleNode("one:T", onenotenamespace).InnerText);

                            oec.AppendChild(oechildren);
                            oechildren.AppendChild(tagchildren);
                            oechildren.AppendChild(tchildren);
                            tchildren.AppendChild(cdata);
                            oec = oechildren;
                        }                        
                    }
                    recursion(child, oec, ref targetpagexml);
                }
            }
        }
        
        //move the finished task to the top
        public void sortcurrentpage()
        {
            XmlDocument pagexml = GetPageContent(OneNoteApplication.Windows.CurrentWindow.CurrentPageId);
            sortcompletedtask(pagexml.DocumentElement, ref pagexml);
            OneNoteApplication.UpdatePageContent(pagexml.InnerXml, DateTime.MinValue);
        }

        //creat new task based on the input from newtimeline form
        public void creatnewtask()
        {

        }

        //set page name
        //pageid pagename -> newpage
        public void setpagename(XmlDocument pagexml,string pagename)
        {
            XmlNode titlenode = pagexml.GetElementsByTagName("one:Title").Item(0);
            XmlNode tnode = titlenode.SelectSingleNode("//one:T", onenotenamespace);
            XmlCDataSection cdata = pagexml.CreateCDataSection(pagename);

            tnode.RemoveAll();
            tnode.AppendChild(cdata);
        }
        //move completed tag to front.
        public void sortcompletedtask(XmlNode oe, ref XmlDocument targetpagexml)
        {
            if (oe.HasChildNodes)
            {
                XmlNodeList childlist = oe.SelectNodes("//one:OEChildren", onenotenamespace);
                //if (!(node.Name == "one:OE" & node.SelectSingleNode("//one:Tag", GetNSManager(targetpagexml.NameTable)) == null))         

                foreach (XmlNode child in childlist)
                {
                    XmlNodeList oenodes = child.ChildNodes;
                    List<XmlNode> completednodes = new List<XmlNode>();
                    foreach (XmlNode oenode in oenodes)
                    {
                        if (oenode.SelectSingleNode("./one:Tag[@completed = \"true\"]",onenotenamespace) != null)
                        {
                            completednodes.Add(oenode);
                        }
                    }
                    foreach (XmlNode completednode in completednodes)
                    {
                        child.RemoveChild(completednode);
                        child.PrependChild(completednode);
                    }

                }
            }
        }

        public void removecompletedtask()//remove completed tag.
        {
            string id = OneNoteApplication.Windows.CurrentWindow.CurrentPageId;
            var pagexml = GetPageContent(id);
            var x = pagexml.SelectNodes("//one:Tag[@completed = \"true\"]", onenotenamespace);
            foreach (XmlNode tagnode in x)
            {
                try
                {
                    XmlNode oenode = tagnode.ParentNode;
                    oenode.ParentNode.RemoveChild(oenode);

                }
                catch (Exception e)
                {
                    throw new ApplicationException("Error in GetPageContent: " + e.Message, e);
                }
            }
            writetotxt(pagexml.InnerXml);
            MessageBox.Show("finishloop");
            removeblank(ref pagexml);
            writetotxt(pagexml.InnerXml);
            OneNoteApplication.UpdatePageContent(pagexml.InnerXml, DateTime.MinValue);
            MessageBox.Show("finishupdate");
        }


        public void test()//add more attributes to element,failed
        {           
            string id = this.OneNoteApplication.Windows.CurrentWindow.CurrentPageId;
            XmlDocument xml = GetPageContent(id);
            XmlNode oecnode = xml.SelectSingleNode("//one:OEChildren",onenotenamespace);

            XmlElement oenode = xml.CreateElement("one", "OE", NS);
            XmlElement tnode = xml.CreateElement("one", "T", NS);
            oenode.SetAttribute("Metadata", "11");
            XmlCDataSection cdata = xml.CreateCDataSection("哈哈哈");

            oecnode.AppendChild(oenode);
            oenode.AppendChild(tnode);
            tnode.AppendChild(cdata);
            writetotxt(xml.InnerXml);
            this.OneNoteApplication.UpdatePageContent(xml.InnerXml, DateTime.MinValue);
        }
        public void writetotxt(string content)//write to f:/test.txt
        {
            FileStream f = new FileStream(@"F:\test.txt", FileMode.Create);
            StreamWriter s = new StreamWriter(f);
            s.Write(content);
            s.Flush();
            s.Close();
            f.Close();
        }
        public void removeblank(ref XmlDocument xml)
        {
            XmlNodeList nodelist = xml.SelectNodes("//one:Outline", onenotenamespace);
            if (nodelist.Count != 0)
            {
                foreach (XmlNode x in nodelist)
                {
                    if (x.SelectSingleNode(".//one:T", onenotenamespace) == null)
                    {
                        x.ParentNode.RemoveChild(x);
                    }
                }
            }
            nodelist = xml.SelectNodes("//one:OE", onenotenamespace);
            if (nodelist.Count != 0)
            {
                foreach (XmlNode x in nodelist)
                {
                    if (x.SelectSingleNode(".//one:T", onenotenamespace) == null)
                    {
                        x.ParentNode.RemoveChild(x);
                    }
                }
            }
            nodelist = xml.SelectNodes("//one:OEChildren", onenotenamespace);
            if (nodelist.Count != 0)
            {
                foreach (XmlNode x in nodelist)
                {
                    if (x.SelectSingleNode(".//one:T", onenotenamespace) == null)
                    {
                        x.ParentNode.RemoveChild(x);
                    }
                }
            }
            nodelist = xml.SelectNodes("//one:Tag", onenotenamespace);
            XmlNode tagdef = xml.SelectSingleNode("//one:TagDef", onenotenamespace);
            if (nodelist.Count == 0 & tagdef !=null)
            {
                tagdef.ParentNode.RemoveChild(tagdef);
            }
            nodelist = xml.SelectNodes("//one:Outline", onenotenamespace);
            XmlNode QuickStyleDef = xml.SelectSingleNode("//one:QuickStyleDef[@name = \"p\"]", onenotenamespace);
            if (nodelist.Count == 0 & QuickStyleDef != null)
            {
                QuickStyleDef.ParentNode.RemoveChild(QuickStyleDef);
            }
        }


    }
}
