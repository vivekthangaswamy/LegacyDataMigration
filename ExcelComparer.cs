using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Xml;
using System.Data.SqlClient;
using System.Linq;
using System.Linq.Expressions;
using System.Xml.Linq;



namespace ExcelRead
{
    /// <summary>
    /// Summary description for setShowTimes.
    /// </summary>
    public class DataLoader 
    {

        DataSet dsTheater = new DataSet();
        DataTable dbTable = new DataTable();
        
        string strDate = System.DateTime.Now.ToShortDateString();
        XmlDocument xoXLdata = new XmlDocument();
        System.Xml.XmlDocument xDoc = new System.Xml.XmlDocument();

        

        public void fileMigrate(XmlDocument xDocList)
        {

        ////    Sulekha.Content.CityTheaters ct = new CityTheaters();
        //    // 
            XmlElement xoUploadData = (XmlElement)xDocList.SelectSingleNode("DocumentList");

        //    // Theater Operations
        //    if (xoUploadData != null)
        //    {
        //        if (xoUploadData.HasChildNodes)
        //        {
        //            XmlDocument xoDBdata = new XmlDocument();
        //            xoDBdata.LoadXml(ct.GetTheaterST().OuterXml);
        //            string strThrXL;
        //            XmlElement xoDBLoad = (XmlElement)xoDBdata.SelectSingleNode("THEATERS");

        //            // Set the contentid value as NEW for identity
        //            foreach (XmlNode xoSibData in xoUploadData)
        //            {
        //                XmlElement tempxmlEle = (XmlElement)xoSibData;
        //                strThrXL = tempxmlEle.GetAttribute("name");
        //                strThrXL = strThrXL.Trim();
        //                strAirportcode = tempxmlEle.GetAttribute("airportcode");
        //                strAirportcode = strAirportcode.Trim();

        //                XmlElement xotemp = (XmlElement)xoDBLoad.SelectSingleNode("ATTRMAP[@name='" + strThrXL + "' and @AirportCode ='" + strAirportcode + "']");
                       
        //                if (xotemp == null)
        //                {
        //                    tempxmlEle.SetAttribute("contentid", "NEW");
        //                }
        //                else
        //                {
        //                    tempxmlEle.SetAttribute("contentid", xotemp.GetAttribute("contentid"));
        //                }
        //            }

        //            /* Compare the XL data with the DB Data and Insert the data 
        //               from the XL to DB which is not exist in the DB	by using the identity 'NEW' 	*/
        //            string strContent = "NEW";
        //            string strCategory = "Movie Theaters";
        //            foreach (XmlNode xoSibData in xoUploadData)
        //            {
        //                //string strContent;
        //                //strContent = xoUploadData.GetAttribute("contentid");
        //                XmlElement xoTempInsert = (XmlElement)xoSibData;

        //                string strContentID = xoTempInsert.GetAttribute("contentid");
        //                //XmlElement xoInsert = (XmlElement) xoTempInsert.SelectSingleNode("THEATER[@contentid='" + strContent+ "']");
        //                string xoInsert = xoTempInsert.GetAttribute("@contentid='" + strContent + "'");

        //                //if ( xoInsert != "" || xoInsert != null)
        //                if (strContentID == "" || strContentID == "NEW")
        //                {
        //                    XmlElement xoElem = (XmlElement)xoSibData;
        //                    /*sql = "Select * from sulekha..YP where contentid=-1";*/

        //                    try
        //                    {
        //                        sql = "INSERT INTO Sulekha..YP (Name, Address1, City, State, eMail, Phone, airportcode, websiteurl, category) VALUES ( ";
        //                        sql += "'" + xoElem.GetAttribute("name") + "', "
        //                            + "'" + xoElem.GetAttribute("address1") + "', "
        //                            + "'" + xoElem.GetAttribute("city") + "', "
        //                            + "'" + xoElem.GetAttribute("state") + "', "
        //                            + "'" + xoElem.GetAttribute("email") + "', "
        //                            + "'" + xoElem.GetAttribute("phone") + "', "
        //                            + "'" + xoElem.GetAttribute("airportcode") + "', "
        //                            + "'" + xoElem.GetAttribute("websiteurl") + "', "
        //                            + "'" + strCategory + "');";

        //                        Insertdata(sql);
        //                        xoTempInsert.SetAttribute("error", "");
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        xoTempInsert.SetAttribute("error", "Error");
        //                    }
        //                }
        //                else
        //                {
        //                    tempSTr = xoXLdata.OuterXml;
        //                }
        //            }
        //        }
        //        tempSTr = xoXLdata.OuterXml;
        //    }
        //    //Response.WriteFile(tempSTr);
        //   // Response.Write(tempSTr);
        }



        // Distinct of client ID
        public void DataOperation(XmlDocument xDocList)
        {
            XDocument doc = XDocument.Load(xDocList.OuterXml);

            var result = doc.Element("DOCUMENTLISTS")
                    .Elements("DOCUMENTLIST")
                    .Select(e => (string)e.Attribute("ClientId"))
                    .Distinct()
                    .ToList();

        }

        string Check(string str, int opt)
        {
            string output = "";

            switch (opt)
            {
                case 0:
                    if (str.Trim().Length > 0)
                    {
                        output = str;
                    }
                    else
                    {
                        output = "null";
                    }
                    break;

                case 1:
                    if (str.Trim().Length > 0)
                    {
                        output = "'" + str.Replace("'", "''") + "'";
                    }
                    else
                    {
                        output = "null";
                    }
                    break;
            }
            return output;

        }


        string Checkin(string str)
        {
            string output = "";

            if (str.Trim().Length > 0)
            {
                output = "'" + str.Replace("'", "''") + "'";
            }
            else
            {
                output = "null";
            }
            return output;

        }

    }

}

