using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Xml;
using System.Data.SqlClient;

using Sulekha.Content;

namespace Sulekha.MOVIES.Showtimes
{
	/// <summary>
	/// Summary description for setShowTimes.
	/// </summary>
	public class setShowTimes : System.Web.UI.Page
	{
		string sql="";
		string tempSTr;
		string strDate = System.DateTime.Now.ToShortDateString();
		string strAirportcode;
		SqlConnection  dbCon ;
		SqlDataAdapter daAdapter ;
			
		DataSet dsTheater = new DataSet();			
		DataTable dbTable = new DataTable();
		
		DataRow RecInsert ;

		XmlDocument xoXLdata = new XmlDocument();
		System.Xml.XmlDocument xDoc = new System.Xml.XmlDocument () ;

		private void Page_Load(object sender, System.EventArgs e)
		{
			// Put user code to initialize the page here

			
			XmlTextReader xmlRead = new XmlTextReader(Request.InputStream);
			XmlDocument xoXLdata = new XmlDocument();
			xoXLdata.Load(xmlRead);
			
			Sulekha.Content.CityTheaters ct = new CityTheaters();
			// 
			XmlElement xoUploadData = (XmlElement)xoXLdata.SelectSingleNode("THEATERS");
			XmlElement xoMvUploadData = (XmlElement)xoXLdata.SelectSingleNode("MOVIES");
			XmlElement xoShTmUploadData = (XmlElement)xoXLdata.SelectSingleNode("SHOWTIMES");
			
			// Theater Operations
			if (xoUploadData !=null)
			{
				if ( xoUploadData.HasChildNodes)
				{	
					XmlDocument xoDBdata = new XmlDocument();
					xoDBdata.LoadXml(ct.GetTheaterST().OuterXml);
					string strThrXL;
					XmlElement xoDBLoad = (XmlElement)xoDBdata.SelectSingleNode("THEATERS"); 
				
					// Set the contentid value as NEW for identity
					foreach(XmlNode xoSibData in xoUploadData)
					{
						XmlElement tempxmlEle=(XmlElement) xoSibData;
						strThrXL = tempxmlEle.GetAttribute("name");
						strThrXL = strThrXL.Trim();
						strAirportcode = tempxmlEle.GetAttribute("airportcode");
						strAirportcode = strAirportcode.Trim();

						XmlElement xotemp = (XmlElement)xoDBLoad.SelectSingleNode("ATTRMAP[@name='"+strThrXL+"' and @AirportCode ='"+ strAirportcode +"']");
						//XmlElement xotemp = (XmlElement)xoDBLoad.SelectSingleNode("ATTRMAP[@name='"+strThrXL+"']");
						if (xotemp == null)
						{
							tempxmlEle.SetAttribute("contentid", "NEW");
						}
						else
						{
							tempxmlEle.SetAttribute("contentid", xotemp.GetAttribute("contentid") );
						}
					}
			
					/* Compare the XL data with the DB Data and Insert the data 
					   from the XL to DB which is not exist in the DB	by using the identity 'NEW' 	*/
					string strContent = "NEW";
					string strCategory = "Movie Theaters";
					foreach(XmlNode xoSibData in xoUploadData)
					{
						//string strContent;
						//strContent = xoUploadData.GetAttribute("contentid");
						XmlElement xoTempInsert = (XmlElement) xoSibData;

						string strContentID = xoTempInsert.GetAttribute("contentid");
						//XmlElement xoInsert = (XmlElement) xoTempInsert.SelectSingleNode("THEATER[@contentid='" + strContent+ "']");
						string  xoInsert = xoTempInsert.GetAttribute("@contentid='" + strContent+ "'");
				
						//if ( xoInsert != "" || xoInsert != null)
						if ( strContentID == "" || strContentID == "NEW" )
						{
							XmlElement xoElem = (XmlElement) xoSibData;
							/*sql = "Select * from sulekha..YP where contentid=-1";*/
					
							try
							{
								sql = "INSERT INTO Sulekha..YP (Name, Address1, City, State, eMail, Phone, airportcode, websiteurl, category) VALUES ( ";
								sql += "'" + xoElem.GetAttribute("name") + "', " 
									+ "'" + xoElem.GetAttribute("address1") + "', " 
									+ "'" + xoElem.GetAttribute("city") + "', " 
									+ "'" + xoElem.GetAttribute("state") + "', " 
									+ "'" + xoElem.GetAttribute("email") + "', " 
									+ "'" + xoElem.GetAttribute("phone") + "', " 
									+ "'" + xoElem.GetAttribute("airportcode") + "', " 
									+ "'" + xoElem.GetAttribute("websiteurl") + "', " 
									+ "'" + strCategory + "');" ; 
					
								Insertdata(sql);
								xoTempInsert.SetAttribute("error", "");
							}
							catch(Exception ex)
							{
								xoTempInsert.SetAttribute("error", "Error");
							}
						}
						else
						{
							tempSTr=xoXLdata.OuterXml;
						}

						
					}
				}
				tempSTr=xoXLdata.OuterXml;
			}


			// Movies Operations
			if (xoMvUploadData!=null) 
			{
				if (xoMvUploadData.HasChildNodes ) 
				{
					XmlDocument xoDBdata = new XmlDocument();
					xoDBdata.LoadXml(ct.GetMovieST().OuterXml);
					string strMovXL;
					XmlElement xoDBLoad = (XmlElement)xoDBdata.SelectSingleNode("MOVIES"); 
				
					// Set the contentid value as NEW for identity
					foreach(XmlNode xoSibData in xoMvUploadData)
					{
						XmlElement tempxmlEle=(XmlElement) xoSibData;
						strMovXL =tempxmlEle.GetAttribute("title");
						strMovXL = strMovXL.Trim();
						XmlElement xotemp = (XmlElement)xoDBLoad.SelectSingleNode("ATTRMAP[@Title='"+strMovXL+"']");
						if (xotemp == null)
						{
							tempxmlEle.SetAttribute("contentid", "NEW");
						}
						else
						{
							tempxmlEle.SetAttribute("contentid", xotemp.GetAttribute("contentid") );
						}
					}
			
					/* Compare the XL data with the DB Data and Insert the data 
						   from the XL to DB which is not exist in the DB			*/
					string strContent = "NEW";

					foreach(XmlNode xoSibData in xoMvUploadData)
					{
						XmlElement xoTempInsert = (XmlElement) xoSibData;
						string  xoInsert = xoTempInsert.GetAttribute("@contentid='" + strContent+ "'");
				
						string strContentID = xoTempInsert.GetAttribute("contentid");

						if ( strContentID == "" || strContentID == "NEW" )
							//if ( xoInsert != "" || xoInsert != null)
						{
							XmlElement xoElem = (XmlElement) xoSibData;

							string strArtist = xoElem.GetAttribute("artists");
							string strContentTypeID = "12000";
							string strFilter = "ShowTimes";
							string strMovieName = xoElem.GetAttribute("title");

							strArtist = strArtist.Replace("'","''");
							strMovieName = strMovieName.Replace("'","''");
							
							try
							{
								sql = "INSERT INTO Sulekha..Movies (Title, Lang, Category, Artists, URL, ContentTypeID, FilterBy, CrDate) VALUES ( ";
								sql += "'" + strMovieName + "', " 
									+ "'" + xoElem.GetAttribute("language") + "', " 
									+ "'" + xoElem.GetAttribute("category") + "', " 
									+ "'" + strArtist + "', "
									+ "'" + xoElem.GetAttribute("url") + "', " 
									+ "'" + strContentTypeID + "', " 
									+ "'" + strFilter + "', " 
									+ "'" + strDate + "');" ; 
					
								Insertdata(sql);
								xoTempInsert.SetAttribute("error", "");
							}
							catch(Exception ex)
							{
								xoElem.SetAttribute("error", "Error");
								xoTempInsert.SetAttribute("error", "Error");
							}
						}
						else
						{
							tempSTr=xoXLdata.OuterXml;
						}
					}
				}
				tempSTr=xoXLdata.OuterXml;
			}


			// ShowTime Operations
			if (xoShTmUploadData!=null) 
			{
				if (xoShTmUploadData.HasChildNodes ) 
				{
					// Get the Movies List from the Movies table
					XmlDocument xoMovdataDB = new XmlDocument();
					xoMovdataDB.LoadXml(ct.GetMovieST().OuterXml);
					
					string strMovXLtitleDB;
					string strMovXLcontent;
					string strMovXL;
					
					XmlElement tempEle = (XmlElement) xoMovdataDB.SelectSingleNode("MOVIES");
					
					// Set the contentid value as NEW for identity
					// Get the content id from the movie table and adding it to the XL Data
					try
					{
						foreach(XmlNode xoSibData in xoShTmUploadData)
						{
							XmlElement tempxmlEle=(XmlElement) xoSibData;
							strMovXL = tempxmlEle.GetAttribute("movietitle");
							strMovXL = strMovXL.Trim();
							strAirportcode = tempxmlEle.GetAttribute( "airportcode"); 
							strAirportcode = strAirportcode.Trim();

							XmlElement xotemp = (XmlElement)xoMovdataDB.SelectSingleNode("//ATTRMAP[@Title='"+strMovXL+"']");

							if (xotemp == null)
							{
								tempxmlEle.SetAttribute("movieid", "NEW");
							}
							else
							{
								tempxmlEle.SetAttribute("movieid", xotemp.GetAttribute("contentid") );
							}
						}
					}
					catch(Exception ex)
					{
						
					}

					tempSTr=xoXLdata.OuterXml;

					// Get the Theaters List from the YP table
					
					XmlDocument xoThrdata = new XmlDocument();
					xoThrdata.LoadXml(ct.GetTheaterST().OuterXml);
					string strThrXL;
					string strContent = "NEW";
					XmlElement xoThrdata1 = (XmlElement)xoThrdata.SelectSingleNode("THEATERS"); 

					// Set the contentid value as NEW for identity
					//
					try
					{
						foreach(XmlNode xoSibData in xoShTmUploadData)
						{
							XmlElement tempxmlEle=(XmlElement) xoSibData;
							strThrXL = tempxmlEle.GetAttribute("theatername");
							strAirportcode = tempxmlEle.GetAttribute( "airportcode"); 
							strThrXL = strThrXL.Trim();
							strAirportcode = strAirportcode.Trim();

							//XmlElement xotemp = (XmlElement)xoThrdata.SelectSingleNode("//ATTRMAP[@name='"+strThrXL+"'] | //ATTRMAP[@airportcode ='"+ strAirportcode +"']");
							XmlElement xotemp = (XmlElement)xoThrdata.SelectSingleNode("//ATTRMAP[@name='"+strThrXL+"' and @AirportCode ='"+ strAirportcode +"']");
							if (xotemp == null)
							{
								tempxmlEle.SetAttribute("theaterid", "NEW");
							}
							else
							{
								tempxmlEle.SetAttribute("theaterid", xotemp.GetAttribute("contentid") );
							}
						}
					}
					catch (Exception ex)
					{

					}
					tempSTr=xoXLdata.OuterXml;
			
					/* Compare the XL data with the DB Data (Movies and Theaters ) and Insert the data 
						   from the XL to DB 	*/
			
					foreach(XmlNode xoSibData in xoShTmUploadData)
					{
						XmlElement xoTempInsert = (XmlElement) xoSibData;

							XmlElement xoElem = (XmlElement) xoSibData;
							string strStartDate = xoElem.GetAttribute("startdate");  
							string strShowTime = xoElem.GetAttribute("showtime");
							string strShowDateTime = strStartDate + " " + strShowTime ;
							string strThrID = xoElem.GetAttribute("theaterid");	
							string strMovID = xoElem.GetAttribute("movieid");

						try
						{
							if ( strThrID == "" || strThrID == "NEW" || strMovID == "" || strMovID == "NEW" )
							{
								tempSTr=xoXLdata.OuterXml;
							}
							else
							{
								sql = "INSERT INTO Sulekha..Showtimes (movieid, theaterid, showtime, daysofweek, startdate, enddate) VALUES ( ";
								sql += "'" + xoElem.GetAttribute("movieid") + "', " 
									+ "'" + xoElem.GetAttribute("theaterid") + "', " 
									+ "'" + strShowDateTime + "', " 
									+ "'" + xoElem.GetAttribute("daysofweek") + "', " 
									+ "'" + xoElem.GetAttribute("startdate") + "', " 
									+ "'" + xoElem.GetAttribute("enddate")  + "');" ; 
					
								Insertdata(sql);
								xoTempInsert.SetAttribute("error", ""); 
								//xoTempInsert.SetAttribute("error", "You Have No Rights to Upload Movies");
							}
						}
						catch (Exception ex)
						{
							xoTempInsert.SetAttribute("error", "Error");
							//xoTempInsert.SetAttribute("error", "You Have No Rights to Upload Movies");
						}
					}
				}
				tempSTr=xoXLdata.OuterXml;
			}


			//Response.WriteFile(tempSTr);
			Response.Write(tempSTr);
			
		}


		string ConnectStr()
		{
			// Connection String 
			//string strCon = "Initial Catalog=sulekha_profiles;Data Source=sql.sulekha.com;User ID=sa;password=4732gasql0318;";
			string strCon = "Initial Catalog=sulekha_profiles;Data Source=localhost;User ID=sa;password=;";
			return strCon;

		}
		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: This call is required by the ASP.NET Web Form Designer.
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{    
			this.Load += new System.EventHandler(this.Page_Load);
		}
		#endregion


		protected void Insertdata(string sql)
		{
			SqlConnection sqlConn = new SqlConnection(ConnectStr());
			IDbCommand dbCommand	= new SqlCommand();
			dbCommand.CommandText	= sql;
			dbCommand.Connection	= sqlConn;

			int rowsAffected = 0;
			sqlConn.Open();

//			try
//			{
				rowsAffected = dbCommand.ExecuteNonQuery();
//			}
//			catch (Exception ex)
//			{
//				Response.Write("error");
//			}
//			finally
//			{
				dbCommand.Dispose();
				sqlConn.Close();
//			}
		}

		public void TheaterOperation()
		{

		}
		
		string Check(string str, int opt)
		{
			string output="";

			switch(opt)
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
						output = "'" + str.Replace("'","''") + "'";
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
			string output="";

					if (str.Trim().Length > 0)
					{
						output = "'" + str.Replace("'","''") + "'";
					}
					else
					{
						output = "null";
					}				
			return output;

		}

	}

}
		
	