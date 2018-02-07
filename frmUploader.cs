using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using MSXML2;
using System.Xml;
using System.IO;
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace ExcelRead
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		# region Control
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.MainMenu mainMenu1;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem2;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ListBox listbox_sheets;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btn_load;
		private System.Windows.Forms.ListBox lbl_status;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Timer timer1;
		private System.Windows.Forms.MenuItem menuItem3;
		private System.Windows.Forms.MenuItem menuItem4;
		private System.Windows.Forms.MenuItem menuItem5;
		# endregion
 

		#region Decalartion
		public ArrayList UploadList = new ArrayList();
		private Microsoft.Office.Interop.Excel.Application  ExcelObj = null;
		int indexID;
		string[] strArray;
		string[] myvalues;
		string strSheetdata = "";
		int countcell = 0;
        string strError;
        Microsoft.Office.Interop.Excel.Workbook theWorkbook = null;
        Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
        Microsoft.Office.Interop.Excel.Range range = null;
		string shtName ="";
		string FileName="";
		string strMsgop;
		string strContentid;
		string strDate = System.DateTime.Now.ToShortDateString();
		string strTime = System.DateTime.Now.ToLongTimeString();
		System.Xml.XmlDocument xoExcelRecords;
		XmlDocument xoDoc = new XmlDocument () ;

        DataLoader objDataLoader = new DataLoader();
		#endregion



		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			// Splash Screen
			/*Thread th = new Thread(new ThreadStart(DoSplash));
			th.Start();
			Thread.Sleep(4000);
			th.Abort();
			Thread.Sleep(1000); */
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.mainMenu1 = new System.Windows.Forms.MainMenu(this.components);
            this.menuItem1 = new System.Windows.Forms.MenuItem();
            this.menuItem2 = new System.Windows.Forms.MenuItem();
            this.menuItem5 = new System.Windows.Forms.MenuItem();
            this.menuItem3 = new System.Windows.Forms.MenuItem();
            this.menuItem4 = new System.Windows.Forms.MenuItem();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.listbox_sheets = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btn_load = new System.Windows.Forms.Button();
            this.lbl_status = new System.Windows.Forms.ListBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainMenu1
            // 
            this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem1,
            this.menuItem3});
            // 
            // menuItem1
            // 
            this.menuItem1.Index = 0;
            this.menuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem2,
            this.menuItem5});
            this.menuItem1.Text = " File ";
            // 
            // menuItem2
            // 
            this.menuItem2.Index = 0;
            this.menuItem2.Text = "Open";
            this.menuItem2.Click += new System.EventHandler(this.menuItem2_Click);
            // 
            // menuItem5
            // 
            this.menuItem5.Index = 1;
            this.menuItem5.Text = "Exit";
            this.menuItem5.Click += new System.EventHandler(this.menuItem5_Click);
            // 
            // menuItem3
            // 
            this.menuItem3.Index = 1;
            this.menuItem3.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem4});
            this.menuItem3.Text = "Help";
            // 
            // menuItem4
            // 
            this.menuItem4.Index = 0;
            this.menuItem4.Text = "About";
            this.menuItem4.Click += new System.EventHandler(this.menuItem4_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "*.xls";
            this.openFileDialog1.Filter = "Excel File (*.xls) | All Files (*.*) ||";
            this.openFileDialog1.Title = "Choose an Excel File";
            // 
            // listbox_sheets
            // 
            this.listbox_sheets.BackColor = System.Drawing.SystemColors.Info;
            this.listbox_sheets.Location = new System.Drawing.Point(8, 40);
            this.listbox_sheets.MultiColumn = true;
            this.listbox_sheets.Name = "listbox_sheets";
            this.listbox_sheets.Size = new System.Drawing.Size(256, 134);
            this.listbox_sheets.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(8, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 16);
            this.label1.TabIndex = 1;
            this.label1.Text = "Sheets";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btn_load);
            this.groupBox1.Controls.Add(this.listbox_sheets);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(8, 8);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(272, 216);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Excel Sheet Load";
            // 
            // btn_load
            // 
            this.btn_load.Location = new System.Drawing.Point(8, 184);
            this.btn_load.Name = "btn_load";
            this.btn_load.Size = new System.Drawing.Size(256, 23);
            this.btn_load.TabIndex = 2;
            this.btn_load.Text = "Load Sheets";
            this.btn_load.Click += new System.EventHandler(this.btn_load_Click);
            // 
            // lbl_status
            // 
            this.lbl_status.BackColor = System.Drawing.SystemColors.Info;
            this.lbl_status.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_status.HorizontalScrollbar = true;
            this.lbl_status.Location = new System.Drawing.Point(8, 16);
            this.lbl_status.Name = "lbl_status";
            this.lbl_status.Size = new System.Drawing.Size(256, 95);
            this.lbl_status.TabIndex = 4;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.lbl_status);
            this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(8, 232);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(272, 120);
            this.groupBox3.TabIndex = 5;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Status";
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(296, 420);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Menu = this.mainMenu1;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Uploader";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			System.Windows.Forms.Application.Run(new Form1());
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			xoExcelRecords = new  XmlDocument();
		}
		
		string[] ConvertToStringArray(System.Array values)
		{
			string[] theArray = new string[values.Length];
			for (int i = 1; i <= values.Length; i++)
			{
				if (values.GetValue(1, i) == null)
					theArray[i-1] = "";
				else
					theArray[i-1] = (string)values.GetValue(1, i).ToString();
			}
			return theArray;
		}

		//*** Start ***//
		// Load the XL sheets to the List Control
		public void menuItem2_Click(object sender, System.EventArgs e)
		{
			this.openFileDialog1.FileName = "*.xls; *.xlsx";
			if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				GetSheetList();
			}
		}

		//*** End ***//
		
		//*** Collects the sheet list in the selected Excel file
		public void GetSheetList()
		{
            ExcelObj = new Microsoft.Office.Interop.Excel.ApplicationClass();
			FileName = openFileDialog1.FileName;
            Microsoft.Office.Interop.Excel.Workbook theWorkbook = ExcelObj.Workbooks.Open(FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Microsoft.Office.Interop.Excel.Sheets sheets = theWorkbook.Worksheets;
			listbox_sheets.Items.Clear();
			for(int j=1;j<=sheets.Count;j++)
			{
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(j);
				listbox_sheets.Items.Add(worksheet.Name);					
			}
			ExcelObj.Application.Quit();
		} 
		//***

		
		//*** Excel Sheet Data to XML
		XmlDocument LoadExcelSheetToXml(string SheetName)
		{
			XmlDocument xoDoc = new XmlDocument () ;

            string[] FileDataColumns = { "ClientId", "DocumentType", "PlannerCode", "ContentType", "FileName", "RegionCode", "Path", "DateModified", "PlannerType", "Category" };

			string[] Columns = FileDataColumns ;

            if (SheetName.ToString() == "DocumentList")
			{
				Columns = FileDataColumns ;
			}
			else 
			{
                MessageBox.Show("Column list miss match");
			}

            ExcelObj = new Microsoft.Office.Interop.Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Workbook theWorkbook = ExcelObj.Workbooks.Open(FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, null, null);

            Microsoft.Office.Interop.Excel.Sheets sheets = theWorkbook.Worksheets;
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(SheetName);

			string			strName = "<" + SheetName.ToUpper ().Replace (" ", "_")+ "S/>";
			xoDoc.LoadXml ( strName) ;

			XmlElement xoRoot = xoDoc.DocumentElement ;

			int Row = 2 ;
			while ( true )
			{
                Microsoft.Office.Interop.Excel.Range rg1 = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[Row, 1];
				if ( rg1.Text != null )
				{
					if (rg1.Text.ToString()=="")
						break;
				}
				else
					break;
				
				XmlElement xoSelectedSheet = xoDoc.CreateElement (SheetName.ToUpper ().Replace (" ", "_") );

				// Reads the coloumns
				for ( int j = 0 ; j < Columns.Length ; j++ )
				{
                    Microsoft.Office.Interop.Excel.Range rg = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[Row, j + 1];
					string Val = "" ;

					if ( rg.Text != null )
						Val = rg.Text.ToString () ;
					xoSelectedSheet.SetAttribute (Columns[j].ToLower ().Replace (" ", "_"), Val );
				}
				Row++;
				xoRoot.AppendChild (xoSelectedSheet);
			}
			theWorkbook.Close (false, null, null);
			ExcelObj.Application.Quit(); 
			return xoDoc ;
		}

		//*** Loads The Sheet Data 
		private void btn_load_Click(object sender, System.EventArgs e)
		{
			if (listbox_sheets.Items.Count>0)
			{
				if (listbox_sheets.SelectedIndex !=-1)
				{
					shtName = listbox_sheets.SelectedItem.ToString();
					try
					{
                        if (shtName == "DocumentList")
						{
                            XmlDocument xoDoc = LoadExcelSheetToXml("DocumentList");
							MigratePost(xoDoc);
						}
						else
						{
							lbl_status.Items.Add("Select a valid sheet");
						}
					}
					catch(Exception ex)
					{
						lbl_status.Items.Add("Unformatted sheet" + ex.Message);
					}
				}
			}
		}

		// Post the XML to The .ASPX file and get the response as XML
		public void MigratePost(XmlDocument xoDetails)
		{
			try 
			{
                objDataLoader.fileMigrate(xoDetails);
				
				//SaveXLLog(xoDetails.OuterXml);
				FormatXLlogfile(xoDetails.OuterXml);
			}
			catch ( Exception ex)
			{
				string msg = ex.Message;
				lbl_status.Items.Add("Error While Loading The Data" + strDate + strTime + ex.Message );
			}
		}

		
		// Save The OutPut File as Excel File 
		public void SaveXLLog(string xml)
		{
			string strFileName = "OutPut_LDM";
			string path = @"C:\";
			
			FileStream file=new FileStream(path+"\\"+strFileName+".xls" ,FileMode.Create);
			StreamWriter wr=new StreamWriter(file);
			wr.Write(xml);   
			wr.Close();
			file.Close(); 
			Message();
            lbl_status.Items.Add(@" The Data in the Excel Sheet is added and the details are logged in c:\OutPut_LDM.xls  " + strDate + strTime);
			
			FileStream file1=new FileStream(path+"\\"+strFileName+".txt" ,FileMode.Create);
			StreamWriter wr1=new StreamWriter(file1);
			wr1.Write(xml);   
			wr1.Close();
			file1.Close();  
		}

		//
		public void Message()
		{
			MessageBox.Show("Upload Successfull");
		}

		// Creates the XL for with the oputput data 
		public void FormatXLlogfile(string xmlOPdata)
		{
            ExcelObj = new Microsoft.Office.Interop.Excel.ApplicationClass();
			ExcelObj.Visible = false;
            theWorkbook = (Microsoft.Office.Interop.Excel.Workbook)(ExcelObj.Workbooks.Add(Type.Missing));
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)theWorkbook.ActiveSheet;
				
			XmlDocument xoDocXL = new XmlDocument();
			xoDocXL.LoadXml(xmlOPdata);

            XmlElement xoOPdata = (XmlElement)xoDocXL.SelectSingleNode("DocumentList");

			// Document List Log XL File Creation
			if ( xoOPdata != null)
			{
				if (xoOPdata.HasChildNodes )
				{
					try 
					{
                        worksheet.Cells[1,1] = "ClientId";
						worksheet.Cells[1,2] = "DocumentType";
						worksheet.Cells[1,3] = "PlannerCode";
						worksheet.Cells[1,4] = "ContentType";
						worksheet.Cells[1,5] = "FileName";
						worksheet.Cells[1,6] = "RegionCode";
						worksheet.Cells[1,7] = "Path";
						worksheet.Cells[1,8] = "DateModified";
						worksheet.Cells[1,9] = "PlannerType";
                        worksheet.Cells[1,10] = "Category";
						worksheet.Cells[1,11] = "Status";
						worksheet.Cells[1,12] = "Error";

                        Microsoft.Office.Interop.Excel.Range sRg = (Microsoft.Office.Interop.Excel.Range)worksheet.get_Range("A1", "L1");
						sRg.Font.Bold = true;

						try
						{
							int Row = 2 ;
							foreach(XmlNode xoOutData in xoOPdata)
							{
                                xoOPdata.GetAttribute("ClientId");
								XmlElement xoTemp =  (XmlElement)xoOutData;

                                strContentid = xoTemp.GetAttribute("ClientId");

                                worksheet.Cells[Row, 1] = xoTemp.GetAttribute("ClientId");

                                worksheet.Cells[Row, 2] = xoTemp.GetAttribute("DocumentType");
                                worksheet.Cells[Row, 3] = xoTemp.GetAttribute("PlannerCode");
                                worksheet.Cells[Row, 4] = xoTemp.GetAttribute("ContentType");
                                worksheet.Cells[Row, 5] = xoTemp.GetAttribute("FileName");
                                worksheet.Cells[Row, 6] = xoTemp.GetAttribute("RegionCode");
                                worksheet.Cells[Row, 7] = xoTemp.GetAttribute("Path");
                                worksheet.Cells[Row, 8] = xoTemp.GetAttribute("DateModified");
                                worksheet.Cells[Row, 9] = xoTemp.GetAttribute("PlannerType");
                                worksheet.Cells[Row, 9] = xoTemp.GetAttribute("Category");
                                //worksheet.Cells[Row, 9] = xoTemp.GetAttribute("PlannerType");

								if (strError == "Error")
								{
									worksheet.Cells[Row,11] = "Error found in this record";
                                    Microsoft.Office.Interop.Excel.Range sRgx2 = (Microsoft.Office.Interop.Excel.Range)worksheet.get_Range("K" + Row, "K" + Row);
									sRgx2.Font.Color = Color.Red.ToKnownColor();
									sRgx2.Font.Bold = true;
								}
								else
								{
									worksheet.Cells[Row,11] = "";
									if (strContentid == "NEW")
									{
										worksheet.Cells[Row,10] = "New Inserted";
                                        Microsoft.Office.Interop.Excel.Range sRgx = (Microsoft.Office.Interop.Excel.Range)worksheet.get_Range("L" + Row, "L" + Row);
										sRgx.Font.Color = Color.Green.ToKnownColor();
									}
									else
									{
										/*worksheet.Cells[Row,10] = "The Record Already Exist";
                                        Microsoft.Office.Interop.Excel.Range sRgx1 = (Microsoft.Office.Interop.Excel.Range)worksheet.get_Range("J" + Row, "J" + Row);
										sRgx1.Font.Color = Color.RoyalBlue.ToKnownColor();*/
									}
								}
								Row++;	
							}
						}
						catch(Exception ex)
						{
							strMsgop = ex.Message;
							lbl_status.Items.Add("Error while Forming Theater Output Data  " + strDate + strTime  + ex.Message);
						}
						string FileNameOP = "MigratedOutPut"+strDate+strTime;
						FileNameOP=FileNameOP.Replace("/","_");
						FileNameOP=FileNameOP.Replace(":","_");
						FileNameOP=FileNameOP.Replace(" ","");
						FileNameOP=FileNameOP.Trim();

                        theWorkbook.SaveAs(FileNameOP, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, null, null, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, false, false, null, null, null);
						theWorkbook.Application.Quit();
					}
					catch(Exception op)
					{
						strMsgop = op.Message;
						lbl_status.Items.Add("Error while Creating report Output File  " + strDate + strTime + op.Message);
						theWorkbook.Application.Quit();
					}
					lbl_status.Items.Add("Successfully Created Your Migration Output File in your MyDocuments Folder " + strDate + strTime );
				}
			}
		}

		// Splash Screen
		/*private void DoSplash()
		{
			Splash sp = new Splash();
			sp.ShowDialog();
			//sp.Close();
		}*/

		private void timer1_Tick(object sender, System.EventArgs e)
		{
		}

		private void menuItem5_Click(object sender, System.EventArgs e)
		{
			System.Windows.Forms.Application.Exit();
		}

		private void menuItem4_Click(object sender, System.EventArgs e)
		{
			
		}
		
		

	}
}
