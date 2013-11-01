using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using System.Collections;

namespace gr.pnevmatikos.office
{
   /// <summary>
   /// This class uses NPOI ver. 1.2.3 -> http://NPOI.codeplex.com/
   /// </summary>
   public class ExcelNPOI : IExcelManager
   {

      private string m_Filename;
      private FileStream m_ExcelFileStreamREAD = null;
      private MemoryStream m_ExcelFileStreamWRITE = null;
      HSSFWorkbook m_CurrWorkbook = null;

      public string Filename
      {
         get { return this.m_Filename; }
         private set { this.m_Filename = value; }
      }

      private FileStream ExcelReadStream
      {
         get { return this.m_ExcelFileStreamREAD; }
         set { this.m_ExcelFileStreamREAD = value; }
      }

      private MemoryStream ExcelWriteStream
      {
         get { return this.m_ExcelFileStreamWRITE; }
         set { this.m_ExcelFileStreamWRITE = value; }
      }

      private FileStream GetInputStream()
      {
         if (this.m_ExcelFileStreamREAD != null)
         {
            return this.m_ExcelFileStreamREAD;
         }
         else
         {
            try
            {
               this.m_ExcelFileStreamREAD = new FileStream(Filename, FileMode.Open, FileAccess.ReadWrite);
               return m_ExcelFileStreamREAD;
            }
            catch (Exception)
            {
               throw new Exception("Could not open " + Filename);
            }
         }//else
      }

      private MemoryStream GetOutputStream()
      {
         if (m_ExcelFileStreamWRITE != null)
         {
            return m_ExcelFileStreamWRITE;
         }
         else
         {
            m_ExcelFileStreamWRITE = new MemoryStream();
            m_CurrWorkbook.Write(m_ExcelFileStreamWRITE);
            return m_ExcelFileStreamWRITE;
         }//else
      }

      public ExcelNPOI()
      {
      }


      public ExcelNPOI(string a_Filename, bool ForceCreate)
      {
         Filename = a_Filename;
         FileInfo a_FileInfo = new FileInfo(Filename);

         //if ForceCreate == true then check if the file exists and delete it.
         if (ForceCreate == true)
         {
            if (a_FileInfo.Exists == true)
            {
               a_FileInfo.Delete();
            }
            m_CurrWorkbook = new HSSFWorkbook();
         }
         else //check if the file exists
         {
            if (!a_FileInfo.Exists)
            {
               throw new FileNotFoundException("File not found");
            }
         }
      }

      public bool Create()
      {
         return true;
      }

      public bool Create(string a_Sheet)
      {
         return true;
      }

      public bool Open()
      {
         if (m_CurrWorkbook == null)
         {
            m_CurrWorkbook = new HSSFWorkbook(GetInputStream(), true);
         }
         return true;
      }

      /// <summary>
      /// Store's all changes made to the workbook.
      /// </summary>
      /// <returns></returns>
      public bool Save()
      {
         try
         {
            Write(this.Filename);
         }
         catch (Exception e)
         {
            throw e;
         }
         return true;
      }

      /// <summary>
      /// Reads an entire Excel file and populates a Dataset with the data found.
      /// A dataset represents the Excel file, the datatables represent the sheets
      /// DataColumns are created and named after the first excel row.
      /// IMPORTANT: The data columns within the Excel should be continuous!
      /// </summary>
      /// <returns>a DataSet holding all values found in the Excel workbook.</returns>
      public DataSet ReadAll()
      {
         HSSFSheet a_sheet = null;
         HSSFRow a_row = null;
         IEnumerator a_RowEnum = null;

         Open();
         //Dataset containing all Excel sheets
         DataSet a_dataset = new DataSet();

         //Get all worksheets...
         for (int a_SheetIndex = 0; a_SheetIndex < m_CurrWorkbook.NumberOfSheets; a_SheetIndex++)
         {
            a_sheet = (HSSFSheet)m_CurrWorkbook.GetSheetAt(a_SheetIndex);
            //Create a new DataTable for the current sheet.
            a_dataset.Tables.Add(new DataTable(a_sheet.SheetName));
            //get the rows
            a_RowEnum = a_sheet.GetRowEnumerator();
            //for each row...
            while (a_RowEnum.MoveNext())
            {
               //get a row
               a_row = (HSSFRow)a_RowEnum.Current;

               for (int a_ColumnIndex = 0; a_ColumnIndex < a_row.LastCellNum; a_ColumnIndex++)
               {
                  // ---- PARSE FIRST ROW ----
                  //if the rownum = 1 then instantiate a datatable for each Column header.
                  //Stop when a null cell is found.
                  if (a_row.RowNum == 0)
                  {
                     if (a_row.GetCell(a_ColumnIndex) != null)
                     {
                        DataColumn a_dataCol = new DataColumn(a_row.GetCell(a_ColumnIndex).ToString());
                        a_dataset.Tables[a_SheetIndex].Columns.Add(a_dataCol);
                     }

                  }//if
                  // ---- PARSE FIRST ROW ----
                  else // ---- PARSE REST ROWS.... ----
                  {
                     if (a_row.GetCell(a_ColumnIndex) != null)
                     {
                        if (a_ColumnIndex == 0)
                        {
                           a_dataset.Tables[a_SheetIndex].Rows.Add(a_dataset.Tables[a_SheetIndex].NewRow());
                        }
                        a_dataset.Tables[a_SheetIndex].Rows[a_dataset.Tables[a_SheetIndex].Rows.Count - 1][a_ColumnIndex] = a_row.GetCell(a_ColumnIndex).ToString();
                     }//if
                  }//else
               }//for
            }//while
         }//for

         //Finally...
         return a_dataset;
      }

      public void Write(string a_Filename)
      {
         if (a_Filename.Trim().Length == 0)
         {
            return;
         }
         try
         {
            File.WriteAllBytes(a_Filename, GetOutputStream().ToArray());
         }
         catch (Exception)
         {
            throw new Exception("Could not write " + a_Filename);
         }
      }

      public int Replace(string a_OldValue, string a_NewValue)
      {
         int occurences = 0;
         HSSFSheet a_sheet = null;

         Open();

         // Parse all sheets
         for (int SheetIndex = 0; SheetIndex < m_CurrWorkbook.NumberOfSheets; SheetIndex++)
         {
            //Get the sheet
            a_sheet = (HSSFSheet)m_CurrWorkbook.GetSheetAt(SheetIndex);

            //get the rows
            IEnumerator a_enum = a_sheet.GetRowEnumerator();

            //Parse the rows for the value
            while (a_enum.MoveNext())
            {
               HSSFRow arow = (HSSFRow)a_enum.Current;
               //The LastCellNum property indicates the far right cell
               //available in the sheet. DO NOT USE Cells.Count .
               for (int i = 0; i < arow.LastCellNum; i++)
               {
                  if (arow.GetCell(i) != null)
                  {
                     if (arow.GetCell(i).ToString().Equals(a_OldValue))
                     {
                        arow.GetCell(i).SetCellValue(a_NewValue);
                        occurences += 1;
                     }
                  }
               }//for
            }//while

         }//for
         return occurences;
      }

      public int Replace(string a_OldValue, string a_NewValue, string a_SheetName)
      {
         int occurences = 0;
         HSSFSheet a_sheet = null;

         Open();

         // Parse all sheets
         for (int SheetIndex = 0; SheetIndex < m_CurrWorkbook.NumberOfSheets; SheetIndex++)
         {
            //Get the sheet
            a_sheet = (HSSFSheet)m_CurrWorkbook.GetSheetAt(SheetIndex);

            if (a_sheet.SheetName.Equals(a_SheetName))
            {
               //get the rows
               IEnumerator a_enum = a_sheet.GetRowEnumerator();

               //Parse the rows for the value
               int aint = 0;
               while (a_enum.MoveNext())
               {
                  Console.WriteLine(aint);
                  aint += 1;
                  HSSFRow arow = (HSSFRow)a_enum.Current;
                  //The LastCellNum property indicates the far right cell
                  //available in the sheet.DO NOT USE Cells.Count.
                  for (int i = 0; i < arow.LastCellNum; i++)
                  {
                     if (arow.GetCell(i) != null)
                     {
                        if (arow.GetCell(i).ToString().Equals(a_OldValue))
                        {
                           arow.GetCell(i).SetCellValue(a_NewValue);
                           occurences += 1;
                        }
                     }
                  }//for
               }//while
            }//if
         }//for
         return occurences;
      }

      /// <summary>
      /// Sets a given cell's value.If the sheet referred to , does not exist
      /// then it will be created.
      /// </summary>
      /// <param name="a_Value">the cell's value</param>
      /// <param name="a_SheetName">the sheet's name in which values should be set</param>
      /// <param name="a_CellRow">the row number</param>
      /// <param name="a_CellColumn">the column number</param>
      public void SetCellValue(string a_Value, string a_SheetName, int a_CellRow, int a_CellColumn)
      {
         HSSFSheet a_Sheet = null;

         Open();

         a_Sheet = (HSSFSheet)m_CurrWorkbook.GetSheet(a_SheetName);

         //Check if the sheet exists.If not then create it.
         if (a_Sheet == null)
         {
            m_CurrWorkbook.CreateSheet(a_SheetName);
            a_Sheet = (HSSFSheet)m_CurrWorkbook.GetSheet(a_SheetName);
         }

         //Check if the row exists.If not then create it.
         if (a_Sheet.GetRow(a_CellRow) == null)
         {
            a_Sheet.CreateRow(a_CellRow);
         }

         //Check if the cell exists.If not then create it.
         if (a_Sheet.GetRow(a_CellRow).GetCell(a_CellColumn) == null)
         {
            a_Sheet.GetRow(a_CellRow).CreateCell(a_CellColumn);
         }

         //Finally set the value.
         
         a_Sheet.GetRow(a_CellRow).GetCell(a_CellColumn).SetCellValue(a_Value);
      }

      /// <summary>
      /// Updates a workbook with the datatable provided.
      /// </summary>
      /// <param name="a_Datatable"></param>
      /// <returns></returns>
      public bool UpdateWorkbook(DataTable a_Datatable)
      {
         string a_Value = "";

         if (a_Datatable == null)
         {
            throw new NullReferenceException("Datatable empty");
         }

         //Parse all rows 
         for (int rowCounter = 0; rowCounter < a_Datatable.Rows.Count; rowCounter++)
         {
            for (int columnCounter = 0; columnCounter < a_Datatable.Columns.Count; columnCounter++)
            {
               //if an exception is thrown then set an empty value.
               try
               {
                  a_Value = a_Datatable.Rows[rowCounter].ItemArray.GetValue(columnCounter).ToString();
               }
               catch (Exception)
               {
                  a_Value = "";
               }
               //Finally set the cell's value
               SetCellValue(a_Value, a_Datatable.TableName, rowCounter, columnCounter);

            }//for
         }//foreach

         //Shift all rows by +1 in order to add the column headers.
         m_CurrWorkbook.GetSheet(a_Datatable.TableName).ShiftRows(0,
             m_CurrWorkbook.GetSheet(a_Datatable.TableName).PhysicalNumberOfRows, 1);

         //Add the column headers
         for (int columnIndex = 0; columnIndex < a_Datatable.Columns.Count; columnIndex++)
         {
            SetCellValue(a_Datatable.Columns[columnIndex].ColumnName, a_Datatable.TableName, 0, columnIndex);
         }

         return true;
      }

      public bool Search(string a_Text)
      {
         throw new Exception("The method or operation is not implemented.");
      }

      
   }//END OF CLASS
}
