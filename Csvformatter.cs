

using System;
using System.Data;
using System.Data.OleDb;

using System.IO;
using System.Windows.Forms;

namespace csvformattingtools
{
  public class Csvformatter
  {
    public static DataTable GetCsvFile(string fileName)
    {
      try
      {
        string directoryName = Path.GetDirectoryName(fileName);
        string fileName1 = Path.GetFileName(fileName);
        using (OleDbConnection oleDbConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + directoryName + ";Extended Properties=Text;"))
        {
          using (OleDbCommand selectCommand = new OleDbCommand())
          {
            if (fileName1 == null)
              return new DataTable();
            selectCommand.CommandText = "SELECT * FROM " + fileName1;
            selectCommand.Connection = oleDbConnection;
            using (OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(selectCommand))
            {
              oleDbConnection.Open();
              DataTable dataTable = new DataTable();
              oleDbDataAdapter.Fill(dataTable);
              return dataTable;
            }
          }
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
        return new DataTable();
      }
    }

    public static void CreateCsvFile(DataTable dt, string strFilePath)
    {
      try
      {
        using (StreamWriter streamWriter = new StreamWriter(strFilePath, false))
        {
          int count = dt.Columns.Count;
          for (int index = 0; index < count; ++index)
          {
            streamWriter.Write((object) dt.Columns[index]);
            if (index < count - 1)
              streamWriter.Write(",");
          }
          streamWriter.Write(streamWriter.NewLine);
          foreach (DataRow dataRow in (InternalDataCollectionBase) dt.Rows)
          {
            for (int index = 0; index < count; ++index)
            {
              if (!Convert.IsDBNull(dataRow[index]))
              {
                streamWriter.Write("\"");
                streamWriter.Write(dataRow[index].ToString());
                streamWriter.Write("\"");
              }
              if (index < count - 1)
                streamWriter.Write(",");
            }
            streamWriter.Write(streamWriter.NewLine);
          }
        }
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    public static void Checkcommas(string filename)
    {
      string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
      using (StreamWriter streamWriter1 = new StreamWriter(folderPath + "\\importthis.csv"))
      {
        using (StreamWriter streamWriter2 = new StreamWriter(folderPath + "\\importissues.csv"))
        {
          using (StreamReader streamReader = new StreamReader(filename))
          {
            string str;
            while ((str = streamReader.ReadLine()) != null)
            {
              bool flag = true;
              int num = 0;
              foreach (char ch in str)
              {
                if ((int) ch == 34)
                  flag = !flag;
                else if (flag && (int) ch == 44)
                  ++num;
              }
              if (num != 11 && num != 13 && num != 19)
                streamWriter2.WriteLine(str);
              else
                streamWriter1.WriteLine(str);
            }
          }
        }
      }
    }

    public static void ParseFullName(string fullName, out string firstName, out string middleInitial, out string lastName)
    {
      string str1 = string.Empty;
      middleInitial = "";
      if (fullName != "" || fullName != string.Empty)
      {
        string str2 = fullName;
        if (str2.Length - str2.Replace(",", "").Length <= 1)
        {
          string[] strArray = fullName.Split(',');
          lastName = strArray[0];
          string str3 = strArray[1];
          char[] chArray = new char[1]
          {
            ' '
          };
          foreach (string str4 in str3.Split(chArray))
          {
            if (str4.Length > 1)
              str1 = str1 + str4 + " ";
            else if (str4.Length == 1)
              middleInitial = str4;
          }
          firstName = str1;
        }
        else
        {
          string[] strArray = fullName.Split(',');
          lastName = strArray[0] + " " + strArray[1];
          string str3 = strArray[2];
          char[] chArray = new char[1]
          {
            ' '
          };
          foreach (string str4 in str3.Split(chArray))
          {
            if (str4.Length > 1)
              str1 = str1 + str4 + " ";
            else if (str4.Length == 1)
              middleInitial = str4;
          }
          firstName = str1;
        }
      }
      else
      {
        firstName = "";
        middleInitial = "";
        lastName = "";
      }
    }

    private static DataTable SetColumnOrdinals(DataTable temp)
    {
      temp.Columns["PERSONAL_NUM"].SetOrdinal(0);
      temp.Columns["FIRST_NAME"].SetOrdinal(1);
      temp.Columns["MIDDLE_INITIAL"].SetOrdinal(2);
      temp.Columns["LAST_NAME"].SetOrdinal(3);
      temp.Columns["ADDRESS"].SetOrdinal(4);
      temp.Columns["ADDRESS2"].SetOrdinal(5);
      temp.Columns["CITY"].SetOrdinal(6);
      temp.Columns["STATE"].SetOrdinal(7);
      temp.Columns["ZIP"].SetOrdinal(8);
      temp.Columns["HOME_PHONE"].SetOrdinal(9);
      temp.Columns["OFFICE_PHONE"].SetOrdinal(10);
      temp.Columns["FAX_NUM"].SetOrdinal(11);
      temp.Columns["PAGER_NUM"].SetOrdinal(12);
      temp.Columns["MOBILE_PHONE"].SetOrdinal(13);
      temp.Columns["EMAIL_ADDRESS"].SetOrdinal(14);
      temp.Columns["DEPT"].SetOrdinal(15);
      temp.Columns["SALUTATION"].SetOrdinal(16);
      temp.Columns["TITLE"].SetOrdinal(17);
      return temp;
    }

    public static DataTable FormatForKeyWizardImport(DataTable temp)
    {
      try
      {
        switch (temp.Columns.Count)
        {
          case 12:
            temp.Columns.Add("MIDDLE_INITIAL", typeof (string));
            temp.Columns.Add("LAST_NAME", typeof (string));
            temp.Columns.Add("OFFICE_PHONE", typeof (string));
            temp.Columns.Add("FAX_NUM", typeof (string));
            temp.Columns.Add("PAGER_NUM", typeof (string));
            temp.Columns.Add("MOBILE_PHONE", typeof (string));
            temp.Columns.Add("EMAIL_ADDRESS", typeof (string));
            temp.Columns.Add("SALUTATION", typeof (string));
            temp.Columns.Remove("Date of Birth");
            temp.Columns.Remove("Level 2");
            temp.Columns["Level 1"].ColumnName = "DEPT";
            temp.Columns["Position Title"].ColumnName = "TITLE";
            temp.Columns["Employee Number"].ColumnName = "PERSONAL_NUM";
            temp.Columns["Name - Complete"].ColumnName = "FIRST_NAME";
            temp.Columns["Address Line 1"].ColumnName = "ADDRESS";
            temp.Columns["Address Line 2"].ColumnName = "ADDRESS2";
            temp.Columns["City"].ColumnName = "Bleh";
            temp.Columns["Bleh"].ColumnName = "CITY";
            temp.Columns["State/Prov"].ColumnName = "STATE";
            temp.Columns["Zip Code"].ColumnName = "ZIP";
            temp.Columns["Telephone"].ColumnName = "HOME_PHONE";
            foreach (DataRow dataRow in (InternalDataCollectionBase) temp.Rows)
            {
              string firstName;
              string middleInitial;
              string lastName;
              Csvformatter.ParseFullName(dataRow["FIRST_NAME"].ToString(), out firstName, out middleInitial, out lastName);
              dataRow["FIRST_NAME"] = (object) firstName;
              dataRow["MIDDLE_INITIAL"] = (object) middleInitial;
              dataRow["LAST_NAME"] = (object) lastName;
            }
            temp = Csvformatter.SetColumnOrdinals(temp);
            break;
          case 14:
            temp.Columns.Add("MIDDLE_INITIAL", typeof (string));
            temp.Columns.Add("LAST_NAME", typeof (string));
            temp.Columns.Add("OFFICE_PHONE", typeof (string));
            temp.Columns.Add("FAX_NUM", typeof (string));
            temp.Columns.Add("PAGER_NUM", typeof (string));
            temp.Columns.Add("MOBILE_PHONE", typeof (string));
            temp.Columns.Add("EMAIL_ADDRESS", typeof (string));
            temp.Columns.Add("SALUTATION", typeof (string));
            temp.Columns.Remove("Date of Birth");
            temp.Columns.Remove("Dept#");
            temp.Columns.Remove("ADOH");
            temp.Columns.Remove("ER#");
            temp.Columns["Telephone"].ColumnName = "HOME_PHONE";
            temp.Columns["Employee Number"].ColumnName = "PERSONAL_NUM";
            temp.Columns["Division"].ColumnName = "DEPT";
            temp.Columns["Address Line 1"].ColumnName = "ADDRESS";
            temp.Columns["Address Line 2"].ColumnName = "ADDRESS2";
            temp.Columns["City"].ColumnName = "Bleh";
            temp.Columns["Bleh"].ColumnName = "CITY";
            temp.Columns["State"].ColumnName = "STATE";
            temp.Columns["Zip Code"].ColumnName = "ZIP";
            temp.Columns["Position Title"].ColumnName = "TITLE";
            temp.Columns["Complete Name"].ColumnName = "FIRST_NAME";
            foreach (DataRow dataRow in (InternalDataCollectionBase) temp.Rows)
            {
              string firstName;
              string middleInitial;
              string lastName;
              Csvformatter.ParseFullName(dataRow["FIRST_NAME"].ToString(), out firstName, out middleInitial, out lastName);
              dataRow["FIRST_NAME"] = (object) firstName;
              dataRow["MIDDLE_INITIAL"] = (object) middleInitial;
              dataRow["LAST_NAME"] = (object) lastName;
            }
            temp = Csvformatter.SetColumnOrdinals(temp);
            break;
          case 20:
            temp.Columns.Remove("DateCreated");
            temp.Columns.Remove("DateHired");
            temp.Columns.Remove("DateModified");
            temp.Columns.Remove("Division");
            temp.Columns.Remove("PropertyGUID");
            temp.Columns.Remove("SupervisorLevel");
            temp.Columns.Remove("GamingRelated");
            temp.Columns.Remove("SundayOff");
            temp.Columns.Remove("MondayOff");
            temp.Columns.Remove("TuesdayOff");
            temp.Columns.Remove("WednesdayOff");
            temp.Columns.Remove("ThursdayOff");
            temp.Columns.Remove("FridayOff");
            temp.Columns.Remove("SaturdayOff");
            temp.Columns["EmployeeID"].ColumnName = "PERSONAL_NUM";
            temp.Columns["Firstname"].ColumnName = "FIRST_NAME";
            temp.Columns["Middlename"].ColumnName = "MIDDLE_INITIAL";
            temp.Columns["Lastname"].ColumnName = "LAST_NAME";
            temp.Columns["Department"].ColumnName = "DEPT";
            temp.Columns["JobPosition"].ColumnName = "TITLE";
            temp.Columns.Add("OFFICE_PHONE", typeof (string));
            temp.Columns.Add("FAX_NUM", typeof (string));
            temp.Columns.Add("PAGER_NUM", typeof (string));
            temp.Columns.Add("MOBILE_PHONE", typeof (string));
            temp.Columns.Add("EMAIL_ADDRESS", typeof (string));
            temp.Columns.Add("SALUTATION", typeof (string));
            temp.Columns.Add("ADDRESS", typeof (string));
            temp.Columns.Add("ADDRESS2", typeof (string));
            temp.Columns.Add("CITY", typeof (string));
            temp.Columns.Add("STATE", typeof (string));
            temp.Columns.Add("ZIP", typeof (string));
            temp.Columns.Add("HOME_PHONE", typeof (string));
            temp = Csvformatter.SetColumnOrdinals(temp);
            break;
        }
        return temp;
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
        return new DataTable();
      }
    }

    public static DataTable AddAddressessToDataTable(DataTable original, DataTable emaildt)
    {
       
        //getting null results here (exception error)
      //  var drresults = from myrow in original.AsEnumerable() select myrow;
        
      //  foreach (DataRow x in drresults)
      //  {
      //      string employeeid = x["PERSONAL_NUM"].ToString();
      //   var results = from myRow in emaildt.AsEnumerable()    where myRow.Field<string>("Name").Equals(employeeid)   select myRow.Field<string>("PrimarySmtpAddress");

      //// x["EMAIL_ADDRESS"] = 0;


      //  }









        return emaildt;
    }

  }
}
