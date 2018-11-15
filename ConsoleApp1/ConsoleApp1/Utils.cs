using System;
using System.Collections.Generic;
using System.Text;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using MyBot;

namespace MCSTDiscord
{
    public static class Utils
    {
        /// <summary>
        /// Creates a string dictionary so you an easily pull out values from cells - originally intended for info command
        /// </summary>
        /// <param name="user">the target in the command</param>
        /// <param name="columnNames">Option to pass in specific column names to have</param>
        /// <returns></returns>
        public static Dictionary<string, string> GetInfoFromColumns(string user, string[] columnNames = null)
        {   //Dictionary
            Dictionary<string, string> returnDictionary = new Dictionary<string, string>();
            //find user -TODO - Make this more robust and search other users like grep.
            if (user == "me")
            {
                user = Program.context.User.ToString();
            }
            //fix range for this user
            string range = GetRowByDiscord(user).ToString();
            range = String.Format("B{0}:F", range);
            //regular get
            SpreadsheetsResource.ValuesResource.GetRequest get = Commands.service.Spreadsheets.Values.Get(Commands.sheetid, range);
            ValueRange response = get.Execute();
            //format key /stringvalue pair
            var i = 0;
            foreach (var value in response.Values)
            {
                foreach (var innervalue in value)
                {
                    returnDictionary.Add(("value" + i++).ToString(), innervalue.ToString());
                }
            }
            return returnDictionary;

        }

        public static int GetRowByDiscord(string discordhandle)
        {
            string range = "B2:B";
            SpreadsheetsResource.ValuesResource.GetRequest get = Commands.service.Spreadsheets.Values.Get(Commands.sheetid, range);
            ValueRange response = get.Execute();
            if (response.Values == null) { return 0; }
            IList<IList<object>> values = response.Values;
            int rownum = 2;
            foreach (var row in values)
            {
                foreach (var cell in row)
                {
                    if ((string)cell == discordhandle)
                    {
                        return rownum;
                    }
                    rownum++;
                }
            }
            return 0;
        }
        public static int GetRowByName(string name)
        {
            string range = "A2:A";
            SpreadsheetsResource.ValuesResource.GetRequest get = Commands.service.Spreadsheets.Values.Get(Commands.sheetid, range);
            ValueRange response = get.Execute();
            if (response.Values == null) { return 0; }
            IList<IList<object>> values = response.Values;
            int rownum = 2;
            foreach (var row in values)
            {
                foreach (var cell in row)
                {
                    if ((string)cell == name)
                    {
                        return rownum;
                    }
                    rownum++;
                }
            }
            return 0;
        }

        public static int GetRowByAlt(int altnum, string name)
        {
            string range = null;
            if (altnum > 3 || altnum == 0) { return 0; }
            switch (altnum)
            {
                case 1:
                    range = "G2:G";
                    break;
                case 2:
                    range = "I2:I";
                    break;
            }
            SpreadsheetsResource.ValuesResource.GetRequest get = Commands.service.Spreadsheets.Values.Get(Commands.sheetid, range);
            ValueRange response = get.Execute();
            if (response.Values == null) { return 0; }
            IList<IList<object>> values = response.Values;
            int rownum = 2;
            foreach (var row in values)
            {
                foreach (var cell in row)
                {
                    if ((string)cell == name)
                    {
                        return rownum;
                    }
                    rownum++;
                }
            }
            return 0;
        }
        public static bool CellEmpty(string range)
        {
            SpreadsheetsResource.ValuesResource.GetRequest get = Commands.service.Spreadsheets.Values.Get(Commands.sheetid, range);
            ValueRange response = get.Execute();
            if (response.Values == null) { return true; }
            IList<IList<object>> values = response.Values;
            int rownum = 2;
            foreach (var row in values)
            {
                foreach (var cell in row)
                {
                    if ((string)cell == "")
                    {
                        return true;
                    }
                    else { return false; }
                }
            }
            return false;
        }
    }
}
