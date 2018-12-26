using CsvHelper;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointContentOverview
{
    class Program
    {
        static void Main(string[] args)
        {
            string siteUrl = "https://YourSiteName.sharepoint.com/"; //site URL of the tenant
            OfficeDevPnP.Core.AuthenticationManager authManager = new OfficeDevPnP.Core.AuthenticationManager();
            ClientContext context = authManager.GetWebLoginClientContext(siteUrl);

            DataTable resultDataTable = new DataTable();

            DataColumn title = new DataColumn("Title");
            DataColumn path = new DataColumn("Path");
            DataColumn lastmodifieddate = new DataColumn("Last Modified Date");
            DataColumn lastmodifiedby = new DataColumn("Last Modified By");

            resultDataTable.Columns.Add(title);
            resultDataTable.Columns.Add(path);
            resultDataTable.Columns.Add(lastmodifieddate);
            resultDataTable.Columns.Add(lastmodifiedby);

            int currentRowIndex = 0;

            ResultTable resultTable = getSearchResults(context, currentRowIndex);

            if (null != resultTable && resultTable.TotalRows > 0)
            {
                while (resultTable.TotalRows > resultDataTable.Rows.Count)
                {
                    foreach (var resultRow in resultTable.ResultRows)
                    {
                        DataRow row = resultDataTable.NewRow();
                        row["Title"] = resultRow["Title"];
                        row["Path"] = resultRow["Path"];
                        row["Last Modified Date"] = resultRow["LastModifiedTime"];
                        row["Last Modified By"] = resultRow["EditorOWSUSER"];

                        resultDataTable.Rows.Add(row);
                    }

                    //Update the current row index
                    currentRowIndex = resultDataTable.Rows.Count;

                    resultTable = null;

                    resultTable = getSearchResults(context, currentRowIndex);

                    if (null != resultTable && resultTable.TotalRows > 0)
                    {
                        if (resultTable.RowCount <= 0)
                            break;
                    }
                    else
                        break;
                }
                exportCSV(resultDataTable);
            }
        }

        public static ResultTable getSearchResults(ClientContext clientContext, int startIndex)
        {
            KeywordQuery keywordQuery = new KeywordQuery(clientContext);

            keywordQuery.QueryText = " Your Search Query ";

            keywordQuery.SelectProperties.Add("Title");
            keywordQuery.SelectProperties.Add("Path");
            keywordQuery.SelectProperties.Add("LastModifiedTime");
            keywordQuery.SelectProperties.Add("EditorOWSUSER");

            //Specify the number of rows to return, 500 is MAX
            keywordQuery.RowLimit = 500;
            //Specify the number of rows to return in a page, 500 is MAX
            keywordQuery.RowsPerPage = 500;
            //Whether to remove duplicate results or not
            keywordQuery.TrimDuplicates = true;
            //Specify the timeout
            keywordQuery.Timeout = 10000; //10 minutes

            SearchExecutor searchExecutor = new SearchExecutor(clientContext);

            ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(keywordQuery);

            clientContext.ExecuteQuery();

            return results.Value.FirstOrDefault(v => v.TableType.Equals(KnownTableTypes.RelevantResults));
        }

        public static void exportCSV(DataTable DT)
        {
            //writing contents of datatable into memory stream
            MemoryStream ms = new MemoryStream();
            StreamWriter sw = new StreamWriter(ms);

            using (var csv = new CsvWriter(sw))
            {
                // Write columns
                foreach (DataColumn column in DT.Columns)
                {
                    csv.WriteField(column.ColumnName);
                }
                csv.NextRecord();

                // Write row values
                foreach (DataRow row in DT.Rows)
                {
                    for (var i = 0; i < DT.Columns.Count; i++)
                    {
                        csv.WriteField(row[i]);
                    }
                    csv.NextRecord();
                }

                sw.Flush();
                ms.Position = 0;  // read from the start of what was written             

                FileStream file = new FileStream("c:\\file.csv", FileMode.Create, FileAccess.Write);

                ms.WriteTo(file);
                file.Close();
                ms.Close();
            }

        }
    }
}
