using ExcelDataReader;
using Microsoft.TeamFoundation.VersionControl.Client;
using System;
using System.Data;
using System.IO;

namespace KamenokoSoft.ExcelDiff
{
    public class ExcelDiffToolProvider : IToolProvider
    {
        public string Extension => ".xlsx";

        public ToolOperations Operation => ToolOperations.Compare;

        public bool CanOperateOnInMemoryFiles => true;

        public IToolExecutionResult Execute(EventHandler exitHandler, AdvancedToolParameters advancedParameters, params string[] arguments)
        {
            var originalFile = this.CreateTempTsvFile(arguments[0]);
            var modifiedFile = this.CreateTempTsvFile(arguments[1]);

            var sourceFileTag = arguments[2];
            var targetFileTag = arguments[3];
            var sourceFileLabel = arguments[5];
            var targetFileLabel = arguments[6];

            Difference.VisualDiffFiles(originalFile, modifiedFile, sourceFileTag, targetFileTag, sourceFileLabel, targetFileLabel, true, true, true, true);

            return new EmptyToolExecution();
        }

        public string CreateTempTsvFile(string fileName)
        {
            var tmpFilePath = Path.GetTempFileName();

            FileStream stream;

            try
            {
                stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
            }
            catch (IOException)
            {
                // second attempt: copy file and read
                var tmpCopyFilePath = Path.GetTempFileName();
                File.Copy(fileName, tmpCopyFilePath, true);

                stream = File.Open(tmpCopyFilePath, FileMode.Open, FileAccess.Read);
            }

            var writer = new StreamWriter(tmpFilePath, false, System.Text.Encoding.UTF8);
            var reader = ExcelReaderFactory.CreateReader(stream);

            try
            {
                var xlsxData = reader.AsDataSet();

                foreach (DataTable table in xlsxData.Tables)
                {
                    foreach (DataRow row in table.Rows)
                    {
                        writer.Write($"[{table.TableName}]\t");

                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            writer.Write($"{row[i]}\t");
                        }

                        writer.WriteLine();
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                reader.Dispose();
                writer.Close();
                stream?.Dispose();
            }

            return tmpFilePath;
        }
    }

    public class EmptyToolExecution : IToolExecutionResult
    {
        public int Id => 0;

        public string Name => "";

        public bool HasExited => true;

        public int ExitCode => 0;

        public string ExitMessage => "";

        public bool PromptUserForMergeConfirmation => false;

        public void Cancel()
        {
        }

        public void WaitForOperationEnd()
        {
        }
    }
}
