using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using EloGroup.ConvertWordToPDF.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;

namespace EloGroup.ConvertWordToPDF.Activities
{
    [LocalizedDisplayName(nameof(Resources.ConvertDocToPDF_DisplayName))]
    [LocalizedDescription(nameof(Resources.ConvertDocToPDF_Description))]
    public class ConvertDocToPDF : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.Timeout_DisplayName))]
        [LocalizedDescription(nameof(Resources.Timeout_Description))]
        public InArgument<int> TimeoutMS { get; set; } = 60000;

        [LocalizedDisplayName(nameof(Resources.ConvertDocToPDF_FileNameDoc_DisplayName))]
        [LocalizedDescription(nameof(Resources.ConvertDocToPDF_FileNameDoc_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> FileNameDoc { get; set; }

        [LocalizedDisplayName(nameof(Resources.ConvertDocToPDF_FileNamePDF_DisplayName))]
        [LocalizedDescription(nameof(Resources.ConvertDocToPDF_FileNamePDF_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> FileNamePDF { get; set; }

        #endregion


        #region Constructors

        public ConvertDocToPDF()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (FileNameDoc == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(FileNameDoc)));
            if (FileNamePDF == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(FileNamePDF)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var timeout = TimeoutMS.Get(context);
            var filenameDoc = FileNameDoc.Get(context);
            var filenamePDF = FileNamePDF.Get(context);

            // Set a timeout on the execution
            var task = ExecuteWithTimeout(context, cancellationToken);
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);

            // Outputs
            return (ctx) => {
            };
        }

        private async Task ExecuteWithTimeout(AsyncCodeActivityContext context, CancellationToken cancellationToken = default)
        {
            var filenameDoc = FileNameDoc.Get(context);
            var filenamePDF = FileNamePDF.Get(context);

            // Create a new Microsoft Word application object
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            // C# doesn't have optional arguments so we'll need a dummy value
            object oMissing = System.Reflection.Missing.Value;

            // Get list of Word files in specified directory

            System.IO.FileInfo wordFile = new System.IO.FileInfo(filenameDoc);

            word.Visible = false;
            word.ScreenUpdating = false;

            // Cast as Object for word Open method
            Object filename = (Object)wordFile.FullName;

            // Use the dummy value as a placeholder for optional arguments
            Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref filename, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            doc.Activate();

            object outputFileName = filenamePDF;
            object fileFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;

            // Save document into PDF Format
            doc.SaveAs(ref outputFileName,
                ref fileFormat, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            // Close the Word document, but leave the Word application open.
            // doc has to be cast to type _Document so that it will find the
            // correct Close method.                
            object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
            doc.Close(ref saveChanges, ref oMissing, ref oMissing);
            doc = null;


            // word has to be cast to type _Application so that it will find
            // the correct Quit method.
            word.Quit(ref oMissing, ref oMissing, ref oMissing);
            word = null;
        }

        #endregion
    }
}

