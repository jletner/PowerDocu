using System.IO;
using PowerDocu.Common;

namespace PowerDocu.BPFDocumenter
{
    public class BPFDocumentationContent
    {
        public string folderPath, filename;
        public BPFEntity bpf;
        public DocumentationContext context;

        public string headerOverview = "Overview";
        public string headerStages = "Stages";
        public string headerStageDetails = "Stage Details";
        public string headerProperties = "Properties";
        public string headerDocumentationGenerated = "Documentation generated at";

        public BPFDocumentationContent(BPFEntity bpf, string path, DocumentationContext context)
        {
            NotificationHelper.SendNotification("Preparing documentation content for Business Process Flow: " + bpf.GetDisplayName());
            this.bpf = bpf;
            this.context = context;
            folderPath = path + CharsetHelper.GetSafeName(@"\BPFDoc " + bpf.GetDisplayName() + @"\");
            Directory.CreateDirectory(folderPath);
            filename = CharsetHelper.GetSafeName(bpf.GetDisplayName());
        }

        public string GetTableDisplayName(string schemaName)
        {
            if (string.IsNullOrEmpty(schemaName)) return schemaName;
            return context?.GetTableDisplayName(schemaName) ?? schemaName;
        }
    }
}
