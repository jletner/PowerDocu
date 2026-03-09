using System.IO;
using PowerDocu.Common;

namespace PowerDocu.AIModelDocumenter
{
    class AIModelDocumentationContent
    {
        public string folderPath, filename;
        public AIModel aiModel;
        public DocumentationContext context;

        public AIModelDocumentationContent(AIModel aiModel, string path, DocumentationContext context = null)
        {
            NotificationHelper.SendNotification("Preparing documentation content for AI Model: " + aiModel.getName());
            this.aiModel = aiModel;
            this.context = context;
            folderPath = path + CharsetHelper.GetSafeName(@"\AIModelDoc " + aiModel.getName() + @"\");
            Directory.CreateDirectory(folderPath);
            filename = CharsetHelper.GetSafeName(aiModel.getName());
        }
    }
}
