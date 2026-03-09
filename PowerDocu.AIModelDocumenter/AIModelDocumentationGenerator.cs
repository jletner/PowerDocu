using System;
using System.Collections.Generic;
using System.IO;
using PowerDocu.Common;

namespace PowerDocu.AIModelDocumenter
{
    public static class AIModelDocumentationGenerator
    {
        /// <summary>
        /// Generates documentation output for AI Models found in the solution customizations.
        /// </summary>
        public static void GenerateOutput(DocumentationContext context, string path)
        {
            List<AIModel> aiModels = context.Customizations?.getAIModels();
            if (aiModels == null || aiModels.Count == 0) return;

            DateTime startDocGeneration = DateTime.Now;
            NotificationHelper.SendNotification($"Found {aiModels.Count} AI Model(s) in the solution.");

            if (context.FullDocumentation)
            {
                foreach (AIModel aiModel in aiModels)
                {
                    AIModelDocumentationContent content = new AIModelDocumentationContent(aiModel, path, context);

                    string wordTemplate = (!String.IsNullOrEmpty(context.Config.wordTemplate) && File.Exists(context.Config.wordTemplate))
                        ? context.Config.wordTemplate : null;
                    if (context.Config.outputFormat.Equals(OutputFormatHelper.Word) || context.Config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating Word documentation for AI Model: " + aiModel.getName());
                        AIModelWordDocBuilder wordDoc = new AIModelWordDocBuilder(content, wordTemplate);
                    }
                    if (context.Config.outputFormat.Equals(OutputFormatHelper.Markdown) || context.Config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating Markdown documentation for AI Model: " + aiModel.getName());
                        AIModelMarkdownBuilder markdownDoc = new AIModelMarkdownBuilder(content);
                    }
                    if (context.Config.outputFormat.Equals(OutputFormatHelper.Html) || context.Config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating HTML documentation for AI Model: " + aiModel.getName());
                        AIModelHtmlBuilder htmlDoc = new AIModelHtmlBuilder(content);
                    }
                }
            }

            DateTime endDocGeneration = DateTime.Now;
            NotificationHelper.SendNotification(
                $"AIModelDocumenter: Processed {aiModels.Count} AI Model(s) in {(endDocGeneration - startDocGeneration).TotalSeconds} seconds."
            );
        }
    }
}
