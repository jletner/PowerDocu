using System;
using System.IO;
using PowerDocu.Common;

namespace PowerDocu.BPFDocumenter
{
    public static class BPFDocumentationGenerator
    {
        public static void GenerateOutput(DocumentationContext context, string path)
        {
            if (context.BusinessProcessFlows == null || context.BusinessProcessFlows.Count == 0 || !context.Config.documentBusinessProcessFlows) return;

            DateTime startDocGeneration = DateTime.Now;
            NotificationHelper.SendNotification($"Found {context.BusinessProcessFlows.Count} Business Process Flow(s) in the solution.");

            if (context.FullDocumentation)
            {
                foreach (BPFEntity bpf in context.BusinessProcessFlows)
                {
                    BPFDocumentationContent content = new BPFDocumentationContent(bpf, path, context);

                    string wordTemplate = (!String.IsNullOrEmpty(context.Config.wordTemplate) && File.Exists(context.Config.wordTemplate))
                        ? context.Config.wordTemplate : null;
                    if (context.Config.outputFormat.Equals(OutputFormatHelper.Word) || context.Config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating Word documentation for Business Process Flow: " + bpf.GetDisplayName());
                        BPFWordDocBuilder wordDoc = new BPFWordDocBuilder(content, wordTemplate);
                    }
                    if (context.Config.outputFormat.Equals(OutputFormatHelper.Markdown) || context.Config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating Markdown documentation for Business Process Flow: " + bpf.GetDisplayName());
                        BPFMarkdownBuilder markdownDoc = new BPFMarkdownBuilder(content);
                    }
                    if (context.Config.outputFormat.Equals(OutputFormatHelper.Html) || context.Config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating HTML documentation for Business Process Flow: " + bpf.GetDisplayName());
                        BPFHtmlBuilder htmlDoc = new BPFHtmlBuilder(content);
                    }
                }
            }

            DateTime endDocGeneration = DateTime.Now;
            NotificationHelper.SendNotification(
                $"BPFDocumenter: Processed {context.BusinessProcessFlows.Count} Business Process Flow(s) in {(endDocGeneration - startDocGeneration).TotalSeconds} seconds."
            );
        }
    }
}
