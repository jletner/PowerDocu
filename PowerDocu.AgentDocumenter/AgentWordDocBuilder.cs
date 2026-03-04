using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerDocu.Common;

namespace PowerDocu.AgentDocumenter
{
    class AgentWordDocBuilder : WordDocBuilder
    {
        private readonly AgentDocumentationContent content;

        public AgentWordDocBuilder(AgentDocumentationContent contentDocumentation, string template)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);
            string filename = InitializeWordDocument(content.folderPath + content.filename, template);
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(filename, true);
            mainPart = wordDocument.MainDocumentPart;
            body = mainPart.Document.Body;
            PrepareDocument(!String.IsNullOrEmpty(template));
            addAgentOverview();
            addAgentChannels();
            addAgentSettings();
            addAgentTopicsOverview();
            addAgentTopicDetails();
            NotificationHelper.SendNotification("Created Word documentation for " + contentDocumentation.filename);
        }

        private void addAgentOverview()
        {
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Agent - " + content.agent.Name));
            ApplyStyleToParagraph("Heading1", para);

            Table table = CreateTable();
            table.Append(CreateRow(new Text("Agent Name"), new Text(content.agent.Name)));
            if (!String.IsNullOrEmpty(content.agent.IconBase64))
            {
                try
                {
                    Bitmap agentLogo = ImageHelper.ConvertBase64ToBitmap(content.agent.IconBase64);
                    string logoPath = content.folderPath + $"agentlogo-{content.filename.Replace(" ", "-")}.png";
                    agentLogo.Save(logoPath);
                    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                    using (FileStream stream = new FileStream(logoPath, FileMode.Open))
                    {
                        imagePart.FeedData(stream);
                    }
                    Drawing icon = InsertImage(mainPart.GetIdOfPart(imagePart), 64, 64);
                    table.Append(CreateRow(new Text("Agent Logo"), icon));
                    agentLogo.Dispose();
                }
                catch { }
            }
            table.Append(CreateRow(new Text(content.headerDocumentationGenerated), new Text(PowerDocuReleaseHelper.GetTimestampWithVersion())));
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Details section
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Details));
            ApplyStyleToParagraph("Heading2", para);

            // Description
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Description));
            ApplyStyleToParagraph("Heading3", para);
            body.AppendChild(new Paragraph(CreateRunWithLinebreaks(content.agent.GetDescription() ?? "")));

            // Orchestration
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Orchestration));
            ApplyStyleToParagraph("Heading3", para);
            body.AppendChild(new Paragraph(new Run(new Text($"{content.OrchestrationText} - {content.agent.GetOrchestration()}"))));

            // Response Model
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.ResponseModel));
            ApplyStyleToParagraph("Heading3", para);
            body.AppendChild(new Paragraph(new Run(new Text(content.agent.GetResponseModel()))));

            // Instructions
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Instructions));
            ApplyStyleToParagraph("Heading3", para);
            body.AppendChild(new Paragraph(CreateRunWithLinebreaks(content.agent.GetInstructions() ?? "")));

            // Knowledge
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Knowledge));
            ApplyStyleToParagraph("Heading3", para);
            foreach (BotComponent knowledgeSource in content.agent.GetKnowledge())
            {
                body.AppendChild(new Paragraph(new Run(new Text(knowledgeSource.Name))));
            }

            // Suggested Prompts
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.SuggestedPrompts));
            ApplyStyleToParagraph("Heading3", para);
            body.AppendChild(new Paragraph(new Run(new Text(content.SuggestedPromptsText))));
            Dictionary<string, string> conversationStarters = content.agent.GetSuggestedPrompts();
            if (conversationStarters.Count > 0)
            {
                Table promptsTable = CreateTable();
                promptsTable.Append(CreateHeaderRow(new Text("Prompt Title"), new Text("Prompt")));
                foreach (var kvp in conversationStarters.OrderBy(x => x.Key))
                {
                    promptsTable.Append(CreateRow(new Text(kvp.Key), new Text(kvp.Value)));
                }
                body.Append(promptsTable);
            }
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addAgentTopicsOverview()
        {
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Topics));
            ApplyStyleToParagraph("Heading2", para);

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Name"), new Text("Type"), new Text("Trigger"), new Text("Kind")));
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name).ToList())
            {
                string topicType = topic.GetComponentTypeDisplayName();
                string triggerType = topic.GetTriggerTypeForTopic();
                string topicKind = topic.GetTopicKind() == "KnowledgeSourceConfiguration" ? "Knowledge" : triggerType;
                table.Append(CreateRow(new Text(topic.Name), new Text(topicType), new Text(triggerType), new Text(topicKind)));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addAgentTopicDetails()
        {
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name).ToList())
            {
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Topic: " + topic.Name));
                ApplyStyleToParagraph("Heading2", para);

                // Metadata table
                Table table = CreateTable();
                table.Append(CreateRow(new Text("Name"), new Text(topic.Name)));
                table.Append(CreateRow(new Text("Type"), new Text(topic.GetComponentTypeDisplayName())));
                table.Append(CreateRow(new Text("Trigger"), new Text(topic.GetTriggerTypeForTopic())));
                table.Append(CreateRow(new Text("Topic Kind"), new Text(topic.GetTopicKind())));
                if (!string.IsNullOrEmpty(topic.Description))
                {
                    table.Append(CreateRow(new Text("Description"), new Text(topic.Description)));
                }
                string modelDesc = topic.GetModelDescription();
                if (!string.IsNullOrEmpty(modelDesc))
                {
                    table.Append(CreateRow(new Text("Model Description"), new Text(modelDesc)));
                }
                string startBehavior = topic.GetStartBehavior();
                if (!string.IsNullOrEmpty(startBehavior))
                {
                    table.Append(CreateRow(new Text("Start Behavior"), new Text(startBehavior)));
                }
                body.Append(table);
                body.AppendChild(new Paragraph(new Run(new Break())));

                // Trigger queries
                List<string> triggerQueries = topic.GetTriggerQueries();
                if (triggerQueries.Count > 0)
                {
                    para = body.AppendChild(new Paragraph());
                    run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Trigger Queries"));
                    ApplyStyleToParagraph("Heading3", para);

                    Table triggerTable = CreateTable();
                    triggerTable.Append(CreateHeaderRow(new Text("Query")));
                    foreach (string query in triggerQueries)
                    {
                        triggerTable.Append(CreateRow(new Text(query)));
                    }
                    body.Append(triggerTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Knowledge source details
                if (topic.GetTopicKind() == "KnowledgeSourceConfiguration")
                {
                    var (sourceKind, skillConfig) = topic.GetKnowledgeSourceDetails();
                    para = body.AppendChild(new Paragraph());
                    run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Knowledge Source"));
                    ApplyStyleToParagraph("Heading3", para);

                    Table ksTable = CreateTable();
                    if (!string.IsNullOrEmpty(sourceKind))
                        ksTable.Append(CreateRow(new Text("Source Kind"), new Text(sourceKind)));
                    if (!string.IsNullOrEmpty(skillConfig))
                        ksTable.Append(CreateRow(new Text("Skill Configuration"), new Text(skillConfig)));
                    body.Append(ksTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Variables
                var variables = topic.GetTopicVariables();
                if (variables.Count > 0)
                {
                    para = body.AppendChild(new Paragraph());
                    run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Variables"));
                    ApplyStyleToParagraph("Heading3", para);

                    Table varTable = CreateTable();
                    varTable.Append(CreateHeaderRow(new Text("Variable"), new Text("Context")));
                    foreach (var (variable, context) in variables)
                    {
                        varTable.Append(CreateRow(new Text(variable), new Text(context)));
                    }
                    body.Append(varTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Topic flow diagram
                string graphFile = topic.getTopicFileName() + "-detailed.png";
                string graphFilePath = Path.Combine(content.folderPath, "Topics", graphFile);
                if (File.Exists(graphFilePath))
                {
                    para = body.AppendChild(new Paragraph());
                    run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Topic Flow"));
                    ApplyStyleToParagraph("Heading3", para);

                    try
                    {
                        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                        int imageWidth, imageHeight;
                        using (FileStream stream = new FileStream(graphFilePath, FileMode.Open))
                        {
                            using (var image = Image.FromStream(stream, false, false))
                            {
                                imageWidth = image.Width;
                                imageHeight = image.Height;
                            }
                            stream.Position = 0;
                            imagePart.FeedData(stream);
                        }
                        int usedWidth = (imageWidth > 600) ? 600 : imageWidth;
                        Drawing drawing = InsertImage(mainPart.GetIdOfPart(imagePart), usedWidth, (int)(usedWidth * imageHeight / imageWidth));
                        body.AppendChild(new Paragraph(new Run(drawing)));
                    }
                    catch { }
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }
            }
        }
        private void addAgentChannels()
        {
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Channels"));
            ApplyStyleToParagraph("Heading1", para);

            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Channels are not exported with the solution and are not documented automatically."));
        }

        private static readonly string NotInExportMessage = "This setting is not available in the solution export.";

        private void addAgentSettings()
        {
            var config = content.agent.Configuration;
            var ai = config?.aISettings;

            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Settings"));
            ApplyStyleToParagraph("Heading1", para);

            // Generative AI
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Generative AI"));
            ApplyStyleToParagraph("Heading2", para);

            Table genAiTable = CreateTable();
            genAiTable.Append(CreateHeaderRow(new Text("Setting"), new Text("Value")));
            genAiTable.Append(CreateRow(new Text("Generative Actions"), new Text(config?.settings?.GenerativeActionsEnabled == true ? "Enabled" : "Disabled")));
            genAiTable.Append(CreateRow(new Text("Use Model Knowledge"), new Text(ai?.useModelKnowledge == true ? "Yes" : "No")));
            genAiTable.Append(CreateRow(new Text("File Analysis"), new Text(ai?.isFileAnalysisEnabled == true ? "Enabled" : "Disabled")));
            genAiTable.Append(CreateRow(new Text("Semantic Search"), new Text(ai?.isSemanticSearchEnabled == true ? "Enabled" : "Disabled")));
            genAiTable.Append(CreateRow(new Text("Content Moderation"), new Text(ai?.contentModeration ?? "Unknown")));
            genAiTable.Append(CreateRow(new Text("Opt-in to Latest Models"), new Text(ai?.optInUseLatestModels == true ? "Yes" : "No")));
            body.Append(genAiTable);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Security
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Security"));
            ApplyStyleToParagraph("Heading2", para);

            Table secTable = CreateTable();
            secTable.Append(CreateHeaderRow(new Text("Setting"), new Text("Value")));
            secTable.Append(CreateRow(new Text("Authentication Mode"), new Text(content.agent.GetAuthenticationModeDisplayName())));
            secTable.Append(CreateRow(new Text("Authentication Trigger"), new Text(content.agent.GetAuthenticationTriggerDisplayName())));
            body.Append(secTable);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Connection settings
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Connection settings"));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Authoring canvas
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Authoring canvas"));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Entities
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Entities"));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Skills
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Skills"));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Voice
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Voice"));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Languages
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Languages"));
            ApplyStyleToParagraph("Heading2", para);

            Table langTable = CreateTable();
            langTable.Append(CreateHeaderRow(new Text("Setting"), new Text("Value")));
            langTable.Append(CreateRow(new Text("Primary Language"), new Text(content.agent.GetLanguageDisplayName())));
            body.Append(langTable);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Language understanding
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Language understanding"));
            ApplyStyleToParagraph("Heading2", para);

            Table luTable = CreateTable();
            luTable.Append(CreateHeaderRow(new Text("Setting"), new Text("Value")));
            luTable.Append(CreateRow(new Text("Recognizer"), new Text(content.agent.GetRecognizerDisplayName())));
            body.Append(luTable);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Component collections
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Component collections"));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Advanced
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Advanced"));
            ApplyStyleToParagraph("Heading2", para);

            Table advTable = CreateTable();
            advTable.Append(CreateHeaderRow(new Text("Setting"), new Text("Value")));
            advTable.Append(CreateRow(new Text("Template"), new Text(content.agent.Template ?? "")));
            advTable.Append(CreateRow(new Text("Runtime Provider"), new Text(content.agent.RuntimeProvider.ToString())));
            body.Append(advTable);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }
    }
}
