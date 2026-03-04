using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using PowerDocu.Common;
using Grynwald.MarkdownGenerator;
using Svg;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PowerDocu.AgentDocumenter
{
    class AgentMarkdownBuilder : MarkdownBuilder
    {
        private readonly AgentDocumentationContent content;
        private readonly string mainDocumentFileName, knowledgeFileName, toolsFileName, agentsFileName, topicsFileName, channelsFileName, settingsFileName;
        private readonly MdDocument mainDocument, knowledgeDocument, toolsDocument, agentsDocument, topicsDocument, channelsDocument, settingsDocument;
        private readonly Dictionary<string, MdDocument> topicsDocuments = new Dictionary<string, MdDocument>();
        private readonly DocumentSet<MdDocument> set;
        private MdTable metadataTable;

        public AgentMarkdownBuilder(AgentDocumentationContent contentdocumentation)
        {
            content = contentdocumentation;
            Directory.CreateDirectory(content.folderPath);
            mainDocumentFileName = ("index " + content.filename + ".md").Replace(" ", "-");
            knowledgeFileName = ("knowledge " + content.filename + ".md").Replace(" ", "-");
            toolsFileName = ("tools " + content.filename + ".md").Replace(" ", "-");
            agentsFileName = ("agents " + content.filename + ".md").Replace(" ", "-");
            topicsFileName = ("topics " + content.filename + ".md").Replace(" ", "-");
            channelsFileName = ("channels " + content.filename + ".md").Replace(" ", "-");
            settingsFileName = ("settings " + content.filename + ".md").Replace(" ", "-");
            set = new DocumentSet<MdDocument>();
            mainDocument = set.CreateMdDocument(mainDocumentFileName);
            knowledgeDocument = set.CreateMdDocument(knowledgeFileName);
            toolsDocument = set.CreateMdDocument(toolsFileName);
            agentsDocument = set.CreateMdDocument(agentsFileName);
            //a dedicated document for each topic
            topicsDocument = set.CreateMdDocument(topicsFileName);
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name).ToList())
            {
                topicsDocuments.Add(topic.getTopicFileName(), set.CreateMdDocument("Topics/" + ("topic " + topic.getTopicFileName() + " " + content.filename + ".md").Replace(" ", "-")));
            }
            channelsDocument = set.CreateMdDocument(channelsFileName);
            settingsDocument = set.CreateMdDocument(settingsFileName);
            addAgentOverview();
            addAgentKnowledgeInfo();
            addAgentTools();
            addAgentAgentsInfo();
            addAgentTopics();
            addAgentChannels();
            addAgentSettings();
            set.Save(content.folderPath);
            NotificationHelper.SendNotification("Created Markdown documentation for " + content.filename);
        }

        private static readonly string NotInExportMessage = "This setting is not available in the solution export.";

        private void addAgentSettings()
        {
            var config = content.agent.Configuration;
            var ai = config?.aISettings;

            // Generative AI
            settingsDocument.Root.Add(new MdHeading("Generative AI", 2));
            List<MdTableRow> genAiRows = new List<MdTableRow>
            {
                new MdTableRow("Generative Actions", config?.settings?.GenerativeActionsEnabled == true ? "Enabled" : "Disabled"),
                new MdTableRow("Use Model Knowledge", ai?.useModelKnowledge == true ? "Yes" : "No"),
                new MdTableRow("File Analysis", ai?.isFileAnalysisEnabled == true ? "Enabled" : "Disabled"),
                new MdTableRow("Semantic Search", ai?.isSemanticSearchEnabled == true ? "Enabled" : "Disabled"),
                new MdTableRow("Content Moderation", ai?.contentModeration ?? "Unknown"),
                new MdTableRow("Opt-in to Latest Models", ai?.optInUseLatestModels == true ? "Yes" : "No")
            };
            settingsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Setting", "Value" }), genAiRows));

            // Security
            settingsDocument.Root.Add(new MdHeading("Security", 2));
            List<MdTableRow> secRows = new List<MdTableRow>
            {
                new MdTableRow("Authentication Mode", content.agent.GetAuthenticationModeDisplayName()),
                new MdTableRow("Authentication Trigger", content.agent.GetAuthenticationTriggerDisplayName())
            };
            settingsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Setting", "Value" }), secRows));

            // Connection settings
            settingsDocument.Root.Add(new MdHeading("Connection settings", 2));
            settingsDocument.Root.Add(new MdParagraph(new MdTextSpan(NotInExportMessage)));

            // Authoring canvas
            settingsDocument.Root.Add(new MdHeading("Authoring canvas", 2));
            settingsDocument.Root.Add(new MdParagraph(new MdTextSpan(NotInExportMessage)));

            // Entities
            settingsDocument.Root.Add(new MdHeading("Entities", 2));
            settingsDocument.Root.Add(new MdParagraph(new MdTextSpan(NotInExportMessage)));

            // Skills
            settingsDocument.Root.Add(new MdHeading("Skills", 2));
            settingsDocument.Root.Add(new MdParagraph(new MdTextSpan(NotInExportMessage)));

            // Voice
            settingsDocument.Root.Add(new MdHeading("Voice", 2));
            settingsDocument.Root.Add(new MdParagraph(new MdTextSpan(NotInExportMessage)));

            // Languages
            settingsDocument.Root.Add(new MdHeading("Languages", 2));
            List<MdTableRow> langRows = new List<MdTableRow>
            {
                new MdTableRow("Primary Language", content.agent.GetLanguageDisplayName())
            };
            settingsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Setting", "Value" }), langRows));

            // Language understanding
            settingsDocument.Root.Add(new MdHeading("Language understanding", 2));
            List<MdTableRow> luRows = new List<MdTableRow>
            {
                new MdTableRow("Recognizer", content.agent.GetRecognizerDisplayName())
            };
            settingsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Setting", "Value" }), luRows));

            // Component collections
            settingsDocument.Root.Add(new MdHeading("Component collections", 2));
            settingsDocument.Root.Add(new MdParagraph(new MdTextSpan(NotInExportMessage)));

            // Advanced
            settingsDocument.Root.Add(new MdHeading("Advanced", 2));
            List<MdTableRow> advRows = new List<MdTableRow>
            {
                new MdTableRow("Template", content.agent.Template ?? ""),
                new MdTableRow("Runtime Provider", content.agent.RuntimeProvider.ToString())
            };
            settingsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Setting", "Value" }), advRows));
        }

        private void addAgentOverview()
        {
            List<MdTableRow> tableRows = new List<MdTableRow>
            {
                new MdTableRow("Agent Name", content.agent.Name)
            };
            //todo where is the agent logo stored?
            if (!String.IsNullOrEmpty(content.agent.IconBase64))
            {
                Directory.CreateDirectory(content.folderPath);
                Bitmap agentLogo = ImageHelper.ConvertBase64ToBitmap(content.agent.IconBase64);

                agentLogo.Save(content.folderPath + $"agentlogo-{content.filename.Replace(" ", "-")}.png");

                tableRows.Add(new MdTableRow("Agent Logo", new MdImageSpan("Agent Logo", $"agentlogo-{content.filename.Replace(" ", "-")}.png")));
                agentLogo.Dispose();
            }

            tableRows.Add(new MdTableRow(content.headerDocumentationGenerated, PowerDocuReleaseHelper.GetTimestampWithVersion()));
            metadataTable = new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), tableRows);
            // prepare the common sections for top-level documents
            foreach (MdDocument doc in new[] { mainDocument, knowledgeDocument, toolsDocument, agentsDocument, topicsDocument, channelsDocument, settingsDocument })
            {
                doc.Root.Add(new MdHeading($"Agent - {content.filename}", 1));
                doc.Root.Add(metadataTable);
                doc.Root.Add(getNavigationLinks());
            }
            // prepare the common sections for topic documents (in Topics subfolder)
            foreach (var kvp in topicsDocuments)
            {
                kvp.Value.Root.Add(new MdHeading($"Agent - {content.filename}", 1));
                kvp.Value.Root.Add(metadataTable);
                kvp.Value.Root.Add(getNavigationLinks(false));
            }

            mainDocument.Root.Add(new MdHeading(content.Details, 2));
            mainDocument.Root.Add(new MdHeading(content.Description, 3));
            AddParagraphsWithLinebreaks(mainDocument, content.agent.GetDescription());
            mainDocument.Root.Add(new MdHeading(content.Orchestration, 3));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"{content.OrchestrationText} - {content.agent.GetOrchestration()}")));
            mainDocument.Root.Add(new MdHeading(content.ResponseModel, 3));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"{content.agent.GetResponseModel()}")));
            mainDocument.Root.Add(new MdHeading(content.Instructions, 3));
            AddParagraphsWithLinebreaks(mainDocument, content.agent.GetInstructions());
            mainDocument.Root.Add(new MdHeading(content.Knowledge, 3));
            foreach (BotComponent knowledgeSource in content.agent.GetKnowledge())
            {
                mainDocument.Root.Add(new MdParagraph(new MdTextSpan(knowledgeSource.Name)));
            }
            mainDocument.Root.Add(new MdHeading(content.WebSearch, 3));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan("TODO")));
            mainDocument.Root.Add(new MdHeading(content.Triggers, 3));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan("TODO")));
            mainDocument.Root.Add(new MdHeading(content.Agents, 3));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan("TODO")));
            mainDocument.Root.Add(new MdHeading(content.Topics, 3));
            List<MdListItem> topicsList = new List<MdListItem>();
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name))
            {
                topicsList.Add(new MdListItem(new MdLinkSpan(topic.Name, "Topics/" + ("topic " + topic.getTopicFileName() + " " + content.filename + ".md").Replace(" ", "-"))));
            }
            mainDocument.Root.Add(new MdBulletList(topicsList));
            mainDocument.Root.Add(new MdHeading(content.SuggestedPrompts, 3));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan(content.SuggestedPromptsText)));
            tableRows = new List<MdTableRow>();
            Dictionary<string, string> conversationStarters = content.agent.GetSuggestedPrompts();
            foreach (var kvp in conversationStarters.OrderBy(x => x.Key))
            {
                tableRows.Add(new MdTableRow(kvp.Key, kvp.Value));
            }
            mainDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Prompt Title", "Prompt" }), tableRows));
        }

        private MdBulletList getNavigationLinks(bool topLevel = true)
        {
            MdListItem[] navItems = new MdListItem[] {
                new MdListItem(new MdLinkSpan("Overview", topLevel ? mainDocumentFileName : "../" + mainDocumentFileName)),
                new MdListItem(new MdLinkSpan("Knowledge", topLevel ? knowledgeFileName : "../" + knowledgeFileName)),
                new MdListItem(new MdLinkSpan("Tools", topLevel ? toolsFileName : "../" + toolsFileName)),
                new MdListItem(new MdLinkSpan("Agents", topLevel ? agentsFileName : "../" + agentsFileName)),
                new MdListItem(new MdLinkSpan("Topics", topLevel ? topicsFileName : "../" + topicsFileName)),
                new MdListItem(new MdLinkSpan("Channels", topLevel ? channelsFileName : "../" + channelsFileName)),
                new MdListItem(new MdLinkSpan("Settings", topLevel ? settingsFileName : "../" + settingsFileName))
                };
            return new MdBulletList(navItems);
        }


        /*
            private void addAppDetails()
            {
                List<MdTableRow> tableRows = new List<MdTableRow>();
                knowledgeDocument.Root.Add(new MdHeading(content.appProperties.headerAppProperties, 2));
                foreach (Expression property in content.appProperties.appProperties)
                {
                    if (!content.appProperties.propertiesToSkip.Contains(property.expressionOperator))
                    {
                        tableRows.Add(new MdTableRow(property.expressionOperator, property.expressionOperands[0].ToString()));
                    }
                }
                if (tableRows.Count > 0)
                {
                    knowledgeDocument.Root.Add(new MdTable(new MdTableRow("App Property", "Value"), tableRows));
                }
                knowledgeDocument.Root.Add(new MdHeading(content.appProperties.headerAppPreviewFlags, 2));
                tableRows = new List<MdTableRow>();
                if (content.appProperties.appPreviewsFlagProperty != null)
                {
                    foreach (Expression flagProp in content.appProperties.appPreviewsFlagProperty.expressionOperands)
                    {
                        tableRows.Add(new MdTableRow(flagProp.expressionOperator, flagProp.expressionOperands[0].ToString()));
                    }
                    if (tableRows.Count > 0)
                    {
                        knowledgeDocument.Root.Add(new MdTable(new MdTableRow("Preview Flag", "Value"), tableRows));
                    }
                }
            }

    */
        private void addAgentKnowledgeInfo()
        {
            knowledgeDocument.Root.Add(new MdHeading(content.Knowledge, 2));
            knowledgeDocument.Root.Add(new MdParagraph(new MdTextSpan("Knowledge sources for this agent.")));
            foreach (BotComponent knowledgeSource in content.agent.GetKnowledge())
            {
                knowledgeDocument.Root.Add(new MdHeading(knowledgeSource.Name, 3));
            }
        }

        private void addAgentTopics()
        {
            topicsDocument.Root.Add(new MdHeading(content.Topics, 2));
            List<MdTableRow> tableRows = new List<MdTableRow>();
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name).ToList())
            {
                string topicType = topic.GetComponentTypeDisplayName();
                string triggerType = topic.GetTriggerTypeForTopic();
                tableRows.Add(new MdTableRow(
                    new MdLinkSpan(topic.Name, "Topics/" + ("topic " + topic.getTopicFileName() + " " + content.filename + ".md").Replace(" ", "-")),
                    topicType,
                    triggerType,
                    topic.GetTopicKind() == "KnowledgeSourceConfiguration" ? "Knowledge" : topic.GetTopicKind()));

                // Fill per-topic document
                topicsDocuments.TryGetValue(topic.getTopicFileName(), out MdDocument topicDoc);
                if (topicDoc != null)
                {
                    addTopicDetails(topicDoc, topic);
                }
            }
            topicsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Name", "Type", "Trigger", "Kind" }), tableRows));
        }

        private void addTopicDetails(MdDocument topicDoc, BotComponent topic)
        {
            topicDoc.Root.Add(new MdHeading("Topic: " + topic.Name, 2));

            // Metadata table
            List<MdTableRow> metaRows = new List<MdTableRow>
            {
                new MdTableRow("Name", topic.Name),
                new MdTableRow("Type", topic.GetComponentTypeDisplayName()),
                new MdTableRow("Trigger", topic.GetTriggerTypeForTopic()),
                new MdTableRow("Topic Kind", topic.GetTopicKind())
            };
            if (!string.IsNullOrEmpty(topic.Description))
            {
                metaRows.Add(new MdTableRow("Description", topic.Description));
            }
            string modelDesc = topic.GetModelDescription();
            if (!string.IsNullOrEmpty(modelDesc))
            {
                metaRows.Add(new MdTableRow("Model Description", modelDesc));
            }
            string startBehavior = topic.GetStartBehavior();
            if (!string.IsNullOrEmpty(startBehavior))
            {
                metaRows.Add(new MdTableRow("Start Behavior", startBehavior));
            }
            topicDoc.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), metaRows));

            // Trigger queries
            List<string> triggerQueries = topic.GetTriggerQueries();
            if (triggerQueries.Count > 0)
            {
                topicDoc.Root.Add(new MdHeading("Trigger Queries", 3));
                List<MdListItem> queryItems = new List<MdListItem>();
                foreach (string query in triggerQueries)
                {
                    queryItems.Add(new MdListItem(query));
                }
                topicDoc.Root.Add(new MdBulletList(queryItems));
            }

            // Knowledge source details for KnowledgeSourceConfiguration topics
            if (topic.GetTopicKind() == "KnowledgeSourceConfiguration")
            {
                var (sourceKind, skillConfig) = topic.GetKnowledgeSourceDetails();
                topicDoc.Root.Add(new MdHeading("Knowledge Source", 3));
                List<MdTableRow> ksRows = new List<MdTableRow>();
                if (!string.IsNullOrEmpty(sourceKind))
                    ksRows.Add(new MdTableRow("Source Kind", sourceKind));
                if (!string.IsNullOrEmpty(skillConfig))
                    ksRows.Add(new MdTableRow("Skill Configuration", skillConfig));
                if (ksRows.Count > 0)
                    topicDoc.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), ksRows));
            }

            // Variables
            var variables = topic.GetTopicVariables();
            if (variables.Count > 0)
            {
                topicDoc.Root.Add(new MdHeading("Variables", 3));
                List<MdTableRow> varRows = new List<MdTableRow>();
                foreach (var (variable, context) in variables)
                {
                    varRows.Add(new MdTableRow(variable, context));
                }
                topicDoc.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Variable", "Context" }), varRows));
            }

            // Topic flow diagram
            string graphFile = topic.getTopicFileName() + "-detailed.svg";
            if (File.Exists(Path.Combine(content.folderPath, "Topics", graphFile)))
            {
                topicDoc.Root.Add(new MdHeading("Topic Flow", 3));
                topicDoc.Root.Add(new MdParagraph(new MdImageSpan("Topic Flow Diagram", graphFile)));
            }
        }

        private MdBulletList CreateControlList(ControlEntity control)
        {
            var svgDocument = SvgDocument.FromSvg<SvgDocument>(AppControlIcons.GetControlIcon(control.Type));
            //generating the PNG from the SVG with a width of 16px because some SVGs are huge and downscaled, thus can't be shown directly
            using (var bitmap = svgDocument.Draw(16, 0))
            {
                bitmap?.Save(content.folderPath + @"resources\" + control.Type + ".png");
            }
            //link to the screen instead of the control directly for the moment, as the directly generated anchor link (#" + control.Name.ToLower()) doesn't work the same way in DevOps and GitHub
            MdBulletList list = new MdBulletList(){
                                     new MdListItem(new MdLinkSpan(
                                            new MdCompositeSpan(
                                                new MdImageSpan(control.Type, "resources/"+control.Type+".png"),
                                                new MdTextSpan(" "+control.Name))
                                        ,("screen " + CharsetHelper.GetSafeName(control.Screen().Name) + " " + content.filename + ".md").Replace(" ", "-")))};

            foreach (ControlEntity child in control.Children.OrderBy(o => o.Name).ToList())
            {
                list.Add(new MdListItem(CreateControlList(child)));
            }
            return list;
        }

        private void addAgentChannels()
        {
            channelsDocument.Root.Add(new MdHeading("Channels", 2));
            channelsDocument.Root.Add(new MdParagraph(new MdTextSpan("Channels are not exported with the solution and are not documented automatically.")));
        }
        /*
                private void addAppControlsTable(ControlEntity control, MdDocument screenDoc)
                {
                    Entity defaultEntity = DefaultChangeHelper.GetEntityDefaults(control.Type);
                    List<MdTableRow> tableRows = new List<MdTableRow>();
                    var svgDocument = SvgDocument.FromSvg<SvgDocument>(AppControlIcons.GetControlIcon(control.Type));
                    //generating the PNG from the SVG with a width of 16px because some SVGs are huge and downscaled, thus can't be shown directly
                    using (var bitmap = svgDocument.Draw(16, 0))
                    {
                        bitmap?.Save(content.folderPath + @"resources\" + control.Type + ".png");
                    }
                    tableRows.Add(new MdTableRow(new MdImageSpan(control.Type, "resources/" + control.Type + ".png"), new MdTextSpan("Type: " + control.Type)));

                    string category = "";
                    foreach (Rule rule in control.Rules.OrderBy(o => o.Category).ThenBy(o => o.Property).ToList())
                    {
                        string defaultValue = defaultEntity?.Rules.Find(r => r.Property == rule.Property)?.InvariantScript;
                        if (String.IsNullOrEmpty(defaultValue))
                            defaultValue = DefaultChangeHelper.DefaultValueIfUnknown;
                        if (!documentChangedDefaultsOnly || (defaultValue != rule.InvariantScript))
                        {
                            if (!content.ColourProperties.Contains(rule.Property))
                            {
                                if (rule.Category != category)
                                {
                                    if (tableRows.Count > 0)
                                    {
                                        screenDoc.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
                                        tableRows = new List<MdTableRow>();
                                    }
                                    category = rule.Category;
                                    screenDoc.Root.Add(new MdHeading(category, 3));
                                }
                                if (rule.InvariantScript.StartsWith("RGBA("))
                                {
                                    tableRows.Add(CreateColorTable(rule, defaultValue));
                                }
                                else
                                {
                                    tableRows.Add(CreateRowForControlProperty(rule, defaultValue));
                                }
                            }
                        }
                    }
                    if (tableRows.Count > 0)
                        screenDoc.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
                    //Colour properties
                    tableRows = new List<MdTableRow>();
                    screenDoc.Root.Add(new MdHeading("Color Properties", 3));
                    foreach (string property in content.ColourProperties)
                    {
                        Rule rule = control.Rules.Find(o => o.Property == property);
                        if (rule != null)
                        {
                            string defaultValue = defaultEntity?.Rules.Find(r => r.Property == rule.Property)?.InvariantScript;
                            if (String.IsNullOrEmpty(defaultValue))
                                defaultValue = DefaultChangeHelper.DefaultValueIfUnknown;
                            if (!documentChangedDefaultsOnly || defaultValue != rule.InvariantScript)
                            {
                                if (rule.InvariantScript.StartsWith("RGBA("))
                                {
                                    tableRows.Add(CreateColorTable(rule, defaultValue));
                                }
                                else
                                {
                                    tableRows.Add(new MdTableRow(rule.Property, $"`{rule.InvariantScript}`"));
                                }
                            }
                        }
                    }
                    if (tableRows.Count > 0)
                        screenDoc.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
                    tableRows = new List<MdTableRow>();
                    screenDoc.Root.Add(new MdHeading("Child & Parent Controls", 3));

                    foreach (ControlEntity childControl in control.Children)
                    {
                        tableRows.Add(new MdTableRow("Child Control", childControl.Name));
                    }
                    if (control.Parent != null)
                    {
                        tableRows.Add(new MdTableRow("Parent Control", control.Parent.Name));
                    }
                    if (tableRows.Count > 0)
                        screenDoc.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
                }


                private MdTableRow CreateColorTable(Rule rule, string defaultValue)
                {
                    StringBuilder colourTable = new StringBuilder("<table border=\"0\">");
                    colourTable.Append("<tr><td>").Append(rule.InvariantScript).Append("</td></tr>");
                    string colour = ColourHelper.ParseColor(rule.InvariantScript[..(rule.InvariantScript.IndexOf(')') + 1)]);
                    if (!String.IsNullOrEmpty(colour))
                    {
                        colourTable.Append("<tr><td style=\"background-color:").Append(colour).Append("\"></td></tr>");
                    }
                    colourTable.Append("</table>");
                    if (showDefaults && defaultValue != rule.InvariantScript && !content.appControls.controlPropertiesToSkip.Contains(rule.Property))
                    {
                        StringBuilder defaultTable = new StringBuilder("<table border=\"0\">");
                        defaultTable.Append("<tr><td>").Append(defaultValue).Append("</td></tr>");
                        string defaultColour = ColourHelper.ParseColor(defaultValue);
                        if (!String.IsNullOrEmpty(defaultColour))
                        {
                            defaultTable.Append("<tr><td style=\"background-color:").Append(defaultColour).Append("\"></td></tr>");
                        }
                        defaultTable.Append("</table>");
                        StringBuilder changesTable = new StringBuilder("<table border=\"0\">");
                        changesTable.Append(CreateChangedDefaultColourRow(colourTable.ToString(), defaultTable.ToString()));
                        return new MdTableRow(rule.Property, new MdRawMarkdownSpan(changesTable.Append("</table>").ToString()));
                    }
                    return new MdTableRow(rule.Property, new MdRawMarkdownSpan(colourTable.ToString()));
                }

                private MdTableRow CreateRowForControlProperty(Rule rule, string defaultValue)
                {
                    if (showDefaults && defaultValue != rule.InvariantScript && !content.appControls.controlPropertiesToSkip.Contains(rule.Property))
                    {
                        StringBuilder table = new StringBuilder("<table border=\"0\">");
                        table.Append("<tr><td style=\"background-color:#ccffcc; width:50%;\">")
                             .Append($"`{rule.InvariantScript}`")
                             .Append("<td style=\"background-color:#ffcccc; width:50%;\">").Append(defaultValue).Append("</td></tr></table>");
                        return new MdTableRow(rule.Property, new MdRawMarkdownSpan(table.ToString()));
                    }
                    return new MdTableRow(rule.Property, $"`{rule.InvariantScript}`");
                }
        */
        private void addAgentTools()
        {
            toolsDocument.Root.Add(new MdHeading(content.Tools, 2));
            toolsDocument.Root.Add(new MdParagraph(new MdTextSpan("Tools available for this agent.")));
        }

        private void addAgentAgentsInfo()
        {
            agentsDocument.Root.Add(new MdHeading(content.Agents, 2));
            agentsDocument.Root.Add(new MdParagraph(new MdTextSpan("Sub-agents for this agent.")));
        }

        private void AddParagraphsWithLinebreaks(MdDocument document, string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                document.Root.Add(new MdParagraph(new MdTextSpan("")));
                return;
            }
            string[] lines = text.Replace("\r\n", "\n").Split('\n');
            foreach (string line in lines)
            {
                document.Root.Add(new MdParagraph(new MdTextSpan(line)));
            }
        }
    }
}
