using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using PowerDocu.Common;
using Rubjerg.Graphviz;
using Svg;
using YamlDotNet.RepresentationModel;

namespace PowerDocu.AgentDocumenter
{
    public class GraphBuilder
    {
        private readonly BotComponent topic;
        private readonly string folderPath;
        private YamlScalarNode actionsKey = new YamlScalarNode("actions");
        private YamlScalarNode conditions = new YamlScalarNode("conditions");
        private YamlScalarNode elseActions = new YamlScalarNode("elseActions");

        public GraphBuilder(string agentName, BotComponent topicToUse, string path)
        {
            topic = topicToUse;
            folderPath = path;
            Directory.CreateDirectory(folderPath + "Resources");
        }

        public void buildTopLevelGraph()
        {
            buildGraph(false);
        }

        public void buildDetailedGraph()
        {
            buildGraph(true);
        }

        private void buildGraph(bool showSubactions)
        {
            RootGraph rootGraph = RootGraph.CreateNew(GraphType.Directed, CharsetHelper.GetSafeName(topic.Name + "(" + topic.getTopicFileName() + ")"));
            Graph.IntroduceAttribute(rootGraph, "rankdir", "TB");
            Graph.IntroduceAttribute(rootGraph, "compound", "true");
            Graph.IntroduceAttribute(rootGraph, "fontname", "helvetica");

            // Add subgraph default attributes
            Graph.IntroduceAttribute(rootGraph, "clusterrank", "local");
            SubGraph.IntroduceAttribute(rootGraph, "style", "filled");
            SubGraph.IntroduceAttribute(rootGraph, "color", "black");
            SubGraph.IntroduceAttribute(rootGraph, "fillcolor", "lightgray");
            SubGraph.IntroduceAttribute(rootGraph, "penwidth", "1");

            Node.IntroduceAttribute(rootGraph, "shape", "rectangle");
            Node.IntroduceAttribute(rootGraph, "color", "");
            Node.IntroduceAttribute(rootGraph, "style", "");
            Node.IntroduceAttribute(rootGraph, "fillcolor", "");
            Node.IntroduceAttribute(rootGraph, "label", "");
            Node.IntroduceAttribute(rootGraph, "fontname", "helvetica");
            Edge.IntroduceAttribute(rootGraph, "label", "");
            var yaml = new YamlStream();
            //pasrse the topic YAML data
            //this is ugly and should't be done, but otherwise YamlDotNet throws an error.... need longterm strategy
            var input = new StringReader(topic.YamlData.Replace("@odata", "odata"));
            yaml.Load(input);
            var mapping = (YamlMappingNode)yaml.Documents[0].RootNode;

            // Guard: skip graph generation for non-AdaptiveDialog topics (e.g. KnowledgeSourceConfiguration)
            if (!mapping.Children.TryGetValue(new YamlScalarNode("beginDialog"), out var beginDialogNode)
                || !(beginDialogNode is YamlMappingNode triggerYamlNode))
            {
                // Create a simple single-node graph for non-dialog topics
                string topicKind = "Unknown";
                if (mapping.Children.TryGetValue(new YamlScalarNode("kind"), out var kindNode))
                    topicKind = kindNode.ToString();
                Node infoNode = rootGraph.GetOrAddNode("info_" + topic.Name);
                var svgDoc = SvgDocument.FromSvg<SvgDocument>(AgentIcon.GetIcon("KnowledgeSource"));
                using (var bmp = svgDoc.Draw(20, 0))
                {
                    bmp?.Save(folderPath + @"Resources\KnowledgeSource.png");
                }
                infoNode.SetAttribute("shape", "plain");
                string infoHtml = createCardStart("Trigger", topicKind + ": " + topic.Name, "KnowledgeSource");
                infoHtml += createCardEnd();
                infoNode.SetAttributeHtml("label", infoHtml);
                rootGraph.CreateLayout();
                NotificationHelper.SendNotification("  - Created Graph " + folderPath + generateImageFiles(rootGraph, showSubactions) + ".png");
                return;
            }
            var triggerYaml = triggerYamlNode;
            Node trigger = addTriggerDetails(triggerYaml, rootGraph);
            addActionNodes(triggerYaml, trigger, rootGraph, actionsKey);
            rootGraph.CreateLayout();
            NotificationHelper.SendNotification("  - Created Graph " + folderPath + generateImageFiles(rootGraph, showSubactions) + ".png");
        }

        private Node addTriggerDetails(YamlMappingNode triggerYaml, RootGraph rootGraph)
        {
            string triggerType = triggerYaml.Children[new YamlScalarNode("kind")].ToString();
            Node trigger = rootGraph.GetOrAddNode("Trigger " + topic.Name + " (" + topic.getTopicFileName() + ")");
            trigger.SetAttribute("shape", "plain");
            var svgDocument = SvgDocument.FromSvg<SvgDocument>(AgentIcon.GetIcon("Trigger"));
            //generating the PNG from the SVG with a width of 20px because some SVGs are huge and downscaled, thus can't be shown directly
            using (var bitmap = svgDocument.Draw(20, 0))
            {
                bitmap?.Save(folderPath + @"Resources\Trigger.png");
            }
            string html = createCardStart("Trigger", $"Trigger - {topic.Name}");
            switch (triggerType)
            {
                case "OnRecognizedIntent":

                    html += createCardBodyRow("<table border=\"1\" cellpadding=\"4\"><tr><td>The agent chooses</td></tr></table>");
                    html += createCardBodyRow("Describe what the topic does:");
                    html += "<tr><td cellpadding=\"8\"><table border=\"1\" cellpadding=\"4\"><tr><td>This tool can handle queries like these:<br/>";
                    var intentYaml = (YamlMappingNode)triggerYaml.Children[new YamlScalarNode("intent")];
                    if (intentYaml.Children.TryGetValue(new YamlScalarNode("triggerQueries"), out var triggerQueryNode) && triggerQueryNode is YamlSequenceNode triggerQuerySequence)
                    {
                        html += string.Join("<br/>", triggerQuerySequence);
                    }
                    html += "</td></tr></table></td></tr>";

                    break;
                case "OnConversationStart":
                    html += createCardBodyRow("<table border=\"1\" cellpadding=\"4\"><tr><td>When a conversation starts</td></tr></table>");
                    break;
                case "OnEscalate":
                    html += createCardBodyRow("<table border=\"1\" cellpadding=\"4\"><tr><td>Talk to a representative</td></tr></table>");
                    if (triggerYaml.Children.TryGetValue(new YamlScalarNode("intent"), out var escIntentNode)
                        && escIntentNode is YamlMappingNode escIntentMapping
                        && escIntentMapping.Children.TryGetValue(new YamlScalarNode("triggerQueries"), out var escTriggerNode)
                        && escTriggerNode is YamlSequenceNode escTriggerSeq)
                    {
                        html += "<tr><td cellpadding=\"8\"><table border=\"1\" cellpadding=\"4\"><tr><td>";
                        html += string.Join("<br/>", escTriggerSeq.Take(10));
                        if (escTriggerSeq.Count() > 10)
                            html += "<br/>... and more";
                        html += "</td></tr></table></td></tr>";
                    }
                    break;
                case "OnUnknownIntent":
                    html += createCardBodyRow("<table border=\"1\" cellpadding=\"4\"><tr><td>Unknown / unmatched intent (Fallback)</td></tr></table>");
                    break;
                case "OnRedirect":
                    html += createCardBodyRow("<table border=\"1\" cellpadding=\"4\"><tr><td>Redirected from another topic</td></tr></table>");
                    break;
                case "OnSystemRedirect":
                    html += createCardBodyRow("<table border=\"1\" cellpadding=\"4\"><tr><td>System redirect</td></tr></table>");
                    break;
                case "OnError":
                    html += createCardBodyRow("<table border=\"1\" cellpadding=\"4\"><tr><td>When an error occurs</td></tr></table>");
                    break;
                case "OnSignIn":
                    html += createCardBodyRow("<table border=\"1\" cellpadding=\"4\"><tr><td>When sign-in is required</td></tr></table>");
                    break;
                case "OnSelectIntent":
                    html += createCardBodyRow("<table border=\"1\" cellpadding=\"4\"><tr><td>When multiple topics matched</td></tr></table>");
                    break;
                default:
                    html += createCardBodyRow($"<table border=\"1\" cellpadding=\"4\"><tr><td>{System.Web.HttpUtility.HtmlEncode(triggerType)}</td></tr></table>");
                    break;
            }
            html += createCardEnd();
            trigger.SetAttributeHtml("label", html);
            return trigger;
        }

        private Node addActionNodes(YamlMappingNode actionsYaml, Node prevNode, RootGraph rootGraph, YamlScalarNode filterKey, SubGraph parentCluster = null)
        {
            Node returnNode = null;
            if (actionsYaml.Children.TryGetValue(filterKey, out var actionsNode) && actionsNode is YamlSequenceNode actionsSequence)
            {
                foreach (var action in actionsSequence)
                {
                    Node clusterExitNode = null;
                    Node actionNode = null;
                    SubGraph conditionCluster = null;
                    string actionName = CharsetHelper.GetSafeName(((YamlMappingNode)action).Children[new YamlScalarNode("kind")].ToString() + ((YamlMappingNode)action).Children[new YamlScalarNode("id")].ToString());
                    string actionType = GetActionType((YamlMappingNode)action);
                    if (actionType != "ConditionGroup")
                    {
                        actionNode = rootGraph.GetOrAddNode(actionName);
                        actionNode.SetAttribute("shape", "plain");
                        parentCluster?.AddExisting(actionNode);
                        returnNode = actionNode;
                    }
                    string displayName = "";
                    if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("displayName"), out var displayNameNode))
                    {
                        displayName = System.Web.HttpUtility.HtmlEncode(displayNameNode.ToString());
                    }
                    var svgDocument = SvgDocument.FromSvg<SvgDocument>(AgentIcon.GetIcon(actionType));
                    //generating the PNG from the SVG with a width of 16px because some SVGs are huge and downscaled, thus can't be shown directly
                    using (var bitmap = svgDocument.Draw(20, 0))
                    {
                        bitmap?.Save(folderPath + @"Resources\" + actionType + ".png");
                    }
                    switch (actionType)
                    {
                        case "AdaptiveCard":

                            string adapativeCardHtml = createCardStart(actionType, $"Adaptive Card: {displayName}");

                            // Parse the card JSON content
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("card"), out var cardNode))
                            {
                                try
                                {
                                    string cardJsonString = cardNode.ToString();
                                    if (cardJsonString.StartsWith('='))
                                    {
                                        cardJsonString = cardJsonString.Substring(1);
                                    }
                                    JObject cardJson = JObject.Parse(cardJsonString);

                                    // Collect rows first to avoid empty table (Graphviz requires at least one <TR>)
                                    var cardRows = new List<string>();

                                    // Process card body elements
                                    if (cardJson["body"] is JArray bodyArray)
                                    {
                                        foreach (var element in bodyArray)
                                        {
                                            string elementHtml = RenderCardElement(element as JObject);
                                            if (!string.IsNullOrEmpty(elementHtml))
                                            {
                                                cardRows.Add($"<tr><td>{elementHtml}</td></tr>");
                                            }
                                        }
                                    }
                                    // Process card actions
                                    if (cardJson["actions"] is JArray actionsArray)
                                    {
                                        bool hasActions = false;
                                        foreach (var cardAction in actionsArray)
                                        {
                                            string actionHtml = RenderCardAction(cardAction as JObject);
                                            if (!string.IsNullOrEmpty(actionHtml))
                                            {
                                                if (!hasActions) { cardRows.Add("<tr><td> </td></tr>"); hasActions = true; }
                                                cardRows.Add($"<tr><td>{actionHtml}</td></tr>");
                                            }
                                        }
                                    }

                                    // Only emit the inner table if there are rows to show
                                    if (cardRows.Count > 0)
                                    {
                                        adapativeCardHtml += "<tr><td cellpadding=\"8\"><table border=\"1\" cellpadding=\"4\">";
                                        adapativeCardHtml += string.Join("", cardRows);
                                        adapativeCardHtml += "</table></td></tr>";
                                    }
                                    else
                                    {
                                        adapativeCardHtml += createCardBodyRow("(empty card)");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    adapativeCardHtml += createCardBodyRow("Error parsing adaptive card: " + System.Web.HttpUtility.HtmlEncode(ex.Message));
                                }
                            }
                            else
                            {
                                adapativeCardHtml += createCardBodyRow("Adaptive Card definition not found");
                            }

                            // Process output binding if present
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("output"), out var outputNode) &&
                                outputNode is YamlMappingNode outputMapping &&
                                outputMapping.Children.TryGetValue(new YamlScalarNode("binding"), out var bindingNode) &&
                                bindingNode is YamlMappingNode outputBinding)
                            {
                                adapativeCardHtml += "<tr><td cellpadding=\"8\"><table border=\"1\" cellpadding=\"4\">";
                                adapativeCardHtml += "<tr><td><b>Output Binding:</b></td></tr>";
                                foreach (var outputItem in outputBinding.Children)
                                {
                                    adapativeCardHtml += $"<tr><td><b>{outputItem.Key}:</b> {formatVariable(outputItem.Value.ToString())}</td></tr>";
                                }
                                adapativeCardHtml += "</table></td></tr>";
                            }

                            adapativeCardHtml += createCardEnd();
                            actionNode.SetAttributeHtml("label", adapativeCardHtml);
                            break;
                        case "Question":
                            string questionHtml = createCardStart(actionType, $"Question: {displayName}");

                            // Get the prompt text
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("prompt"), out var promptNode))
                            {
                                string promptText = promptNode.ToString();
                                questionHtml += "<tr><td cellpadding=\"8\"><table border=\"1\" cellpadding=\"4\">";
                                questionHtml += $"<tr><td>{generateMultiLineText(System.Web.HttpUtility.HtmlEncode(promptText))}</td></tr>";
                                questionHtml += "</table></td></tr>";
                            }


                            // Get the entity type
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("entity"), out var entityNode))
                            {
                                string entityType = entityNode.ToString();
                                questionHtml += createCardBodyRow($"<b>Identify:</b> {entityType}");
                            }

                            // Get the variable name
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("variable"), out var variableNode))
                            {
                                string variableName = variableNode.ToString();
                                questionHtml += createCardBodyRow($"<b>Save user response as:</b><br/> {formatVariable(variableName)}");
                            }

                            // Check for choices if it's a choice question
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("choices"), out var choicesNode) && choicesNode is YamlSequenceNode choicesSequence)
                            {
                                questionHtml += createCardBodyRow("<b>Choices:</b>");
                                questionHtml += "<tr><td cellpadding=\"8\"><table border=\"1\" cellpadding=\"4\">";
                                foreach (var choice in choicesSequence)
                                {
                                    if (choice is YamlMappingNode choiceMap)
                                    {
                                        string choiceText = "";
                                        if (choiceMap.Children.TryGetValue(new YamlScalarNode("value"), out var valueNode))
                                        {
                                            choiceText = valueNode.ToString();
                                        }
                                        if (choiceMap.Children.TryGetValue(new YamlScalarNode("synonyms"), out var synonymsNode) && synonymsNode is YamlSequenceNode synonymsSeq)
                                        {
                                            var synonymsList = synonymsSeq.Select(s => s.ToString()).ToList();
                                            choiceText += $" (synonyms: {string.Join(", ", synonymsList)})";
                                        }
                                        questionHtml += $"<tr><td>{System.Web.HttpUtility.HtmlEncode(choiceText)}</td></tr>";
                                    }
                                    else
                                    {
                                        // Simple string choice
                                        questionHtml += $"<tr><td>{System.Web.HttpUtility.HtmlEncode(choice.ToString())}</td></tr>";
                                    }
                                }
                                questionHtml += "</table></td></tr>";
                            }

                            questionHtml += createCardEnd();
                            actionNode.SetAttributeHtml("label", questionHtml);
                            break;
                        case "ConditionGroup":
                            // Create the condition cluster - use the parentCluster if we're nested
                            conditionCluster = ((Graph)(parentCluster != null ? parentCluster : rootGraph)).GetOrAddSubgraph("cluster_" + CharsetHelper.GetSafeName(actionName));

                            // Use SetAttribute instead of SafeSetAttribute for better reliability
                            conditionCluster.SetAttribute("style", "filled");
                            conditionCluster.SetAttribute("fillcolor", GraphColours.GetFillColourForAction(actionType));
                            conditionCluster.SetAttribute("color", GraphColours.GetColourForAction(actionType));
                            conditionCluster.SetAttribute("penwidth", "1");

                            // Track the last nodes in each condition branch
                            var lastNodes = new List<Node>();
                            // Collect all nodes that should be at the same level
                            var sameLevelNodes = new List<Node>();
                            //get the conditions node, then loop through the items inside which may have actions
                            if (((YamlMappingNode)action).Children.TryGetValue(conditions, out var conditionsNode) && conditionsNode is YamlSequenceNode conditionsSequence)
                            {
                                foreach (var condition in conditionsSequence)
                                {
                                    //add the condition node
                                    Node conditionNode = rootGraph.GetOrAddNode("conditionnode-" + CharsetHelper.GetSafeName(((YamlMappingNode)condition).Children[new YamlScalarNode("id")].ToString()));
                                    conditionNode.SetAttribute("shape", "plain");
                                    string rawCondition = ((YamlMappingNode)condition).Children[new YamlScalarNode("condition")].ToString();
                                    string conditionHtml = createCardStart(actionType, "Condition");
                                    conditionHtml += createCardBodyRow(formatConditionExpression(rawCondition));
                                    conditionHtml += createCardEnd();
                                    conditionNode.SetAttributeHtml("label", conditionHtml);

                                    // Create edge from previous node to condition node
                                    Edge edge = rootGraph.GetOrAddEdge(prevNode, conditionNode, "edge to " + "conditionnode-" + CharsetHelper.GetSafeName(((YamlMappingNode)condition).Children[new YamlScalarNode("id")].ToString()));
                                    edge.SetAttribute("weight", "1");
                                    conditionCluster.AddExisting(conditionNode);

                                    // Process actions and get the last node in this branch - PASS conditionCluster as parentCluster for nesting
                                    Node lastNodeInBranch = addActionNodes((YamlMappingNode)condition, conditionNode, rootGraph, actionsKey, conditionCluster);
                                    lastNodes.Add(lastNodeInBranch ?? conditionNode);
                                    sameLevelNodes.Add(conditionNode);
                                }
                            }

                            // Handle else actions if present
                            if (((YamlMappingNode)action).Children.TryGetValue(elseActions, out var elseActionsYamlNode))
                            {
                                //add the else actions node
                                Node elseActionsNode = rootGraph.GetOrAddNode("elseactionsnode-" + actionName);
                                elseActionsNode.SetAttribute("shape", "plain");
                                string elseHtml = createCardStart(actionType, "Condition");
                                elseHtml += createCardBodyRow("All Other Conditions");
                                elseHtml += createCardEnd();
                                elseActionsNode.SetAttributeHtml("label", elseHtml);
                                Edge edge = rootGraph.GetOrAddEdge(prevNode, elseActionsNode, "edge to " + "elseactionsnode-" + actionName);
                                edge.SetAttribute("weight", "1");
                                conditionCluster.AddExisting(elseActionsNode);

                                // Process else actions - PASS conditionCluster as parentCluster for nesting
                                Node lastElseNode = addActionNodes((YamlMappingNode)action, elseActionsNode, rootGraph, elseActions, conditionCluster);
                                lastNodes.Add(lastElseNode ?? elseActionsNode);
                                sameLevelNodes.Add(elseActionsNode);
                            }

                            // Create exit node AFTER processing all conditions
                            clusterExitNode = rootGraph.GetOrAddNode(actionName + "_cluster_exit");
                            clusterExitNode.SetAttribute("margin", "0");
                            clusterExitNode.SetAttribute("style", "invis");
                            clusterExitNode.SetAttribute("width", "0");
                            clusterExitNode.SetAttribute("height", "0");
                            clusterExitNode.SetAttribute("shape", "point");
                            conditionCluster.AddExisting(clusterExitNode);

                            // Connect all last nodes to the exit node with invisible edges
                            foreach (var lastNode in lastNodes)
                            {
                                if (lastNode != null)
                                {
                                    Edge exitEdge = rootGraph.GetOrAddEdge(lastNode, clusterExitNode, "to_exit_" + lastNode.GetName());
                                    exitEdge.SetAttribute("style", "invis");
                                    exitEdge.SetAttribute("weight", "100");
                                    exitEdge.SetAttribute("minlen", "1");
                                }
                            }

                            returnNode = clusterExitNode;
                            prevNode = clusterExitNode;
                            break;
                        case "AIModel":
                            string aiModelHtml = createCardStart(actionType, !string.IsNullOrEmpty(displayName) ? displayName : "Prompt");

                            // Show AI Model ID if present
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("aIModelId"), out var aiModelIdNode))
                            {
                                aiModelHtml += createCardBodyRow($"<font point-size=\"9\"><b>AI Model ID:</b> {generateMultiLineText(System.Web.HttpUtility.HtmlEncode(aiModelIdNode.ToString()))}</font>");
                            }

                            // Show input bindings
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("input"), out var aiInputNode)
                                && aiInputNode is YamlMappingNode aiInputMapping
                                && aiInputMapping.Children.TryGetValue(new YamlScalarNode("binding"), out var aiBindingNode)
                                && aiBindingNode is YamlMappingNode aiBindingMapping)
                            {
                                aiModelHtml += "<tr><td cellpadding=\"8\"><table border=\"1\" cellpadding=\"4\"><tr><td><b>Inputs:</b></td></tr>";
                                foreach (var kvp in aiBindingMapping.Children)
                                {
                                    aiModelHtml += $"<tr><td><b>{System.Web.HttpUtility.HtmlEncode(kvp.Key.ToString())}:</b> {formatVariable(kvp.Value.ToString())}</td></tr>";
                                }
                                aiModelHtml += "</table></td></tr>";
                            }

                            // Show output bindings
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("output"), out var aiOutputNode)
                                && aiOutputNode is YamlMappingNode aiOutputMapping
                                && aiOutputMapping.Children.TryGetValue(new YamlScalarNode("binding"), out var aiOutputBindingNode)
                                && aiOutputBindingNode is YamlMappingNode aiOutputBindingMapping)
                            {
                                aiModelHtml += "<tr><td cellpadding=\"8\"><table border=\"1\" cellpadding=\"4\"><tr><td><b>Outputs:</b></td></tr>";
                                foreach (var kvp in aiOutputBindingMapping.Children)
                                {
                                    aiModelHtml += $"<tr><td><b>{System.Web.HttpUtility.HtmlEncode(kvp.Key.ToString())}:</b> {formatVariable(kvp.Value.ToString())}</td></tr>";
                                }
                                aiModelHtml += "</table></td></tr>";
                            }

                            aiModelHtml += createCardEnd();
                            actionNode.SetAttributeHtml("label", aiModelHtml);
                            break;
                        case "Message":
                            YamlScalarNode messageYaml = null;
                            var activityYaml = ((YamlMappingNode)action).Children[new YamlScalarNode("activity")];
                            if (activityYaml.GetType().Equals(typeof(YamlMappingNode)))
                            {
                                messageYaml = (YamlScalarNode)((YamlSequenceNode)((YamlMappingNode)activityYaml).Children[new YamlScalarNode("text")]).First();
                                //todo there may also be a speak node, or maybe even no text
                            }
                            else if (activityYaml.GetType().Equals(typeof(YamlScalarNode)))
                            {
                                messageYaml = (YamlScalarNode)activityYaml;
                            }
                            string msgHtml = createCardStart(actionType, "Message");
                            msgHtml += createCardBodyRow(CharsetHelper.GetSafeName(messageYaml.Value).Replace("\n", "<br/>"));
                            msgHtml += createCardEnd();
                            actionNode.SetAttributeHtml("label", msgHtml);
                            break;
                        case "SetVariable":
                            var variableYaml = (YamlScalarNode)((YamlMappingNode)action).Children[new YamlScalarNode("variable")];
                            var valueYaml = (YamlScalarNode)((YamlMappingNode)action).Children[new YamlScalarNode("value")];
                            string setVarHtml = createCardStart(actionType, "Set Variable");
                            setVarHtml += createCardBodyRow($"<table border=\"1\" cellpadding=\"4\"><tr><td>{formatVariable(variableYaml.Value)}</td></tr></table>");
                            setVarHtml += createCardBodyRow("To Value");
                            string variableValue = valueYaml.Value;
                            if (variableValue.StartsWith('='))
                            {
                                variableValue = variableValue.Substring(1);
                            }
                            setVarHtml += createCardBodyRow($"<table border=\"1\" cellpadding=\"4\"><tr><td>{generateMultiLineText(System.Web.HttpUtility.HtmlEncode(variableValue))}</td></tr></table>");
                            setVarHtml += createCardEnd();
                            actionNode.SetAttributeHtml("label", setVarHtml);
                            break;
                        case "CancelAllDialogs":
                            actionNode.SetAttributeHtml("label", createCardStart(actionType, !string.IsNullOrEmpty(displayName) ? displayName : "End all topics") + createCardEnd());
                            break;
                        case "LogCustomTelemetry":
                            actionNode.SetAttributeHtml("label", createCardStart(actionType, !string.IsNullOrEmpty(displayName) ? displayName : "Log custom telemetry event") + createCardEnd());
                            break;
                        case "InvokeFlow":
                            string flowHtml = createCardStart(actionType, !string.IsNullOrEmpty(displayName) ? $"Flow: {displayName}" : "Call a flow");
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("flowId"), out var flowIdNode))
                            {
                                flowHtml += createCardBodyRow($"<font point-size=\"9\"><b>Flow ID:</b> {generateMultiLineText(System.Web.HttpUtility.HtmlEncode(flowIdNode.ToString()))}</font>");
                            }
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("input"), out var flowInputNode)
                                && flowInputNode is YamlMappingNode flowInputMapping
                                && flowInputMapping.Children.TryGetValue(new YamlScalarNode("binding"), out var flowBindingNode)
                                && flowBindingNode is YamlMappingNode flowBindingMapping)
                            {
                                flowHtml += "<tr><td cellpadding=\"8\"><table border=\"1\" cellpadding=\"4\"><tr><td><b>Inputs:</b></td></tr>";
                                foreach (var kvp in flowBindingMapping.Children)
                                {
                                    flowHtml += $"<tr><td><b>{System.Web.HttpUtility.HtmlEncode(kvp.Key.ToString())}:</b> {formatVariable(kvp.Value.ToString())}</td></tr>";
                                }
                                flowHtml += "</table></td></tr>";
                            }
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("output"), out var flowOutputNode)
                                && flowOutputNode is YamlMappingNode flowOutputMapping
                                && flowOutputMapping.Children.TryGetValue(new YamlScalarNode("binding"), out var flowOutputBindingNode)
                                && flowOutputBindingNode is YamlMappingNode flowOutputBindingMapping)
                            {
                                flowHtml += "<tr><td cellpadding=\"8\"><table border=\"1\" cellpadding=\"4\"><tr><td><b>Outputs:</b></td></tr>";
                                foreach (var kvp in flowOutputBindingMapping.Children)
                                {
                                    flowHtml += $"<tr><td><b>{System.Web.HttpUtility.HtmlEncode(kvp.Key.ToString())}:</b> {formatVariable(kvp.Value.ToString())}</td></tr>";
                                }
                                flowHtml += "</table></td></tr>";
                            }
                            flowHtml += createCardEnd();
                            actionNode.SetAttributeHtml("label", flowHtml);
                            break;
                        case "InvokeConnector":
                            string connHtml = createCardStart(actionType, !string.IsNullOrEmpty(displayName) ? $"Connector: {displayName}" : "Call a connector");
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("connectionReference"), out var connRefNode))
                            {
                                connHtml += createCardBodyRow($"<b>Connection:</b> {generateMultiLineText(System.Web.HttpUtility.HtmlEncode(connRefNode.ToString()))}");
                            }
                            connHtml += createCardEnd();
                            actionNode.SetAttributeHtml("label", connHtml);
                            break;
                        case "EndConversation":
                            actionNode.SetAttributeHtml("label", createCardStart(actionType, !string.IsNullOrEmpty(displayName) ? displayName : "End conversation") + createCardEnd());
                            break;
                        case "EndDialog":
                            actionNode.SetAttributeHtml("label", createCardStart(actionType, !string.IsNullOrEmpty(displayName) ? displayName : "End dialog") + createCardEnd());
                            break;
                        case "OAuthInput":
                            string oauthTitle = "Sign in";
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("title"), out var oauthTitleNode))
                                oauthTitle = oauthTitleNode.ToString();
                            string oauthText = "";
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("text"), out var oauthTextNode))
                                oauthText = oauthTextNode.ToString();
                            string oauthHtml = createCardStart(actionType, $"Sign In: {System.Web.HttpUtility.HtmlEncode(oauthTitle)}");
                            if (!string.IsNullOrEmpty(oauthText))
                                oauthHtml += createCardBodyRow(generateMultiLineText(System.Web.HttpUtility.HtmlEncode(oauthText)));
                            oauthHtml += createCardEnd();
                            actionNode.SetAttributeHtml("label", oauthHtml);
                            break;
                        case "SearchAndSummarize":
                            string searchHtml = createCardStart(actionType, !string.IsNullOrEmpty(displayName) ? displayName : "Search and summarize content");
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("variable"), out var searchVarNode))
                            {
                                searchHtml += createCardBodyRow($"<b>Save to:</b> {formatVariable(searchVarNode.ToString())}");
                            }
                            searchHtml += createCardEnd();
                            actionNode.SetAttributeHtml("label", searchHtml);
                            break;
                        case "RedirectToTopic":
                            string redirectHeader = !string.IsNullOrEmpty(displayName) ? $"Redirect: {displayName}" : "Redirect to topic";
                            string redirectHtml = createCardStart(actionType, redirectHeader);

                            // Show target dialog/topic for BeginDialog actions
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("dialog"), out var dialogNode))
                            {
                                string dialogRef = dialogNode.ToString();
                                // Extract the topic name from the schema name (e.g. cr6b0_agent.topic.SessionAudit -> SessionAudit)
                                string topicName = dialogRef.Contains('.') ? dialogRef.Substring(dialogRef.LastIndexOf('.') + 1) : dialogRef;
                                redirectHtml += createCardBodyRow($"<table border=\"1\" cellpadding=\"4\"><tr><td><b>Target Topic:</b> {System.Web.HttpUtility.HtmlEncode(topicName)}</td></tr></table>");
                            }

                            // Show target action ID for GotoAction (internal jump)
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("actionId"), out var actionIdNode))
                            {
                                redirectHtml += createCardBodyRow($"<table border=\"1\" cellpadding=\"4\"><tr><td><b>Go to:</b> {System.Web.HttpUtility.HtmlEncode(actionIdNode.ToString())}</td></tr></table>");
                            }

                            // Show input bindings if present (for parameterized topic calls)
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("input"), out var redirectInputNode)
                                && redirectInputNode is YamlMappingNode redirectInputMapping
                                && redirectInputMapping.Children.TryGetValue(new YamlScalarNode("binding"), out var redirectBindingNode)
                                && redirectBindingNode is YamlMappingNode redirectBindingMapping)
                            {
                                redirectHtml += "<tr><td cellpadding=\"8\"><table border=\"1\" cellpadding=\"4\"><tr><td><b>Inputs:</b></td></tr>";
                                foreach (var kvp in redirectBindingMapping.Children)
                                {
                                    redirectHtml += $"<tr><td><b>{System.Web.HttpUtility.HtmlEncode(kvp.Key.ToString())}:</b> {formatVariable(kvp.Value.ToString())}</td></tr>";
                                }
                                redirectHtml += "</table></td></tr>";
                            }

                            // Show output bindings if present
                            if (((YamlMappingNode)action).Children.TryGetValue(new YamlScalarNode("output"), out var redirectOutputNode)
                                && redirectOutputNode is YamlMappingNode redirectOutputMapping
                                && redirectOutputMapping.Children.TryGetValue(new YamlScalarNode("binding"), out var redirectOutputBindingNode)
                                && redirectOutputBindingNode is YamlMappingNode redirectOutputBindingMapping)
                            {
                                redirectHtml += "<tr><td cellpadding=\"8\"><table border=\"1\" cellpadding=\"4\"><tr><td><b>Outputs:</b></td></tr>";
                                foreach (var kvp in redirectOutputBindingMapping.Children)
                                {
                                    redirectHtml += $"<tr><td><b>{System.Web.HttpUtility.HtmlEncode(kvp.Key.ToString())}:</b> {formatVariable(kvp.Value.ToString())}</td></tr>";
                                }
                                redirectHtml += "</table></td></tr>";
                            }

                            redirectHtml += createCardEnd();
                            actionNode.SetAttributeHtml("label", redirectHtml);
                            break;
                        default:
                            actionNode.SetAttribute("label", CharsetHelper.GetSafeName(actionName));
                            break;
                    }
                    if (conditionCluster == null)
                    {
                        Edge edge = rootGraph.GetOrAddEdge((Node)prevNode, actionNode, actionName);
                    }
                    if (actionType != "ConditionGroup")
                    {
                        prevNode = actionNode;
                    }
                }
            }
            return returnNode;
        }

        /// <summary>
        /// Creates the start of a card-style node label with a colored header bar, similar to Copilot Studio UI.
        /// </summary>
        private string createCardStart(string actionType, string headerText, string iconType = null)
        {
            string borderColor = GraphColours.GetColourForAction(actionType);
            string iconName = iconType ?? actionType;
            return $"<table border=\"2\" cellborder=\"0\" cellspacing=\"0\" cellpadding=\"0\" color=\"{borderColor}\" bgcolor=\"white\" style=\"rounded\">"
                 + $"<tr><td bgcolor=\"{borderColor}\" cellpadding=\"8\" align=\"left\">"
                 + $"<table border=\"0\" cellspacing=\"0\" cellpadding=\"2\"><tr>"
                 + $"<td width=\"20\"><img src=\"{folderPath + @"Resources\" + iconName}.png\" /></td>"
                 + $"<td><font color=\"white\"><b>  {headerText}</b></font></td>"
                 + $"<td>   </td>"
                 + $"</tr></table></td></tr>";
        }

        /// <summary>
        /// Creates a body row for the card with padding.
        /// </summary>
        private string createCardBodyRow(string content)
        {
            return $"<tr><td cellpadding=\"8\" align=\"left\">{content}</td></tr>";
        }

        /// <summary>
        /// Closes the card table.
        /// </summary>
        private string createCardEnd()
        {
            return "</table>";
        }

        /// <summary>
        /// Formats a variable reference with the (x) badge style, similar to Copilot Studio.
        /// </summary>
        private string formatVariable(string varName)
        {
            return $"<font color=\"#0078d4\"><b>(x)</b></font> {System.Web.HttpUtility.HtmlEncode(varName)}";
        }

        /// <summary>
        /// Parses a condition expression and formats it with individual rows and Or/And separators.
        /// </summary>
        private string formatConditionExpression(string conditionString)
        {
            if (conditionString.StartsWith("="))
                conditionString = conditionString.Substring(1).Trim();

            // Try to split by || (OR)
            var orParts = Regex.Split(conditionString, @"\s*\|\|\s*");
            if (orParts.Length > 1)
            {
                string html = "<table border=\"0\" cellpadding=\"3\">";
                for (int i = 0; i < orParts.Length; i++)
                {
                    if (i > 0)
                        html += "<tr><td><b>Or</b></td></tr>";
                    html += $"<tr><td><table border=\"1\" cellpadding=\"4\"><tr><td>{formatSingleCondition(orParts[i].Trim())}</td></tr></table></td></tr>";
                }
                html += "</table>";
                return html;
            }

            // Try && (AND)
            var andParts = Regex.Split(conditionString, @"\s*&&\s*");
            if (andParts.Length > 1)
            {
                string html = "<table border=\"0\" cellpadding=\"3\">";
                for (int i = 0; i < andParts.Length; i++)
                {
                    if (i > 0)
                        html += "<tr><td><b>And</b></td></tr>";
                    html += $"<tr><td><table border=\"1\" cellpadding=\"4\"><tr><td>{formatSingleCondition(andParts[i].Trim())}</td></tr></table></td></tr>";
                }
                html += "</table>";
                return html;
            }

            // Single condition
            return $"<table border=\"1\" cellpadding=\"4\"><tr><td>{formatSingleCondition(conditionString)}</td></tr></table>";
        }

        /// <summary>
        /// Formats a single condition clause with variable badge, operator text, and value.
        /// </summary>
        private string formatSingleCondition(string condition)
        {
            condition = condition.Trim();
            // Try to parse: Variable operator "value" or Variable operator value
            var match = Regex.Match(condition, @"^(.+?)\s*(<>|!=|>=|<=|=|>|<)\s*(.+)$");
            if (match.Success)
            {
                string variable = match.Groups[1].Value.Trim();
                string op = match.Groups[2].Value;
                string value = match.Groups[3].Value.Trim();

                string opText = op switch
                {
                    "=" => "is equal to",
                    "<>" or "!=" => "is not equal to",
                    ">" => "is greater than",
                    "<" => "is less than",
                    ">=" => "is greater than or equal to",
                    "<=" => "is less than or equal to",
                    _ => op
                };

                return $"{formatVariable(variable)}<br/>{System.Web.HttpUtility.HtmlEncode(opText)}<br/>{System.Web.HttpUtility.HtmlEncode(value)}";
            }

            return generateMultiLineText(System.Web.HttpUtility.HtmlEncode(condition));
        }

        private string GetActionType(YamlMappingNode action)
        {
            string actionType = ((YamlMappingNode)action).Children[new YamlScalarNode("kind")].ToString();
            switch (actionType)
            {
                case "AdaptiveCardPrompt":
                    return "AdaptiveCard";
                case "ConditionGroup":
                    return "ConditionGroup";
                case "InvokeAIBuilderModelAction":
                    return "AIModel";
                case "SendActivity":
                    return "Message";
                case "SetVariable":
                    return "SetVariable";
                case "LogCustomTelemetryEvent":
                    return "LogCustomTelemetry";
                case "Question":
                    return "Question";
                case "InvokeFlowAction":
                    return "InvokeFlow";
                case "InvokeConnectorAction":
                    return "InvokeConnector";
                case "EndConversation":
                    return "EndConversation";
                case "EndDialog":
                    return "EndDialog";
                case "OAuthInput":
                    return "OAuthInput";
                case "SearchAndSummarizeContent":
                    return "SearchAndSummarize";
                case "GotoAction":
                case "BeginDialog":
                    return "RedirectToTopic";
                case "CancelAllDialogs":
                    return "CancelAllDialogs";
                default:
                    return actionType;
            }
        }

        private string generateImageFiles(RootGraph rootGraph, bool showSubactions)
        {
            //Generate image files
            string filename = topic.getTopicFileName() + (showSubactions ? "-detailed" : "");
            rootGraph.ToPngFile(folderPath + filename + ".png");
            rootGraph.ToSvgFile(folderPath + filename + ".svg");
            // Post-process SVG to embed referenced PNG icons as base64 data URIs,
            // so the SVG is self-contained and renders correctly when viewed from HTML pages.
            EmbedImagesInSvg(folderPath + filename + ".svg");
            return filename;
        }

        /// <summary>
        /// Reads an SVG file and replaces all xlink:href/href image references pointing
        /// to local PNG files with inline base64 data URIs.
        /// </summary>
        private static void EmbedImagesInSvg(string svgFilePath)
        {
            if (!File.Exists(svgFilePath)) return;
            string svgContent = File.ReadAllText(svgFilePath);
            // Match xlink:href="..." or href="..." attributes that point to .png files
            string pattern = @"(xlink:href|href)=""([^""]+\.png)""";
            string result = Regex.Replace(svgContent, pattern, match =>
            {
                string attr = match.Groups[1].Value;
                string imagePath = match.Groups[2].Value;
                // Resolve relative paths against the SVG's directory
                if (!Path.IsPathRooted(imagePath))
                {
                    imagePath = Path.Combine(Path.GetDirectoryName(svgFilePath), imagePath);
                }
                if (File.Exists(imagePath))
                {
                    byte[] imageBytes = File.ReadAllBytes(imagePath);
                    string base64 = Convert.ToBase64String(imageBytes);
                    return $"{attr}=\"data:image/png;base64,{base64}\"";
                }
                return match.Value;
            });
            File.WriteAllText(svgFilePath, result);
        }

        //splits a text into multiple lines (<br/> for line breaks), with each line having a maximum of 65 characters
        private string generateMultiLineText(string text)
        {
            const int maxLineLength = 65;
            string[] words = text.Split(' ');
            string multiLineText = "";
            int lineLength = 0;
            for (var counter = 0; counter < words.Length; counter++)
            {
                string word = words[counter];
                // Force-break words longer than the max line length
                while (word.Length > maxLineLength)
                {
                    if (lineLength > 0)
                    {
                        multiLineText += "<br/>";
                        lineLength = 0;
                    }
                    multiLineText += word.Substring(0, maxLineLength) + "<br/>";
                    word = word.Substring(maxLineLength);
                }
                if (lineLength + word.Length + 1 >= maxLineLength && lineLength > 0)
                {
                    multiLineText += "<br/>";
                    lineLength = 0;
                }
                multiLineText = multiLineText + word + " ";
                lineLength += word.Length + 1;
            }
            return multiLineText;
        }

        private string RenderCardElement(JObject element)
        {
            if (element == null) return "";

            string elementType = element["type"]?.ToString() ?? "";
            return elementType switch
            {
                "TextBlock" => RenderTextBlock(element),
                "Input.Text" => RenderInputText(element),
                "Input.ChoiceSet" => RenderInputChoiceSet(element),
                /*"Input.Date" => RenderInputDate(element),
                "Input.Number" => RenderInputNumber(element),
                "Input.Toggle" => RenderInputToggle(element),
                "Container" => RenderContainer(element),
                "ColumnSet" => RenderColumnSet(element),
                "Image" => RenderImage(element),
                "FactSet" => RenderFactSet(element),
                "ActionSet" => RenderActionSet(element),*/
                _ => $"<i>Element: {System.Web.HttpUtility.HtmlEncode(elementType)}</i>"
            };
        }

        private string RenderCardAction(JObject action)
        {
            if (action == null) return "";

            string actionType = action["type"]?.ToString() ?? "";
            string title = action["title"]?.ToString() ?? "Action";

            return actionType switch
            {
                "Action.Submit" => $"<table border=\"1\" cellpadding=\"4\" bgcolor=\"#f3f2f1\"><tr><td>{System.Web.HttpUtility.HtmlEncode(title)}</td></tr></table>",
                "Action.OpenUrl" => $"<table border=\"1\" cellpadding=\"4\" bgcolor=\"#f3f2f1\"><tr><td>{System.Web.HttpUtility.HtmlEncode(title)}</td></tr></table>",
                "Action.ShowCard" => $"<table border=\"1\" cellpadding=\"4\" bgcolor=\"#f3f2f1\"><tr><td>{System.Web.HttpUtility.HtmlEncode(title)}</td></tr></table>",
                "Action.Execute" => $"<table border=\"1\" cellpadding=\"4\" bgcolor=\"#f3f2f1\"><tr><td>{System.Web.HttpUtility.HtmlEncode(title)}</td></tr></table>",
                "Action.ToggleVisibility" => $"<table border=\"1\" cellpadding=\"4\" bgcolor=\"#f3f2f1\"><tr><td>{System.Web.HttpUtility.HtmlEncode(title)}</td></tr></table>",
                _ => $"<table border=\"1\" cellpadding=\"4\" bgcolor=\"#f3f2f1\"><tr><td>{System.Web.HttpUtility.HtmlEncode(title)}</td></tr></table>"
            };
        }

        private string RenderTextBlock(JObject element)
        {
            string text = element["text"]?.ToString() ?? "";
            bool isSubtle = element["isSubtle"]?.ToObject<bool>() ?? false;
            string weight = element["weight"]?.ToString() ?? "";
            string encodedText = generateMultiLineText(System.Web.HttpUtility.HtmlEncode(text));
            if (weight == "Bolder")
                encodedText = $"<b>{encodedText}</b>";
            if (isSubtle)
                encodedText = $"<i>{encodedText}</i>";
            return encodedText;
        }

        private string RenderInputText(JObject element)
        {
            string id = element["id"]?.ToString() ?? "";
            string placeholder = element["placeholder"]?.ToString() ?? "";
            string label = element["label"]?.ToString() ?? "";
            bool isMultiline = element["isMultiline"]?.ToObject<bool>() ?? false;
            bool isRequired = element["isRequired"]?.ToObject<bool>() ?? false;

            // Graphviz HTML-label cells cannot mix text and <table> children.
            // Wrap everything in a single table so a <td> only ever contains one table.
            string requiredMark = isRequired ? "<font color=\"red\"> *</font>" : "";
            string html = "<table border=\"0\" cellpadding=\"2\">";
            if (!string.IsNullOrEmpty(label))
            {
                html += $"<tr><td>{generateMultiLineText(System.Web.HttpUtility.HtmlEncode(label))}{requiredMark}</td></tr>";
            }
            if (!string.IsNullOrEmpty(placeholder))
            {
                html += $"<tr><td><table border=\"1\" cellpadding=\"4\" color=\"#cccccc\"><tr><td><font color=\"#999999\">{generateMultiLineText(System.Web.HttpUtility.HtmlEncode(placeholder))}</font></td></tr></table></td></tr>";
            }
            html += "</table>";
            return html;
        }

        private string RenderInputChoiceSet(JObject element)
        {
            string id = element["id"]?.ToString() ?? "";
            string label = element["label"]?.ToString() ?? "";
            bool isMultiSelect = element["isMultiSelect"]?.ToObject<bool>() ?? false;

            string choicesText = "";
            string indicator = isMultiSelect ? "[ ]" : "( )";
            if (element["choices"] is JArray choices)
            {
                var choiceList = new List<string>();
                foreach (var choice in choices)
                {
                    string title = choice["title"]?.ToString() ?? "";
                    choiceList.Add($"  {indicator} {System.Web.HttpUtility.HtmlEncode(title)}");
                }
                choicesText = string.Join("<br/>", choiceList);
            }

            bool isRequired = element["isRequired"]?.ToObject<bool>() ?? false;
            string requiredMark = isRequired ? "<font color=\"red\"> *</font>" : "";
            string selectLabel = !string.IsNullOrEmpty(label) ? $"{generateMultiLineText(System.Web.HttpUtility.HtmlEncode(label))}{requiredMark}" : "";
            return $"{selectLabel}<br/>{choicesText}";
        }

        private string RenderInputDate(JObject element)
        {
            string id = element["id"]?.ToString() ?? "";
            string label = element["label"]?.ToString() ?? "";
            string min = element["min"]?.ToString() ?? "";
            string max = element["max"]?.ToString() ?? "";

            string constraints = "";
            if (!string.IsNullOrEmpty(min) || !string.IsNullOrEmpty(max))
            {
                constraints = $"<br/>Range: {min} to {max}";
            }

            return $"<b>Date Input:</b> {label} (ID: {id}){constraints}";
        }

        private string RenderInputNumber(JObject element)
        {
            string id = element["id"]?.ToString() ?? "";
            string label = element["label"]?.ToString() ?? "";
            string min = element["min"]?.ToString() ?? "";
            string max = element["max"]?.ToString() ?? "";

            string constraints = "";
            if (!string.IsNullOrEmpty(min) || !string.IsNullOrEmpty(max))
            {
                constraints = $"<br/>Range: {min} to {max}";
            }

            return $"<b>Number Input:</b> {label} (ID: {id}){constraints}";
        }

        private string RenderInputToggle(JObject element)
        {
            string id = element["id"]?.ToString() ?? "";
            string title = element["title"]?.ToString() ?? "";
            string valueOn = element["valueOn"]?.ToString() ?? "true";
            string valueOff = element["valueOff"]?.ToString() ?? "false";

            return $"<b>Toggle:</b> {title} (ID: {id})<br/>Values: {valueOff}/{valueOn}";
        }

        private string RenderContainer(JObject element)
        {
            int itemCount = 0;
            if (element["items"] is JArray items)
            {
                itemCount = items.Count;
            }

            return $"<b>Container</b> with {itemCount} nested item(s)";
        }

        private string RenderColumnSet(JObject element)
        {
            int columnCount = 0;
            if (element["columns"] is JArray columns)
            {
                columnCount = columns.Count;
            }

            return $"<b>Column Set</b> with {columnCount} column(s)";
        }

        private string RenderImage(JObject element)
        {
            string url = element["url"]?.ToString() ?? "";
            string altText = element["altText"]?.ToString() ?? "";
            string size = element["size"]?.ToString() ?? "auto";

            return $"<b>Image:</b> {(!string.IsNullOrEmpty(altText) ? altText : "No alt text")}<br/>Size: {size}";
        }

        private string RenderFactSet(JObject element)
        {
            int factCount = 0;
            if (element["facts"] is JArray facts)
            {
                factCount = facts.Count;
                var factList = new List<string>();
                foreach (var fact in facts.Take(3)) // Show first 3 facts
                {
                    string title = fact["title"]?.ToString() ?? "";
                    string value = fact["value"]?.ToString() ?? "";
                    factList.Add($"{title}: {value}");
                }
                string factText = string.Join("<br/>", factList);
                if (factCount > 3) factText += $"<br/>... and {factCount - 3} more";
                return $"<b>Fact Set</b> ({factCount} facts):<br/>{factText}";
            }

            return $"<b>Fact Set</b> with {factCount} fact(s)";
        }

        private string RenderActionSet(JObject element)
        {
            int actionCount = 0;
            if (element["actions"] is JArray actions)
            {
                actionCount = actions.Count;
            }

            return $"<b>Action Set</b> with {actionCount} action(s)";
        }

        // Add this method to create constraint edges between parallel nodes
        //TODO potentially no longer required
        private void CreateLevelConstraints(RootGraph rootGraph, SubGraph cluster, List<Node> parallelNodes)
        {
            if (parallelNodes.Count <= 1) return;
            Console.WriteLine(topic.Name);
            foreach (Node node in parallelNodes)
            {
                Console.WriteLine("Parallel Node: " + node.GetName());
            }
            Console.WriteLine();
            // Create invisible edges between parallel nodes to maintain same level
            for (int i = 0; i < parallelNodes.Count - 1; i++)
            {
                Edge constraintEdge = rootGraph.GetOrAddEdge(parallelNodes[i], parallelNodes[i + 1], "constraint_" + i);
                constraintEdge.SetAttribute("style", "invis");
                constraintEdge.SetAttribute("constraint", "false"); // Don't affect layout, just ranking
                constraintEdge.SetAttribute("weight", "0");
            }
            // Create a rank subgraph
            string rankName = "rank_" + parallelNodes[0].GetName() + "_group";
            SubGraph rankGroup = cluster.GetOrAddSubgraph(rankName);
            rankGroup.SetAttribute("rank", "same");
            rankGroup.SetAttribute("style", "invis");

            foreach (var node in parallelNodes)
            {
                rankGroup.AddExisting(node);
            }
        }
    }

    public static class GraphColours
    {
        public static string TriggerColour = "#0077ff";
        public static string TriggerFillColour = "#e7f4ff";
        public static string SetVariableFillColour = "#e7f4ff";
        public static string SetVariableColour = "#118dff";
        public static string ConditionFillColour = "#e7f4ff";
        public static string ConditionColour = "#118dff";
        public static string MessageColour = "#672367";
        public static string MessageFillColour = "#f0e9f0";
        public static string AdaptiveCardColour = "#672367";
        public static string AdaptiveCardFillColour = "#f0e9f0";
        public static string AIModelCardColour = "#672367";
        public static string AIModelCardFillColour = "#f0e9f0";
        public static string CancelAllDialogsColour = "#6bb700";
        public static string CancelAllDialogsFillColour = "#f0f8e6";
        public static string LogCustomTelemetryColour = "#242424";
        public static string LogCustomTelemetryFillColour = "#edebe9";
        public static string QuestionColour = "#672367";
        public static string QuestionFillColour = "#f0e9f0";
        public static string InvokeFlowColour = "#0078d4";
        public static string InvokeFlowFillColour = "#e5f1fb";
        public static string InvokeConnectorColour = "#0078d4";
        public static string InvokeConnectorFillColour = "#e5f1fb";
        public static string EndConversationColour = "#6bb700";
        public static string EndConversationFillColour = "#f0f8e6";
        public static string EndDialogColour = "#6bb700";
        public static string EndDialogFillColour = "#f0f8e6";
        public static string OAuthInputColour = "#0078d4";
        public static string OAuthInputFillColour = "#e5f1fb";
        public static string SearchAndSummarizeColour = "#0078d4";
        public static string SearchAndSummarizeFillColour = "#e5f1fb";
        public static string RedirectToTopicColour = "#0077ff";
        public static string RedirectToTopicFillColour = "#e7f4ff";

        public static string GetColourForAction(string actionType)
        {
            return actionType switch
            {
                "Trigger" => TriggerColour,
                "Message" => MessageColour,
                "Question" => QuestionColour,
                "CancelAllDialogs" => CancelAllDialogsColour,
                "SetVariable" => SetVariableColour,
                "ConditionGroup" => ConditionColour,
                "LogCustomTelemetry" => LogCustomTelemetryColour,
                "AdaptiveCard" => AdaptiveCardColour,
                "AIModel" => AIModelCardColour,
                "InvokeFlow" => InvokeFlowColour,
                "InvokeConnector" => InvokeConnectorColour,
                "EndConversation" => EndConversationColour,
                "EndDialog" => EndDialogColour,
                "OAuthInput" => OAuthInputColour,
                "SearchAndSummarize" => SearchAndSummarizeColour,
                "RedirectToTopic" => RedirectToTopicColour,
                _ => "black",
            };
        }

        public static string GetFillColourForAction(string actionType)
        {
            string colour = actionType switch
            {
                "Trigger" => TriggerFillColour,
                "Message" => MessageFillColour,
                "Question" => QuestionFillColour,
                "CancelAllDialogs" => CancelAllDialogsFillColour,
                "SetVariable" => SetVariableFillColour,
                "ConditionGroup" => ConditionFillColour,
                "LogCustomTelemetry" => LogCustomTelemetryFillColour,
                "AdaptiveCard" => AdaptiveCardFillColour,
                "AIModel" => AIModelCardFillColour,
                "InvokeFlow" => InvokeFlowFillColour,
                "InvokeConnector" => InvokeConnectorFillColour,
                "EndConversation" => EndConversationFillColour,
                "EndDialog" => EndDialogFillColour,
                "OAuthInput" => OAuthInputFillColour,
                "SearchAndSummarize" => SearchAndSummarizeFillColour,
                "RedirectToTopic" => RedirectToTopicFillColour,
                _ => "red",
            };
            return colour;
        }
    }
}
