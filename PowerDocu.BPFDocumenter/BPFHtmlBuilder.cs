using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using PowerDocu.Common;

namespace PowerDocu.BPFDocumenter
{
    class BPFHtmlBuilder : HtmlBuilder
    {
        private readonly BPFDocumentationContent content;
        private readonly string mainFileName;

        public BPFHtmlBuilder(BPFDocumentationContent contentDocumentation)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);
            WriteDefaultStylesheet(content.folderPath);
            mainFileName = ("bpf-" + content.filename + ".html").Replace(" ", "-");

            addOverviewPage();
            NotificationHelper.SendNotification("Created HTML documentation for Business Process Flow: " + content.bpf.GetDisplayName());
        }

        private string getNavigationHtml()
        {
            var navItemsList = new List<(string label, string href)>();
            if (content.context?.Solution != null)
            {
                if (content.context?.Config?.documentSolution == true)
                    navItemsList.Add(("Solution", "../" + CrossDocLinkHelper.GetSolutionDocHtmlPath(content.context.Solution.UniqueName)));
                else
                    navItemsList.Add((content.context.Solution.UniqueName, ""));
            }
            navItemsList.AddRange(new (string label, string href)[]
            {
                ("Overview", "#overview"),
                ("Stages", "#stages"),
                ("Stage Details", "#stage-details"),
                ("Properties", "#properties")
            });
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"<div class=\"nav-title\">{Encode(content.bpf.GetDisplayName())}</div>");
            sb.Append(NavigationList(navItemsList));
            return sb.ToString();
        }

        private void addOverviewPage()
        {
            StringBuilder body = new StringBuilder();

            // Overview
            body.AppendLine(HeadingWithId(1, content.bpf.GetDisplayName(), "overview"));

            body.Append(TableStart("Property", "Value"));
            body.Append(TableRow("Name", content.bpf.GetDisplayName()));
            body.Append(TableRow("Unique Name", content.bpf.UniqueName ?? ""));
            body.Append(TableRow("Primary Entity", content.GetTableDisplayName(content.bpf.PrimaryEntity)));
            body.Append(TableRow("State", content.bpf.GetStateLabel()));
            body.Append(TableRow("Number of Stages", content.bpf.Stages.Count.ToString()));
            if (!string.IsNullOrEmpty(content.bpf.Description))
                body.Append(TableRow("Description", content.bpf.Description));
            if (!string.IsNullOrEmpty(content.bpf.IntroducedVersion))
                body.Append(TableRow("Version", content.bpf.IntroducedVersion));
            body.Append(TableRow(content.headerDocumentationGenerated, PowerDocuReleaseHelper.GetTimestampWithVersion()));
            body.AppendLine(TableEnd());

            // Stages overview
            if (content.bpf.Stages.Count > 0)
            {
                body.AppendLine(HeadingWithId(2, content.headerStages, "stages"));
                body.AppendLine($"<p>This Business Process Flow has {content.bpf.Stages.Count} stage(s).</p>");

                body.Append(TableStart("#", "Stage Name", "Entity", "Category", "Steps", "Required Fields"));
                for (int i = 0; i < content.bpf.Stages.Count; i++)
                {
                    BPFStage stage = content.bpf.Stages[i];
                    int requiredCount = stage.Steps.Count(s => s.IsRequired);
                    body.Append(TableRow(
                        (i + 1).ToString(),
                        stage.Name ?? "",
                        content.GetTableDisplayName(stage.EntityName),
                        stage.GetStageCategoryLabel(),
                        stage.Steps.Count.ToString(),
                        requiredCount.ToString()
                    ));
                }
                body.AppendLine(TableEnd());
            }

            // Stage details
            if (content.bpf.Stages.Count > 0)
            {
                body.AppendLine(HeadingWithId(2, content.headerStageDetails, "stage-details"));

                for (int i = 0; i < content.bpf.Stages.Count; i++)
                {
                    BPFStage stage = content.bpf.Stages[i];
                    body.AppendLine(Heading(3, $"Stage {i + 1}: {stage.Name}"));

                    body.Append(TableStart("Property", "Value"));
                    body.Append(TableRow("Entity", content.GetTableDisplayName(stage.EntityName)));
                    body.Append(TableRow("Category", stage.GetStageCategoryLabel()));
                    if (!string.IsNullOrEmpty(stage.NextStageId))
                    {
                        BPFStage nextStage = content.bpf.Stages.FirstOrDefault(s => s.StageId == stage.NextStageId);
                        body.Append(TableRow("Next Stage", nextStage?.Name ?? stage.NextStageId));
                    }
                    body.AppendLine(TableEnd());

                    // Steps/Controls
                    if (stage.Steps.Count > 0)
                    {
                        body.Append(TableStart("#", "Control", "Data Field", "Required"));
                        int stepNum = 1;
                        foreach (BPFStep step in stage.Steps)
                        {
                            foreach (BPFControl control in step.Controls)
                            {
                                body.Append(TableRow(
                                    stepNum.ToString(),
                                    control.DisplayName ?? "",
                                    control.DataFieldName ?? "",
                                    step.IsRequired ? "Yes" : "No"
                                ));
                                stepNum++;
                            }
                            if (step.Controls.Count == 0)
                            {
                                body.Append(TableRow(
                                    stepNum.ToString(),
                                    step.Description ?? "(No control)",
                                    "",
                                    step.IsRequired ? "Yes" : "No"
                                ));
                                stepNum++;
                            }
                        }
                        body.AppendLine(TableEnd());
                    }

                    // Condition branches
                    if (stage.ConditionBranches.Count > 0)
                    {
                        body.AppendLine("<p><strong>Conditional Branching:</strong></p>");
                        body.Append(TableStart("Condition", "Target Stage"));
                        foreach (BPFConditionBranch branch in stage.ConditionBranches)
                        {
                            BPFStage targetStage = content.bpf.Stages.FirstOrDefault(s => s.StageId == branch.TargetStageId);
                            body.Append(TableRow(
                                !string.IsNullOrEmpty(branch.Description) ? branch.Description : "(Default)",
                                targetStage?.Name ?? branch.TargetStageId ?? ""
                            ));
                        }
                        body.AppendLine(TableEnd());
                    }
                }
            }

            // Properties
            body.AppendLine(HeadingWithId(2, content.headerProperties, "properties"));
            body.Append(TableStart("Property", "Value"));
            body.Append(TableRow("ID", content.bpf.ID ?? ""));
            body.Append(TableRow("Unique Name", content.bpf.UniqueName ?? ""));
            body.Append(TableRow("Primary Entity", content.bpf.PrimaryEntity ?? ""));
            body.Append(TableRow("Business Process Type", content.bpf.BusinessProcessType.ToString()));
            body.Append(TableRow("State", content.bpf.GetStateLabel()));
            body.Append(TableRow("Trigger On Create", content.bpf.TriggerOnCreate ? "Yes" : "No"));
            body.Append(TableRow("Is Customizable", content.bpf.IsCustomizable ? "Yes" : "No"));
            if (!string.IsNullOrEmpty(content.bpf.IntroducedVersion))
                body.Append(TableRow("Introduced Version", content.bpf.IntroducedVersion));
            body.AppendLine(TableEnd());

            SaveHtmlFile(Path.Combine(content.folderPath, mainFileName),
                WrapInHtmlPage($"Business Process Flow - {content.bpf.GetDisplayName()}", body.ToString(), getNavigationHtml()));
        }
    }
}
