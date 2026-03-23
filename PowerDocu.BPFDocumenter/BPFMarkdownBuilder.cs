using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PowerDocu.Common;
using Grynwald.MarkdownGenerator;

namespace PowerDocu.BPFDocumenter
{
    class BPFMarkdownBuilder : MarkdownBuilder
    {
        private readonly BPFDocumentationContent content;
        private readonly string mainDocumentFileName;
        private readonly MdDocument mainDocument;
        private readonly DocumentSet<MdDocument> set;

        public BPFMarkdownBuilder(BPFDocumentationContent contentDocumentation)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);
            mainDocumentFileName = ("bpf-" + content.filename + ".md").Replace(" ", "-");
            set = new DocumentSet<MdDocument>();
            mainDocument = set.CreateMdDocument(mainDocumentFileName);

            addOverview();
            addStagesOverview();
            addStageDetails();
            addProperties();

            set.Save(content.folderPath);
            NotificationHelper.SendNotification("Created Markdown documentation for Business Process Flow: " + content.bpf.GetDisplayName());
        }

        private void addOverview()
        {
            mainDocument.Root.Add(new MdHeading(content.bpf.GetDisplayName(), 1));

            if (content.context?.Solution != null)
            {
                if (content.context?.Config?.documentSolution == true)
                    mainDocument.Root.Add(new MdParagraph(new MdCompositeSpan(new MdTextSpan("Solution: "), new MdLinkSpan(content.context.Solution.UniqueName, "../" + CrossDocLinkHelper.GetSolutionDocMdPath(content.context.Solution.UniqueName)))));
                else
                    mainDocument.Root.Add(new MdParagraph(new MdTextSpan("Solution: " + content.context.Solution.UniqueName)));
            }

            List<MdTableRow> tableRows = new List<MdTableRow>
            {
                new MdTableRow("Name", content.bpf.GetDisplayName()),
                new MdTableRow("Unique Name", content.bpf.UniqueName ?? ""),
                new MdTableRow("Primary Entity", content.GetTableDisplayName(content.bpf.PrimaryEntity)),
                new MdTableRow("State", content.bpf.GetStateLabel()),
                new MdTableRow("Number of Stages", content.bpf.Stages.Count.ToString()),
                new MdTableRow(content.headerDocumentationGenerated, PowerDocuReleaseHelper.GetTimestampWithVersion())
            };
            if (!string.IsNullOrEmpty(content.bpf.Description))
                tableRows.Insert(tableRows.Count - 1, new MdTableRow("Description", content.bpf.Description));
            if (!string.IsNullOrEmpty(content.bpf.IntroducedVersion))
                tableRows.Insert(tableRows.Count - 1, new MdTableRow("Version", content.bpf.IntroducedVersion));

            mainDocument.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
        }

        private void addStagesOverview()
        {
            if (content.bpf.Stages.Count == 0) return;

            mainDocument.Root.Add(new MdHeading(content.headerStages, 2));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"This Business Process Flow has {content.bpf.Stages.Count} stage(s).")));

            List<MdTableRow> tableRows = new List<MdTableRow>();
            for (int i = 0; i < content.bpf.Stages.Count; i++)
            {
                BPFStage stage = content.bpf.Stages[i];
                int requiredCount = stage.Steps.Count(s => s.IsRequired);
                tableRows.Add(new MdTableRow(
                    (i + 1).ToString(),
                    stage.Name ?? "",
                    content.GetTableDisplayName(stage.EntityName),
                    stage.GetStageCategoryLabel(),
                    stage.Steps.Count.ToString(),
                    requiredCount.ToString()
                ));
            }
            mainDocument.Root.Add(new MdTable(
                new MdTableRow("#", "Stage Name", "Entity", "Category", "Steps", "Required Fields"),
                tableRows));
        }

        private void addStageDetails()
        {
            if (content.bpf.Stages.Count == 0) return;

            mainDocument.Root.Add(new MdHeading(content.headerStageDetails, 2));

            for (int i = 0; i < content.bpf.Stages.Count; i++)
            {
                BPFStage stage = content.bpf.Stages[i];
                mainDocument.Root.Add(new MdHeading($"Stage {i + 1}: {stage.Name}", 3));

                // Stage info
                List<MdTableRow> infoRows = new List<MdTableRow>
                {
                    new MdTableRow("Entity", content.GetTableDisplayName(stage.EntityName)),
                    new MdTableRow("Category", stage.GetStageCategoryLabel())
                };
                if (!string.IsNullOrEmpty(stage.NextStageId))
                {
                    BPFStage nextStage = content.bpf.Stages.FirstOrDefault(s => s.StageId == stage.NextStageId);
                    infoRows.Add(new MdTableRow("Next Stage", nextStage?.Name ?? stage.NextStageId));
                }
                mainDocument.Root.Add(new MdTable(new MdTableRow("Property", "Value"), infoRows));

                // Steps/Controls
                if (stage.Steps.Count > 0)
                {
                    List<MdTableRow> stepRows = new List<MdTableRow>();
                    int stepNum = 1;
                    foreach (BPFStep step in stage.Steps)
                    {
                        foreach (BPFControl control in step.Controls)
                        {
                            stepRows.Add(new MdTableRow(
                                stepNum.ToString(),
                                control.DisplayName ?? "",
                                control.DataFieldName ?? "",
                                step.IsRequired ? "Yes" : "No"
                            ));
                            stepNum++;
                        }
                        if (step.Controls.Count == 0)
                        {
                            stepRows.Add(new MdTableRow(
                                stepNum.ToString(),
                                step.Description ?? "(No control)",
                                "",
                                step.IsRequired ? "Yes" : "No"
                            ));
                            stepNum++;
                        }
                    }
                    if (stepRows.Count > 0)
                    {
                        mainDocument.Root.Add(new MdTable(
                            new MdTableRow("#", "Control", "Data Field", "Required"),
                            stepRows));
                    }
                }

                // Condition branches
                if (stage.ConditionBranches.Count > 0)
                {
                    mainDocument.Root.Add(new MdParagraph(new MdStrongEmphasisSpan("Conditional Branching:")));
                    List<MdTableRow> branchRows = new List<MdTableRow>();
                    foreach (BPFConditionBranch branch in stage.ConditionBranches)
                    {
                        BPFStage targetStage = content.bpf.Stages.FirstOrDefault(s => s.StageId == branch.TargetStageId);
                        branchRows.Add(new MdTableRow(
                            !string.IsNullOrEmpty(branch.Description) ? branch.Description : "(Default)",
                            targetStage?.Name ?? branch.TargetStageId ?? ""
                        ));
                    }
                    mainDocument.Root.Add(new MdTable(
                        new MdTableRow("Condition", "Target Stage"),
                        branchRows));
                }
            }
        }

        private void addProperties()
        {
            mainDocument.Root.Add(new MdHeading(content.headerProperties, 2));

            List<MdTableRow> tableRows = new List<MdTableRow>
            {
                new MdTableRow("ID", content.bpf.ID ?? ""),
                new MdTableRow("Unique Name", content.bpf.UniqueName ?? ""),
                new MdTableRow("Primary Entity", content.bpf.PrimaryEntity ?? ""),
                new MdTableRow("Business Process Type", content.bpf.BusinessProcessType.ToString()),
                new MdTableRow("State", content.bpf.GetStateLabel()),
                new MdTableRow("Trigger On Create", content.bpf.TriggerOnCreate ? "Yes" : "No"),
                new MdTableRow("Is Customizable", content.bpf.IsCustomizable ? "Yes" : "No")
            };
            if (!string.IsNullOrEmpty(content.bpf.IntroducedVersion))
                tableRows.Add(new MdTableRow("Introduced Version", content.bpf.IntroducedVersion));

            mainDocument.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
        }
    }
}
