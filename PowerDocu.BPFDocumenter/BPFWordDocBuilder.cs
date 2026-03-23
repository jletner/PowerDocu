using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerDocu.Common;

namespace PowerDocu.BPFDocumenter
{
    class BPFWordDocBuilder : WordDocBuilder
    {
        private readonly BPFDocumentationContent content;

        public BPFWordDocBuilder(BPFDocumentationContent contentDocumentation, string template)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);
            string filename = InitializeWordDocument(content.folderPath + content.filename, template);
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filename, true))
            {
                mainPart = wordDocument.MainDocumentPart;
                body = mainPart.Document.Body;
                PrepareDocument(!String.IsNullOrEmpty(template));

                addOverview();
                addStagesOverview();
                addStageDetails();
                addProperties();
            }
            NotificationHelper.SendNotification("Created Word documentation for Business Process Flow: " + content.bpf.GetDisplayName());
        }

        private void addOverview()
        {
            AddHeading(content.bpf.GetDisplayName(), "Heading1");
            body.AppendChild(new Paragraph(new Run()));

            Table table = CreateTable();
            table.Append(CreateRow(new Text("Name"), new Text(content.bpf.GetDisplayName())));
            table.Append(CreateRow(new Text("Unique Name"), new Text(content.bpf.UniqueName ?? "")));
            table.Append(CreateRow(new Text("Primary Entity"), new Text(content.GetTableDisplayName(content.bpf.PrimaryEntity))));
            table.Append(CreateRow(new Text("State"), new Text(content.bpf.GetStateLabel())));
            table.Append(CreateRow(new Text("Number of Stages"), new Text(content.bpf.Stages.Count.ToString())));
            if (!string.IsNullOrEmpty(content.bpf.Description))
                table.Append(CreateRow(new Text("Description"), new Text(content.bpf.Description)));
            if (!string.IsNullOrEmpty(content.bpf.IntroducedVersion))
                table.Append(CreateRow(new Text("Version"), new Text(content.bpf.IntroducedVersion)));
            table.Append(CreateRow(new Text(content.headerDocumentationGenerated),
                new Text(PowerDocuReleaseHelper.GetTimestampWithVersion())));
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addStagesOverview()
        {
            if (content.bpf.Stages.Count == 0) return;

            AddHeading(content.headerStages, "Heading2");
            body.AppendChild(new Paragraph(new Run(
                new Text($"This Business Process Flow has {content.bpf.Stages.Count} stage(s)."))));

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("#"), new Text("Stage Name"), new Text("Entity"), new Text("Category"), new Text("Steps"), new Text("Required Fields")));

            for (int i = 0; i < content.bpf.Stages.Count; i++)
            {
                BPFStage stage = content.bpf.Stages[i];
                int requiredCount = stage.Steps.Count(s => s.IsRequired);
                table.Append(CreateRow(
                    new Text((i + 1).ToString()),
                    new Text(stage.Name ?? ""),
                    new Text(content.GetTableDisplayName(stage.EntityName)),
                    new Text(stage.GetStageCategoryLabel()),
                    new Text(stage.Steps.Count.ToString()),
                    new Text(requiredCount.ToString())
                ));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addStageDetails()
        {
            if (content.bpf.Stages.Count == 0) return;

            AddHeading(content.headerStageDetails, "Heading2");

            for (int i = 0; i < content.bpf.Stages.Count; i++)
            {
                BPFStage stage = content.bpf.Stages[i];
                AddHeading($"Stage {i + 1}: {stage.Name}", "Heading3");

                // Stage info
                Table infoTable = CreateTable();
                infoTable.Append(CreateRow(new Text("Entity"), new Text(content.GetTableDisplayName(stage.EntityName))));
                infoTable.Append(CreateRow(new Text("Category"), new Text(stage.GetStageCategoryLabel())));
                if (!string.IsNullOrEmpty(stage.NextStageId))
                {
                    BPFStage nextStage = content.bpf.Stages.FirstOrDefault(s => s.StageId == stage.NextStageId);
                    infoTable.Append(CreateRow(new Text("Next Stage"), new Text(nextStage?.Name ?? stage.NextStageId)));
                }
                body.Append(infoTable);
                body.AppendChild(new Paragraph(new Run(new Break())));

                // Steps/Controls
                if (stage.Steps.Count > 0)
                {
                    Table stepTable = CreateTable();
                    stepTable.Append(CreateHeaderRow(new Text("#"), new Text("Control"), new Text("Data Field"), new Text("Required")));

                    int stepNum = 1;
                    foreach (BPFStep step in stage.Steps)
                    {
                        foreach (BPFControl control in step.Controls)
                        {
                            stepTable.Append(CreateRow(
                                new Text(stepNum.ToString()),
                                new Text(control.DisplayName ?? ""),
                                new Text(control.DataFieldName ?? ""),
                                new Text(step.IsRequired ? "Yes" : "No")
                            ));
                            stepNum++;
                        }
                        if (step.Controls.Count == 0)
                        {
                            stepTable.Append(CreateRow(
                                new Text(stepNum.ToString()),
                                new Text(step.Description ?? "(No control)"),
                                new Text(""),
                                new Text(step.IsRequired ? "Yes" : "No")
                            ));
                            stepNum++;
                        }
                    }
                    body.Append(stepTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Condition branches
                if (stage.ConditionBranches.Count > 0)
                {
                    body.AppendChild(new Paragraph(new Run(new Text("Conditional Branching:"))));
                    Table branchTable = CreateTable();
                    branchTable.Append(CreateHeaderRow(new Text("Condition"), new Text("Target Stage")));

                    foreach (BPFConditionBranch branch in stage.ConditionBranches)
                    {
                        BPFStage targetStage = content.bpf.Stages.FirstOrDefault(s => s.StageId == branch.TargetStageId);
                        branchTable.Append(CreateRow(
                            new Text(!string.IsNullOrEmpty(branch.Description) ? branch.Description : "(Default)"),
                            new Text(targetStage?.Name ?? branch.TargetStageId ?? "")
                        ));
                    }
                    body.Append(branchTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }
            }
        }

        private void addProperties()
        {
            AddHeading(content.headerProperties, "Heading2");

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Property"), new Text("Value")));
            table.Append(CreateRow(new Text("ID"), new Text(content.bpf.ID ?? "")));
            table.Append(CreateRow(new Text("Unique Name"), new Text(content.bpf.UniqueName ?? "")));
            table.Append(CreateRow(new Text("Primary Entity"), new Text(content.bpf.PrimaryEntity ?? "")));
            table.Append(CreateRow(new Text("Business Process Type"), new Text(content.bpf.BusinessProcessType.ToString())));
            table.Append(CreateRow(new Text("State"), new Text(content.bpf.GetStateLabel())));
            table.Append(CreateRow(new Text("Trigger On Create"), new Text(content.bpf.TriggerOnCreate ? "Yes" : "No")));
            table.Append(CreateRow(new Text("Is Customizable"), new Text(content.bpf.IsCustomizable ? "Yes" : "No")));
            if (!string.IsNullOrEmpty(content.bpf.IntroducedVersion))
                table.Append(CreateRow(new Text("Introduced Version"), new Text(content.bpf.IntroducedVersion)));
            body.Append(table);
        }
    }
}
