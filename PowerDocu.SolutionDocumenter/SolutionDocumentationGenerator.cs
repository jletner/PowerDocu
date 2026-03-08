using System;
using System.Collections.Generic;
using System.IO;
using PowerDocu.Common;
using PowerDocu.AgentDocumenter;
using PowerDocu.AppDocumenter;
using PowerDocu.AppModuleDocumenter;
using PowerDocu.FlowDocumenter;

namespace PowerDocu.SolutionDocumenter
{
    public static class SolutionDocumentationGenerator
    {
        private static List<FlowEntity> flows;
        private static List<AppEntity> apps;
        private static List<AgentEntity> agents;
        private static List<AppModuleEntity> appModules;

        public static void GenerateDocumentation(string filePath, bool fullDocumentation, ConfigHelper config, string outputPath = null)
        {
            if (File.Exists(filePath))
            {
                DateTime startDocGeneration = DateTime.Now;

                flows = FlowDocumentationGenerator.GenerateDocumentation(
                    filePath,
                    fullDocumentation,
                    config,
                    outputPath
                );

                apps = AppDocumentationGenerator.GenerateDocumentation(
                    filePath,
                    fullDocumentation,
                    config,
                    outputPath
                );
                agents = AgentDocumentationGenerator.GenerateDocumentation(
                    filePath,
                    fullDocumentation,
                    config,
                    outputPath
                );

                // Parse and document the solution if enabled
                if (config.documentSolution)
                {
                    SolutionParser solutionParser = new SolutionParser(filePath);
                    if (solutionParser.solution != null)
                    {
                        string path = outputPath == null
                            ? Path.GetDirectoryName(filePath) + @"\Solution " + CharsetHelper.GetSafeName(Path.GetFileNameWithoutExtension(filePath) + @"\")
                            : outputPath + @"\" + CharsetHelper.GetSafeName(Path.GetFileNameWithoutExtension(filePath) + @"\");

                        // Generate Model-Driven App documentation
                        if (config.documentModelDrivenApps)
                        {
                            appModules = AppModuleDocumentationGenerator.GenerateDocumentation(
                                solutionParser.solution,
                                fullDocumentation,
                                config,
                                path,
                                apps
                            );
                        }

                        SolutionDocumentationContent solutionContent = new SolutionDocumentationContent(solutionParser.solution, apps, flows, appModules, path);
                        DataverseGraphBuilder dataverseGraphBuilder = new DataverseGraphBuilder(solutionContent);

                        if (fullDocumentation)
                        {
                            if (config.outputFormat.Equals(OutputFormatHelper.Word) || config.outputFormat.Equals(OutputFormatHelper.All))
                            {
                                NotificationHelper.SendNotification("Creating Solution documentation");
                                SolutionWordDocBuilder wordzip = new SolutionWordDocBuilder(solutionContent, config.wordTemplate, config.documentDefaultColumns, config.addTableOfContents);
                            }
                            if (config.outputFormat.Equals(OutputFormatHelper.Markdown) || config.outputFormat.Equals(OutputFormatHelper.All))
                            {
                                SolutionMarkdownBuilder mdDoc = new SolutionMarkdownBuilder(solutionContent, config.documentDefaultColumns);
                            }
                            if (config.outputFormat.Equals(OutputFormatHelper.Html) || config.outputFormat.Equals(OutputFormatHelper.All))
                            {
                                NotificationHelper.SendNotification("Creating HTML Solution documentation");
                                SolutionHtmlBuilder htmlDoc = new SolutionHtmlBuilder(solutionContent, config.documentDefaultColumns);
                            }
                            // Free cached SVG content now that all output formats have been generated
                            FormSvgBuilder.ClearCache();
                        }
                    }
                }

                DateTime endDocGeneration = DateTime.Now;
                NotificationHelper.SendNotification($"SolutionDocumenter: Created documentation for {filePath}. Total solution documentation completed in {(endDocGeneration - startDocGeneration).TotalSeconds} seconds.");
            }
            else
            {
                NotificationHelper.SendNotification($"File not found: {filePath}");
            }
        }
    }
}