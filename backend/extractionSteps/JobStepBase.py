import logging
from abc import ABC, abstractmethod

from util.excel_utils import Color, addFilterAndFreeze, resizeColumnWidth, writeColoredRow, writeUncoloredRow


class JobStepBase(ABC):
    def __init__(self, componentType: str):
        self.componentType = componentType

    @abstractmethod
    async def extract(self, controllerData):
        """
        Extraction step of AppDynamics data.
        API Calls will be made in this step only.
        """
        pass

    @abstractmethod
    def analyze(self, controllerData, thresholds):
        """
        Analysis step of extracted data.
        No API calls will be made in this step.
        Additional metadata will be calculated from extracted data.
        Report will be generated based on metrics exposed in 'analyze' step and input 'thresholds'.
        """
        pass

    def reportData(
        self,
        workbook,
        controllerData,
        jobStepName,
        useEvaluatedMetrics=True,
        colorRows=True,
    ):
        """
        Creation of Workbook sheet for evaluated analyzed data.
        No data analysis in this step.
        """
        """Create report sheet for raw analysis data."""
        logging.debug(f"Creating workbook sheet for raw details of {jobStepName}")

        metricFolder = "evaluated" if useEvaluatedMetrics else "raw"

        rawDataSheet = workbook.create_sheet(f"{jobStepName}")
        if len(list(controllerData.values())[0][self.componentType]) == 0:
            logging.warning(f"No data found for {jobStepName}")
            return

        rawDataHeaders = list(list(controllerData.values())[0][self.componentType].values())[0][jobStepName][metricFolder].keys()
        headers = ["controller", "application", "applicationId"] + (["description"] if self.componentType == "apm" else []) + list(rawDataHeaders)

        writeUncoloredRow(rawDataSheet, 1, headers)

        # Write Data
        rowIdx = 2
        for host, hostInfo in controllerData.items():
            for application in hostInfo[self.componentType].values():
                if colorRows:
                    data = [
                        ( hostInfo["controller"].host, None),
                        ( application["name"], None),
                        ( application["applicationId"] if self.componentType == "mrum" else application["id"], None),
                        *[application[jobStepName][metricFolder][header] for header in rawDataHeaders]
                    ]
                    if self.componentType == "apm":
                        data.insert(3, ( application["description"], None))
                    writeColoredRow(
                        rawDataSheet,
                        rowIdx,
                        data
                    )
                else:
                    data = [
                        hostInfo["controller"].host,
                        application["name"],
                        application["applicationId"] if self.componentType == "mrum" else application["id"],
                        *[application[jobStepName][metricFolder][header] for header in rawDataHeaders]
                    ]
                    if self.componentType == "apm":
                        data.insert(3, application["description"])
                    writeUncoloredRow(
                        rawDataSheet,
                        rowIdx,
                        data
                    )
                rowIdx += 1

        addFilterAndFreeze(rawDataSheet, "E2") if self.componentType == "apm" else addFilterAndFreeze(rawDataSheet, "D2")
        resizeColumnWidth(rawDataSheet)

    def applyThresholds(self, analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds):
        thresholdLevels = ["platinum", "gold", "silver"]

        # Calculate overall health across all thresholds and metrics a given for JobStep
        # This data goes into the 'Analysis' xlsx sheet.
        score = "bronze"
        color = Color[score]
        for thresholdLevel in thresholdLevels:
            numCriteriaWhichComplyWithCurrentThresholdLevel = 0

            for thresholdLevelMetric in jobStepThresholds[thresholdLevel].keys():
                if jobStepThresholds["direction"][thresholdLevelMetric] == "decreasing":
                    if analysisDataEvaluatedMetrics[thresholdLevelMetric] >= jobStepThresholds[thresholdLevel][thresholdLevelMetric]:
                        numCriteriaWhichComplyWithCurrentThresholdLevel += 1
                else:
                    if analysisDataEvaluatedMetrics[thresholdLevelMetric] <= jobStepThresholds[thresholdLevel][thresholdLevelMetric]:
                        numCriteriaWhichComplyWithCurrentThresholdLevel += 1

            if numCriteriaWhichComplyWithCurrentThresholdLevel == len(jobStepThresholds[thresholdLevel].keys()):
                score = thresholdLevel
                color = Color[score]
                break
        analysisDataRoot["computed"] = [score, color]

        # Calculate individual health of individual metrics.
        # This data goes into the 'JobStep - Metrics' xlsx sheet.
        for thresholdLevelMetric in analysisDataEvaluatedMetrics.keys():
            # Default to bronze, then loop through thresholds to apply correct color
            analysisDataEvaluatedMetrics[thresholdLevelMetric] = [
                analysisDataEvaluatedMetrics[thresholdLevelMetric],
                Color["bronze"],
            ]
            for thresholdLevel in thresholdLevels:
                if jobStepThresholds["direction"][thresholdLevelMetric] == "decreasing":
                    if analysisDataEvaluatedMetrics[thresholdLevelMetric][0] >= jobStepThresholds[thresholdLevel][thresholdLevelMetric]:
                        analysisDataEvaluatedMetrics[thresholdLevelMetric][1] = Color[thresholdLevel]
                        break
                else:
                    if analysisDataEvaluatedMetrics[thresholdLevelMetric][0] <= jobStepThresholds[thresholdLevel][thresholdLevelMetric]:
                        analysisDataEvaluatedMetrics[thresholdLevelMetric][1] = Color[thresholdLevel]
                        break
