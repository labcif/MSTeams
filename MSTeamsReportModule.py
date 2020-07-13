#   LabCif-MSTeams: This project extracts MSTeams Artifacts
#   Copyright (C) 2020
#   Authors: Guilherme Beco and Tiago Garcia
#   Mentors: Miguel Monteiro de Sousa Frade and Patricio Rodrigues Domingues
#   This program is free software: you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation, either version 3 of the License, or
#   any later version.
#
#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.
#
#   You should have received a copy of the GNU General Public License
#   along with this program.  If not, see <https://www.gnu.org/licenses/>.


import os
from java.lang import System
from java.util.logging import Level
from org.sleuthkit.datamodel import TskData
from org.sleuthkit.autopsy.casemodule import Case
from org.sleuthkit.autopsy.coreutils import Logger
from org.sleuthkit.autopsy.report import GeneralReportModuleAdapter
from org.sleuthkit.autopsy.report.ReportProgressPanel import ReportStatus
from org.sleuthkit.autopsy.casemodule.services import FileManager
projectEIAppDataPath = os.getenv('APPDATA') + "\\ProjetoEI\\"

# Class responsible for defining module metadata and logic
class CSVReportModule(GeneralReportModuleAdapter):

    moduleName = "MS Teams Report"

    _logger = None
    def log(self, level, msg):
        if _logger == None:
            _logger = Logger.getLogger(self.moduleName)

        self._logger.logp(level, self.__class__.__name__, inspect.stack()[1][3], msg)

    def getName(self):
        return self.moduleName

    def getDescription(self):
        return "Writes CSV of file names and hash values"

    def getRelativeFilePath(self):
        return "index.html"

    # TODO: Update this method to make a report
    # The 'baseReportDir' object being passed in is a string with the directory that reports are being stored in.   Report should go into baseReportDir + getRelativeFilePath().
    # The 'progressBar' object is of type ReportProgressPanel.
    #   See: http://sleuthkit.org/autopsy/docs/api-docs/latest/classorg_1_1sleuthkit_1_1autopsy_1_1report_1_1_report_progress_panel.html
    def generateReport(self, baseReportDir, progressBar):

        # Setup the progress bar
        progressBar.setIndeterminate(False)
        progressBar.start()        

        f = open(os.path.join(projectEIAppDataPath,"filesToReport.txt"), "r")
        for line in f:
            line = line[:-1]
            reporfile = open(line, "r")

            # Open the output file
            reportName = line.split("\\")
            fileName = os.path.join(baseReportDir, reportName[-1])
            report = open(fileName, 'w')

            for contentline in reporfile:
                report.write(contentline)
            
            report.close()

            # Add the report to the Case, so it is shown in the tree
            if ".css" not in line:
                Case.getCurrentCase().addReport(fileName, self.moduleName, reportName[-1])


        progressBar.complete(ReportStatus.COMPLETE)