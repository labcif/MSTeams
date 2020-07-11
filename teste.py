# Simple data source-level ingest module for Autopsy.
# Search for TODO for the things that you need to change
# See http://sleuthkit.org/autopsy/docs/api-docs/latest/index.html for documentation
import io
import jarray
import os
import time as tim
import inspect
import csv
import subprocess
import inspect
import shutil
import time
import threading
from datetime import datetime
import math
from java.lang import System
from java.util.logging import Level
from org.sleuthkit.datamodel import SleuthkitCase
from org.sleuthkit.datamodel import AbstractFile
from org.sleuthkit.datamodel import ReadContentInputStream
from org.sleuthkit.datamodel import BlackboardArtifact
from org.sleuthkit.datamodel import BlackboardAttribute
from org.sleuthkit.autopsy.ingest import IngestModule
from org.sleuthkit.autopsy.ingest.IngestModule import IngestModuleException
from org.sleuthkit.autopsy.ingest import DataSourceIngestModule
from org.sleuthkit.autopsy.ingest import FileIngestModule
from org.sleuthkit.autopsy.ingest import IngestModuleFactoryAdapter
from org.sleuthkit.autopsy.ingest import IngestMessage
from org.sleuthkit.autopsy.ingest import IngestServices
from org.sleuthkit.autopsy.coreutils import Logger
from org.sleuthkit.autopsy.casemodule import Case
from org.sleuthkit.autopsy.casemodule.services import Services
from org.sleuthkit.autopsy.casemodule.services import FileManager
from org.sleuthkit.autopsy.casemodule.services import Blackboard
tm=""
paths = {}
message=""
projectEIAppDataPath = os.getenv('APPDATA') + "\\ProjetoEI\\"
users=[]
pathsLDB={}

# Factory that defines the name and details of the module and allows Autopsy
# to create instances of the modules that will do the analysis.
# TODO: Rename this to something more specific. Search and replace for it because it is used a few times
class LabcifMSTeamsDataSourceIngestModuleFactory(IngestModuleFactoryAdapter):

    # TODO: give it a unique name.  Will be shown in module list, logs, etc.
    moduleName = "MSTeams Source Module"

    def getModuleDisplayName(self):
        return self.moduleName

    # TODO: Give it a description
    def getModuleDescription(self):
        return "MSTeams extract Teams artifacts."

    def getModuleVersionNumber(self):
        return "1.0"

    def isDataSourceIngestModuleFactory(self):
        return True

    def createDataSourceIngestModule(self, ingestOptions):
        # TODO: Change the class name to the name you'll make below
        return LabcifMSTeamsDataSourceIngestModule()


# Data Source-level ingest module.  One gets created per data source.
# TODO: Rename this to something more specific. Could just remove "Factory" from above name.
class LabcifMSTeamsDataSourceIngestModule(DataSourceIngestModule):

    _logger = Logger.getLogger(LabcifMSTeamsDataSourceIngestModuleFactory.moduleName)

    def log(self, level, msg):
        self._logger.logp(level, self.__class__.__name__, inspect.stack()[1][3], msg)

    def __init__(self):
        self.context = None
        self.art_list = []
    def create_artifact_type(self, art_name, art_desc, blackboard):
        try:
            art = blackboard.getOrAddArtifactType(art_name, "Labcif-MSTEAMS: " + art_desc)
            self.art_list.append(art)
        except Exception as e :
            self.log(Level.INFO, "Error getting or adding artifact type: " + art_desc + " " + str(e))
        return art
    def create_attribute_type(self, att_name, type_name, att_desc, blackboard):
        try:
            att_type = blackboard.getOrAddAttributeType(att_name, type_name, att_desc)
        except Exception as e:
            self.log(Level.INFO, "Error getting or adding attribute type: " + att_desc + " " + str(e))
        return att_type

    # Where any setup and configuration is done
    # 'context' is an instance of org.sleuthkit.autopsy.ingest.IngestJobContext.
    # See: http://sleuthkit.org/autopsy/docs/api-docs/latest/classorg_1_1sleuthkit_1_1autopsy_1_1ingest_1_1_ingest_job_context.html
    # TODO: Add any setup code that you need here.
    def startUp(self, context):
        
        # Throw an IngestModule.IngestModuleException exception if there was a problem setting up
        # raise IngestModuleException("Oh No!")
        self.context = context
        if not os.path.exists(projectEIAppDataPath):
            try:
                os.mkdir(projectEIAppDataPath)
            except OSError:
                print("Creation of the directory %s failed" % projectEIAppDataPath)
            else:
                print("Successfully created the directory %s " % projectEIAppDataPath)  

    # Where the analysis is done.
    # The 'dataSource' object being passed in is of type org.sleuthkit.datamodel.Content.
    # See: http://www.sleuthkit.org/sleuthkit/docs/jni-docs/latest/interfaceorg_1_1sleuthkit_1_1datamodel_1_1_content.html
    # 'progressBar' is of type org.sleuthkit.autopsy.ingest.DataSourceIngestModuleProgress
    # See: http://sleuthkit.org/autopsy/docs/api-docs/latest/classorg_1_1sleuthkit_1_1autopsy_1_1ingest_1_1_data_source_ingest_module_progress.html
    # TODO: Add your analysis code in here.
    def process(self, dataSource, progressBar):

        # we don't know how much work there is yet
        progressBar.switchToIndeterminate()
        self.log(Level.INFO,dataSource.getUniquePath())
        # Use blackboard class to index blackboard artifacts for keyword search
        blackboard = Case.getCurrentCase().getServices().getBlackboard()
        self.art_contacts = self.create_artifact_type("Labcif-MSTeams_CONTACTS_"," Contacts", blackboard)
        self.art_messages = self.create_artifact_type("Labcif-MSTeams_MESSAGES_"," MESSAGES", blackboard)
        self.art_messages_reacts = self.create_artifact_type("Labcif-MSTeams_MESSAGES_REACTS"," REACTS", blackboard)
        self.art_messages_files = self.create_artifact_type("Labcif-MSTeams_MESSAGES_FILES"," FILES", blackboard)
        self.art_call = self.create_artifact_type("Labcif-MSTeams_CALLS_", " Call history", blackboard)
        self.art_call_one_to_one = self.create_artifact_type("Labcif-MSTeams_CALLS_ONE_TO_ONE", " Call history one to one", blackboard)
        self.art_teams = self.create_artifact_type("Labcif-MSTeams_TEAMS_"," Teams", blackboard)
        
        # contactos
        self.att_name = self.create_attribute_type('Labcif-MSTeams_CONTACT_NAME', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Name", blackboard)
        self.att_email = self.create_attribute_type('Labcif-MSTeams_CONTACT_EMAIL', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Email", blackboard)
        self.att_orgid = self.create_attribute_type('Labcif-MSTeams_CONTACT_ORGID', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Orgid", blackboard)
        self.att_user_contacts = self.create_attribute_type('Labcif-MSTeams_USERNAME_CONTACTS', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "User", blackboard)
        self.att_folder_extract_contacts = self.create_attribute_type('Labcif-MSTeams_FOLDER_EXTRACT_CONTACTS', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Folder of extraction", blackboard)
        # reacts
        self.att_message_id_reacts = self.create_attribute_type('Labcif-MSTeams_MESSAGE_ID_REACTS', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Message ID", blackboard)
        self.att_sender_name_react = self.create_attribute_type('Labcif-MSTeams_MESSAGE_SENDER_NAME_REACTS', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Who reacted", blackboard)
        self.att_reacted_with = self.create_attribute_type('Labcif-MSTeams_MESSAGE_FILE_LOCAL_EMOJI_REACTS', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Emoji", blackboard)
        self.att_react_time= self.create_attribute_type('Labcif-MSTeams_MESSAGE_REACT_TIME', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "React time", blackboard)
        self.att_user_message_reacts = self.create_attribute_type('Labcif-MSTeams_USERNAME_MESSAGE_REACTS', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "User", blackboard)
        self.att_folder_extract_reacts = self.create_attribute_type('Labcif-MSTeams_FOLDER_EXTRACT_REACTS', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Folder of extraction", blackboard)
        # mensagens
        self.att_message_id = self.create_attribute_type('Labcif-MSTeams_MESSAGE_ID', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Message ID", blackboard)
        self.att_message = self.create_attribute_type('Labcif-MSTeams_MESSAGE', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Message", blackboard)
        self.att_sender_name = self.create_attribute_type('Labcif-MSTeams_SENDER', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Senders name", blackboard)
        self.att_time = self.create_attribute_type('Labcif-MSTeams_TIME', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Message time", blackboard) 
        self.att_cvid = self.create_attribute_type('Labcif-MSTeams_CONVERSATION_ID', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "CV", blackboard)
        self.att_user_message = self.create_attribute_type('Labcif-MSTeams_USERNAME_MESSAGE', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "User", blackboard)
        self.att_folder_extract_message = self.create_attribute_type('Labcif-MSTeams_FOLDER_EXTRACT_MESSAGES', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Folder of extraction", blackboard)
         # ficheiros
        self.att_message_id_files = self.create_attribute_type('Labcif-MSTeams_MESSAGE_ID', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Message ID", blackboard)
        self.att_file_name = self.create_attribute_type('Labcif-MSTeams_MESSAGE_FILE_NAME', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "File name", blackboard)
        self.att_file_local = self.create_attribute_type('Labcif-MSTeams_MESSAGE_FILE_LINK', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "File Link", blackboard)
        self.att_user_message_files = self.create_attribute_type('Labcif-MSTeams_USERNAME_MESSAGE_FILES', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "User", blackboard)
        self.att_folder_extract_files = self.create_attribute_type('Labcif-MSTeams_FOLDER_EXTRACT_FILES', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Folder of extraction", blackboard)
        # calls one to one 
        self.att_date_start_one_to_one = self.create_attribute_type('Labcif-MSTeams_CALL_ONE_TO_ONE_TIME_START', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Call one to one time start", blackboard) 
        self.att_date_finish_one_to_one = self.create_attribute_type('Labcif-MSTeams_CALL_ONE_TO_ONE_TIME_FINISH', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Call one to one time finish", blackboard) 
        self.att_creator_name_one_to_one = self.create_attribute_type('Labcif-MSTeams_CALL_ONE_TO_ONE_CREATOR_NAME', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Call one to one Creator Name", blackboard)
        self.att_creator_email_one_to_one = self.create_attribute_type('Labcif-MSTeams_CALL_ONE_TO_ONE_CREATOR_EMAIL', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Call one to one Creator Email", blackboard)
        self.att_participant_name_one_to_one = self.create_attribute_type('Labcif-MSTeams_CALL_ONE_TO_ONE_PARTICIPANT_NAME', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Call one to one Participant Name", blackboard)
        self.att_participant_email_one_to_one = self.create_attribute_type('Labcif-MSTeams_CALL_ONE_TO_ONE_PARTICIPANT_EMAIL', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Call one to one Participant Email", blackboard)
        self.att_state_one_to_one = self.create_attribute_type('Labcif-MSTeams_CALL_ONE_TO_ONE_STATE', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Call one to one state", blackboard)
        self.att_user_calls_one_to_one = self.create_attribute_type('Labcif-MSTeams_USERNAME_CALLS', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "User", blackboard)
        self.att_folder_extract_calls_one_to_one = self.create_attribute_type('Labcif-MSTeams_FOLDER_EXTRACT_CALLS_ONE_TO_ONE', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Folder of extraction", blackboard)
        # teams
        self.att_cv_id_teams = self.create_attribute_type('Labcif-MSTeams_CV_ID_TEAMS', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Conversation ID teams", blackboard) 
        self.att_creator_name_teams = self.create_attribute_type('Labcif-MSTeams_TEAMS_CREATOR_NAME', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Teams Creator Name", blackboard)
        self.att_creator_email_teams = self.create_attribute_type('Labcif-MSTeams_TEAMS_CREATOR_EMAIL', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Teams Creator Email", blackboard)
        self.att_participant_name_teams = self.create_attribute_type('Labcif-MSTeams_TEAMS_PARTICIPANT_NAME', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Teams Participant Name", blackboard)
        self.att_participant_email_teams = self.create_attribute_type('Labcif-MSTeams_teams_PARTICIPANT_EMAIL_ONE_TO_ONE', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Teams Participant Email", blackboard)
        self.att_user_teams = self.create_attribute_type('Labcif-MSTeams_USERNAME_TEAMS', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "User", blackboard)
        self.att_folder_extract_teams = self.create_attribute_type('Labcif-MSTeams_FOLDER_EXTRACT_TEAMS', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Folder of extraction", blackboard) 
        # calls
        self.att_date = self.create_attribute_type('Labcif-MSTeams_CALL_DATE', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Call Date", blackboard) 
        self.att_creator_name = self.create_attribute_type('Labcif-MSTeams_CALL_CREATOR_NAME', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Creator Name", blackboard)
        self.att_creator_email = self.create_attribute_type('Labcif-MSTeams_CALL_CREATOR_EMAIL', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Creator Email", blackboard)
        self.att_count_people_in = self.create_attribute_type('Labcif-MSTeams_CALL_AMOUNT_PEOPLE_IN', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Amount of people in call", blackboard)
        self.att_duration = self.create_attribute_type('Labcif-MSTeams_CALL_DURANTION', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Call Duration", blackboard) 
        self.att_participant_name = self.create_attribute_type('Labcif-MSTeams_CALL_PARTICIPANT_NAME', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Participant Name", blackboard)
        self.att_participant_email = self.create_attribute_type('Labcif-MSTeams_CALL_PARTICIPANT_EMAIL', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Participant Email", blackboard)
        self.att_user_calls = self.create_attribute_type('Labcif-MSTeams_USERNAME_CALLS', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "User", blackboard)
        self.att_folder_extract_calls = self.create_attribute_type('Labcif-MSTeams_FOLDER_EXTRACT_CALL', BlackboardAttribute.TSK_BLACKBOARD_ATTRIBUTE_VALUE_TYPE.STRING, "Folder of extraction", blackboard)
        
        
        # For our example, we will use FileManager to get all
        # files with the word "test"
        # in the name and then count and read them
        # FileManager API: http://sleuthkit.org/autopsy/docs/api-docs/latest/classorg_1_1sleuthkit_1_1autopsy_1_1casemodule_1_1services_1_1_file_manager.html
        fileManager = Case.getCurrentCase().getServices().getFileManager()
        files = fileManager.findFiles(dataSource, "%.ldb","https_teams.microsoft.com_")
        

        numFiles = len(files)
        progressBar.switchToDeterminate(numFiles)
        fileCount = 0
        for file in files:

            # Check if the user pressed cancel while we were busy
            if self.context.isJobCancelled():
                return IngestModule.ProcessResult.OK

            fileCount += 1

            # Make an artifact on the blackboard.  TSK_INTERESTING_FILE_HIT is a generic type of
            # artfiact.  Refer to the developer docs for other examples.

            src = file.getParentPath()
            pathSplited=src.split("/")
            user=pathSplited[2]
            if user not in users:
                users.append(user)
            buffer = jarray.zeros(file.getSize(), "b")
            file.read(buffer,0,file.getSize())
            if "lost" not in src and "Roaming" in file.getParentPath() and "ProjetoEI" not in file.getParentPath():
                if src not in paths:
                    tm = datetime.fromtimestamp(math.floor(tim.time())).strftime("%m-%d-%Y_%Hh-%Mm-%Ss")
                    paths[src]="Analysis_Autopsy_LDB_{}_{}".format(user,tm)
                if not os.path.exists(os.path.join(projectEIAppDataPath,paths[src])):
                    try:
                        os.mkdir(os.path.join(projectEIAppDataPath,paths[src]))
                    except OSError:
                        print("Creation of the directory %s failed" % os.path.join(projectEIAppDataPath,paths[src]))
                    else:
                        print("Successfully created the directory %s " % os.path.join(projectEIAppDataPath,paths[src]))
                f = open(os.path.join(os.path.join(projectEIAppDataPath,paths[src]),file.getName()),"wb")
                f.write(buffer.tostring())
                f.close()
            # try:
            #     # index the artifact for keyword search
            #     blackboard.indexArtifact(art)
            # except Blackboard.BlackboardException as e:
            #     self.log(Level.SEVERE, "Error indexing artifact " + art.getDisplayName()+str(e))

            # To further the example, this code will read the contents of the file and count the number of bytes
            # Update the progress bar
            progressBar.progress(fileCount)
            for src, path in paths.items():
                complementaryFiles=fileManager.findFilesByParentPath(dataSource.getId(),src)
                for file in complementaryFiles:
                    if "lost" not in file.getParentPath() and ".ldb" not in file.getName() and "lost" not in file.getName() and "Roaming" in file.getParentPath() and "ProjetoEI" not in file.getParentPath():
                        if file.getName() == "." or file.getName() == ".." or "-slack" in file.getName():
                            continue
                        buffer = jarray.zeros(file.getSize(), "b")
                        self.log(Level.INFO,src)
                        if src not in paths:
                            tm = datetime.fromtimestamp(math.floor(tim.time())).strftime("%m-%d-%Y_%Hh-%Mm-%Ss")
                            paths[src] = "Analysis_Autopsy_LDB_{}_{}".format(user,tm)
                        if not os.path.exists(os.path.join(projectEIAppDataPath,paths[src])):
                            try:
                                os.mkdir(os.path.join(projectEIAppDataPath,paths[src]))
                            except OSError:
                                print("Creation of the directory %s failed" % os.path.join(projectEIAppDataPath,paths[src]))
                            else:
                                print("Successfully created the directory %s " % os.path.join(projectEIAppDataPath,paths[src]))
                        try:
                            f = open(os.path.join(os.path.join(projectEIAppDataPath,paths[src]),file.getName()),"a")
                            file.read(buffer,0,file.getSize())
                            f.write(buffer.tostring())
                            f.close()
                        except :
                            self.log(Level.INFO,"File Crash")
        pathModule = os.path.realpath(__file__)
        indexCutPath=pathModule.rfind("\\")
        pathModule=pathModule[0:indexCutPath+1]
        
        # message = IngestMessage.createMessage(
        #     IngestMessage.MessageType.DATA, Labcif-MSTeamsFactory.moduleName,
        #         str(self.filesFound) + " files found")
        analysisPath = ""
        result = {}
        for key,value in paths.items():
            if key not in result:
                result[key] = value
        self.log(Level.INFO,"ALO")
        for key, value in result.items():
            p = subprocess.Popen([r"{}EI\EI.exe".format(pathModule),"--pathToEI",r"{}EI\ ".format(pathModule), "-a", value],stderr=subprocess.PIPE)
            out = p.stderr.read()
            p.wait()
            # os.system("cmd /c \"{}EI\\EI.exe\" --pathToEI \"{}EI\\\" -a {}".format(pathModule,pathModule,value))
            self.log(Level.INFO,"standalone finished")
        results=[]
        pathResults="Analise Autopsy"
        for u in users:
            pathLDB=""
            for key,value in paths.items():
                if "Analysis_Autopsy_LDB_{}".format(u) in value:
                    pathLDB=value
                    break
            for root, dirs, files in os.walk(projectEIAppDataPath, topdown=False):
                for name in dirs:
                    if pathResults in name and os.stat(os.path.join(projectEIAppDataPath,pathLDB)).st_mtime < os.stat(os.path.join(projectEIAppDataPath,name)).st_mtime:
                        pathsLDB[pathLDB]=os.path.join(projectEIAppDataPath,name)
                        results.append(os.path.join(projectEIAppDataPath,name))
        f = open(os.path.join(projectEIAppDataPath,"filesToReport.txt"),"w")
        nCSS=0
        for r in results:
            for files in os.walk(r,topdown=False):
                for name in files:
                    for fileName in name:
                        if ".csv" in fileName or ".html" in fileName or ".css" in fileName:
                            f.write(os.path.join(r,fileName)+"\n")
        f.close()

        f = open(os.path.join(projectEIAppDataPath,"filesToReport.txt"), "r")
        for line in f:
            line = line.replace("\n","")
            pathExtract=""
            self.log(Level.INFO,line)
            if ".csv" in line:
                # ok
                if "EventCall" in line:
                    rowcount=0
                    for key,value in pathsLDB.items():
                        if value in line:
                            for k,v in paths.items():
                                if v == key:
                                    pathExtract=k
                                    break
                    with io.open(line,encoding="utf-8") as csvfile:
                        reader = csv.reader(x.replace('\0', '') for x in csvfile)  
                        for row in reader: # each row is a list
                            try:
                                row = row[0].split(";")
                                if rowcount!=0:
                                    art = dataSource.newArtifact(self.art_call.getTypeID())
                                    dura=str(int(float(row[4])))
                                    art.addAttribute(BlackboardAttribute(self.att_date, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[0])))
                                    art.addAttribute(BlackboardAttribute(self.att_creator_name, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[1])))
                                    art.addAttribute(BlackboardAttribute(self.att_creator_email, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[2])))
                                    art.addAttribute(BlackboardAttribute(self.att_count_people_in, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[3])))
                                    art.addAttribute(BlackboardAttribute(self.att_duration, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName,dura ))
                                    art.addAttribute(BlackboardAttribute(self.att_participant_name, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[5])))
                                    art.addAttribute(BlackboardAttribute(self.att_participant_email, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[6])))
                                    art.addAttribute(BlackboardAttribute(self.att_user_calls, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[7])))
                                    art.addAttribute(BlackboardAttribute(self.att_folder_extract_calls, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, pathExtract))
                            except:
                                self.log(Level.INFO,"File empty")
                            rowcount+=1
                        csvfile.close()
                # ok
                elif "Conversations" in line:
                    rowcount=0
                    for key,value in pathsLDB.items():
                        if value in line:
                            for k,v in paths.items():
                                if v == key:
                                    pathExtract=k
                                    break
                    with io.open(line,encoding="utf-8") as csvfile:
                        reader = csv.reader(x.replace('\0', '') for x in csvfile)  
                        for row in reader: # each row is a list
                            try:
                                row = row[0].split(";")
                                if rowcount!=0:
                                    art = dataSource.newArtifact(self.art_teams.getTypeID())
                                    art.addAttribute(BlackboardAttribute(self.att_cv_id_teams, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[0])))
                                    art.addAttribute(BlackboardAttribute(self.att_creator_name_teams, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[1])))
                                    art.addAttribute(BlackboardAttribute(self.att_creator_email_teams, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[2])))
                                    art.addAttribute(BlackboardAttribute(self.att_participant_name_teams, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[3])))
                                    art.addAttribute(BlackboardAttribute(self.att_participant_email_teams, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[4])))
                                    art.addAttribute(BlackboardAttribute(self.att_user_teams, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[5])))
                                    art.addAttribute(BlackboardAttribute(self.att_folder_extract_teams, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, pathExtract))
                            except:
                                self.log(Level.INFO,"File empty")
                            rowcount+=1
                        csvfile.close()
                # ok
                elif "CallOneToOne" in line:
                    rowcount=0
                    for key,value in pathsLDB.items():
                        if value in line:
                            for k,v in paths.items():
                                if v == key:
                                    pathExtract=k
                                    break
                    with io.open(line,encoding="utf-8") as csvfile:
                        reader = csv.reader(x.replace('\0', '') for x in csvfile)  
                        for row in reader: # each row is a list
                            try:
                                row = row[0].split(";")
                                if rowcount!=0:
                                    art = dataSource.newArtifact(self.art_call_one_to_one.getTypeID())
                                    art.addAttribute(BlackboardAttribute(self.att_date_start_one_to_one, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[2])))
                                    art.addAttribute(BlackboardAttribute(self.att_date_finish_one_to_one, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[3])))
                                    art.addAttribute(BlackboardAttribute(self.att_creator_name_one_to_one, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[0])))
                                    art.addAttribute(BlackboardAttribute(self.att_creator_email_one_to_one, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[1])))
                                    art.addAttribute(BlackboardAttribute(self.att_participant_name_one_to_one, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[4])))
                                    art.addAttribute(BlackboardAttribute(self.att_participant_email_one_to_one, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[5])))
                                    art.addAttribute(BlackboardAttribute(self.att_state_one_to_one, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[6])))
                                    art.addAttribute(BlackboardAttribute(self.att_user_calls_one_to_one, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[7])))
                                    art.addAttribute(BlackboardAttribute(self.att_folder_extract_calls_one_to_one, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, pathExtract))
                            except:
                                self.log(Level.INFO,"File empty")
                            rowcount+=1
                        csvfile.close()
                elif "Files" in line:
                    rowcount=0
                    for key,value in pathsLDB.items():
                        if value in line:
                            for k,v in paths.items():
                                if v == key:
                                    pathExtract=k
                                    break
                    with io.open(line,encoding="utf-8") as csvfile:
                        reader = csv.reader(x.replace('\0', '') for x in csvfile)  
                        for row in reader: # each row is a list
                            try:
                                row = row[0].split(";")
                                if rowcount!=0:
                                    art = dataSource.newArtifact(self.art_messages_files.getTypeID())
                                    art.addAttribute(BlackboardAttribute(self.att_message_id_files, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[0])))
                                    art.addAttribute(BlackboardAttribute(self.att_file_name, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[1])))
                                    art.addAttribute(BlackboardAttribute(self.att_file_local, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[2])))
                                    art.addAttribute(BlackboardAttribute(self.att_user_message_files, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[3])))
                                    art.addAttribute(BlackboardAttribute(self.att_folder_extract_files, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, pathExtract))
                            except:
                                self.log(Level.INFO,"File empty")
                            rowcount+=1
                        csvfile.close()
                elif "Mensagens" in line:
                    rowcount=0
                    for key,value in pathsLDB.items():
                        if value in line:
                            for k,v in paths.items():
                                if v == key:
                                    pathExtract=k
                                    break
                    idMessage=""
                    message=""
                    sender=""
                    timee=""
                    cvid=""
                    userMessage=""
                    with open(line) as csvfile:
                        reader = csv.reader(x.replace('\0', '') for x in csvfile)                     
                        for row in reader: # each row is a list
                            try:
                                if rowcount!=0:
                                    if len(row) == 1:
                                        row = row[0].split(";")
                                        idMessage=str(row[0])                                
                                        message=str(row[1])
                                        time=str(row[2])
                                        sender=str(row[3])
                                        cvid=str(row[4])
                                        userMessage=str(row[5])
                                    else:
                                        partOne = row[0].split(";")
                                        idMessage=str(partOne[0])
                                        lastPart=row[len(row)-1].split(";")
                                        time=str(lastPart[1])
                                        sender=str(lastPart[2])
                                        cvid=str(lastPart[3])
                                        userMessage=str(lastPart[4])
                                        message=str(partOne[1])+","
                                        if len(row)!=2:
                                            for x in range(1,len(row)-1):
                                                message+=str(row[x])+","
                                        message+=str(lastPart[0])
                                    art = dataSource.newArtifact(self.art_messages.getTypeID())
                                    art.addAttribute(BlackboardAttribute(self.att_message_id, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, idMessage))
                                    art.addAttribute(BlackboardAttribute(self.att_message, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, message))
                                    art.addAttribute(BlackboardAttribute(self.att_sender_name, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, sender))
                                    art.addAttribute(BlackboardAttribute(self.att_time, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, timee))
                                    art.addAttribute(BlackboardAttribute(self.att_cvid, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, cvid))
                                    art.addAttribute(BlackboardAttribute(self.att_user_message, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, userMessage))
                                    art.addAttribute(BlackboardAttribute(self.att_folder_extract_message, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, pathExtract))
                            except:
                                self.log(Level.INFO,"File empty")
                            rowcount+=1
                        csvfile.close()
                elif "Reacts" in line:
                    rowcount=0
                    for key,value in pathsLDB.items():
                        if value in line:
                            for k,v in paths.items():
                                if v == key:
                                    pathExtract=k
                                    break
                    with io.open(line,encoding="utf-8") as csvfile:
                        reader = csv.reader(x.replace('\0', '') for x in csvfile)  
                        for row in reader: # each row is a list
                            try:
                                row = row[0].split(";")
                                if rowcount!=0:
                                    art = dataSource.newArtifact(self.art_messages_reacts.getTypeID())
                                    try:
                                        art.addAttribute(BlackboardAttribute(self.att_message_id_reacts, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[0])))
                                        art.addAttribute(BlackboardAttribute(self.att_reacted_with, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[1])))
                                        art.addAttribute(BlackboardAttribute(self.att_sender_name_react, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[2])))
                                        art.addAttribute(BlackboardAttribute(self.att_react_time, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[3])))
                                        art.addAttribute(BlackboardAttribute(self.att_user_message_reacts, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[4])))
                                        art.addAttribute(BlackboardAttribute(self.att_folder_extract_reacts, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, pathExtract))
                                    except:
                                        pass
                                    else:
                                        pass
                            except:
                                self.log(Level.INFO,"File empty")
                            rowcount+=1
                        csvfile.close()
                elif "Contactos.csv" in line:
                    rowcount=0
                    for key,value in pathsLDB.items():
                        if value in line:
                            for k,v in paths.items():
                                if v == key:
                                    pathExtract=k
                                    break
                    with io.open(line,encoding="utf-8") as csvfile:
                        reader = csv.reader(x.replace('\0', '') for x in csvfile)  
                        for row in reader: # each row is a list
                            try:
                                row = row[0].split(";")
                                if rowcount!=0:
                                    art = dataSource.newArtifact(self.art_contacts.getTypeID())
                                    art.addAttribute(BlackboardAttribute(self.att_name, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[0])))
                                    art.addAttribute(BlackboardAttribute(self.att_email, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[1])))
                                    art.addAttribute(BlackboardAttribute(self.att_orgid, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[2])))
                                    art.addAttribute(BlackboardAttribute(self.att_user_contacts, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, str(row[3])))
                                    art.addAttribute(BlackboardAttribute(self.att_folder_extract_contacts, LabcifMSTeamsDataSourceIngestModuleFactory.moduleName, pathExtract))
                            except:
                                self.log(Level.INFO,"File empty")
                            rowcount+=1
                        csvfile.close()
            rowcount=0
        f.close()
        #Post a message to the ingest messages in box.
        message = IngestMessage.createMessage(IngestMessage.MessageType.DATA,
            "Sample Jython Data Source Ingest Module", "Please run MSTeams Report")
        IngestServices.getInstance().postMessage(message)

        return IngestModule.ProcessResult.OK

