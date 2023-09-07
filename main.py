import json
import logging.handlers
import os.path
import jira
import msal
from jira import JIRA
from requests.auth import HTTPBasicAuth
import requests
import credentials
import datetime


class SentinelOneIssueCreator:
    server = credentials.server
    username = credentials.username
    apiToken = credentials.apiToken
    ApplicationID = credentials.ApplicationID
    tenantID = credentials.tenantID
    SecretValue = credentials.SecretValue
    mail_receiver = credentials.mail_receiver
    mail_subject = credentials.mail_subject
    S1disabledId = credentials.S1disabledId
    S1disabledArchiveID = credentials.S1disabledArchiveID
    S1enabledId = credentials.S1enabledId
    S1enabledArchiveId = credentials.S1enabledArchiveId
    workSpaceId = credentials.workspaceID

    def __init__(self):
        super().__init__()
        try:
            self.jira = JIRA(server=self.server, basic_auth=(self.username, self.apiToken))
            self.jiraAuth = HTTPBasicAuth(self.username, self.apiToken)
        except Exception as e:
            logger.error(f"Can't Connect to JIRA. Error - {e}")

    """
    Generates Access Token for Microsoft Graph API
    :param - None
    :return - Access Token
    :requirements - tenant ID, Application ID and secret Value from Azure
    """

    def get_access_token(self):
        authority = f'https://login.microsoftonline.com/{self.tenantID}'
        scope = ['https://graph.microsoft.com/.default']
        try:
            app = msal.ConfidentialClientApplication(client_id=self.ApplicationID, authority=authority,
                                                     client_credential=self.SecretValue)
            accessToken = app.acquire_token_for_client(scopes=scope)
            return accessToken['access_token']
        except Exception as e:
            logger.error(f'Cannot retrieve Access Token. Error - {e}')
            return None

    """
    Runs a Scan through out the S1 enabled and disabled folder, creates or closes ticket if any S1 mails has been found.
    And moves the read ticket to Archive.
    :param - None
    :return - None
    """

    def run_scan(self):
        try:
            disabledServiceTags = self.search_servicetag_disabledmails()
            if disabledServiceTags:
                if disabledServiceTags[0] == "VALUES":
                    for serviceTag in disabledServiceTags[1]:
                        assetdict = self.get_assetowner_servicetag(serviceTag)
                        self.create_ticket(assetdict)
                if disabledServiceTags[0] == "ERROR":
                    logger.error(f'Faced an error while retrieving Disabled mails. Error - {disabledServiceTags[1]}')
            else:
                logger.info('No Disabled mails in the Disabled Folder')
            enabledServiceTags = self.search_servicetag_enabledmails()
            if enabledServiceTags:
                if enabledServiceTags[0] == "VALUES":
                    for serviceTag in enabledServiceTags[1]:
                        self.closeissue_servicetag(serviceTag)
                if enabledServiceTags[0] == "ERROR":
                    logger.error(f'Faced an error while retrieving Enabled mails. Error - {enabledServiceTags[1]}')
            else:
                logger.info('No Enabled mails in the Enabled Folder')
            return True
        except Exception as e:
            return e

    """
    Searches S1 Enabled Folder and returns list of service tags if there is any mails in the folder
    :param - None
    :return - None or list of service tags
    """

    def search_servicetag_enabledmails(self):
        serviceTags = []
        token = self.get_access_token()
        url = f"https://graph.microsoft.com/v1.0/users/{self.mail_receiver}/mailFolders/{self.S1enabledId}/messages"
        header = {
            "Authorization": f"Bearer {token}",
            "Content-type": "application/json",
            'Accept': 'application/json'
        }
        try:
            response = requests.get(url=url, headers=header)
            data = json.loads(response.text)
            values = data['value']
            if values:
                for value in values:
                    serviceTag = self.get_service_tag(value['subject'])
                    TimeStamp = value['receivedDateTime']
                    serviceTags.append([serviceTag, TimeStamp])
                    payload = {"destinationId": self.S1enabledArchiveId}
                    msg_id = value["id"]
                    url2 = f"https://graph.microsoft.com/v1.0/users/{self.mail_receiver}/messages/{msg_id}/move"
                    _ = requests.post(url=url2, headers=header, data=json.dumps(payload))
                    logger.info(
                        f'Enabled Mail for Service tag - {serviceTag[1]} has been moved to S1EnableArchive Folder ')
                return ["VALUES", serviceTags]
            else:
                return None
        except Exception as e:
            return ["ERROR", e]

    """
    Searches S1 Disabled Folder and returns list of service tags if there is any mails in the folder
    :param - None
    :return - None or list of service tags
    """

    def search_servicetag_disabledmails(self):
        serviceTags = []
        token = self.get_access_token()
        url = f"https://graph.microsoft.com/v1.0/users/{self.mail_receiver}/mailFolders/{self.S1disabledId}/messages"
        header = {
            "Authorization": f"Bearer {token}",
            "Content-type": "application/json",
            'Accept': 'application/json'
        }
        try:
            response = requests.get(url=url, headers=header)
            data = json.loads(response.text)
            values = data['value']
            if values:
                for value in values:
                    serviceTag = self.get_service_tag(value['subject'])
                    serviceTags.append(serviceTag)
                    payload = {"destinationId": self.S1disabledArchiveID}
                    msg_id = value["id"]
                    url2 = f"https://graph.microsoft.com/v1.0/users/{self.mail_receiver}/messages/{msg_id}/move"
                    _ = requests.post(url=url2, headers=header, data=json.dumps(payload))
                    logger.info(f'Disabled Mail for Service tag - {serviceTag[1]} has been moved to S1disableArchive '
                                f'Folder ')
                return ["VALUES", serviceTags]
            return None
        except Exception as e:
            return ['ERROR', e]

    """
    Returns service tag from Predefined subject
    """

    @staticmethod
    def get_service_tag(subject):
        hostName = subject.split("Machine ")[-1]
        serviceTag = hostName[-7:]
        return [serviceTag, hostName]

    """
    Closes a ticket based on the given service tag
    :param - Service Tag
    :return - None
    """

    def closeissue_servicetag(self, service_tag):
        jql = f"project = \"Information Technology/Systems (ITS)\" and summary ~ \"{service_tag[0][1]}\" and status = Open"
        issues = self.jira.search_issues(jql_str=jql)
        if issues:
            for issue in issues:
                try:
                    body = f"Received an email from Sentinel One - Agent enabled - Machine {service_tag[0][1]}.\nMail " \
                           f"recieved time : {service_tag[1]}.\nWill be closing the ticket."
                    comment = self.jira.add_comment(issue, body)
                    ticket = self.jira.issue(issue)
                    self.jira.transition_issue(ticket, '31')
                    logger.info(f'Commented on the Issue - {comment} [Comment ID]')
                    logger.info(f'Ticket {issue} closed. Service Tag - {service_tag[0][1]}')
                except Exception as e:
                    logger.error(f'Cannot close issue for service tag - {service_tag[0][1]}. Error - {e}')
        else:
            logger.warning(f'No Issue found for Service Tag - {service_tag[0][1]}')
    """
    Gets the details of the asset from JIRA to create a ticket
    :param - Service Tag
    :return - Dictionary consist of details required to create an issue
    """

    def get_assetowner_servicetag(self, service_tag):
        headers = {"Accept": "application/json",
                   "Content-Type": "application/json"}
        query = f'"objectType = Laptops AND label = {service_tag[0]}"'
        payload = '{"qlQuery": ' + query + '}'
        url = f"https://api.atlassian.com/jsm/assets/workspace/{self.workSpaceId}/v1/object/aql"
        hostName = service_tag[1]
        my_dict = {'Hostname': hostName, }
        try:
            Response = requests.post(
                url=url,
                headers=headers,
                data=payload,
                auth=self.jiraAuth
            )
            data = json.loads(Response.text)
            values = data['values']
            value = values[0]
            for attribute in value['attributes']:
                attributeID = attribute['objectTypeAttributeId']
                attributeVals = attribute['objectAttributeValues']
                attributeVal = attributeVals[0]
                if int(attributeID) == 27:
                    my_dict['Asset Owner'] = attributeVal['user']['key']
                    my_dict['Display Name'] = attributeVal['user']['displayName']
                if int(attributeID) == 29:
                    location = attributeVal['value']
                    if location == "India-Chennai":
                        my_dict['Reporter'] = "612fa701448a6b0069157259"
            return my_dict
        except Exception as e:
            logger.error(f'Cannot retrieve Asset details from JIRA. Error - {e}')
            my_dict = {
                'Hostname': hostName,
                'Reporter': "612fa701448a6b0069157259",
                'Asset Owner': None,
                'Display Name': None
            }
            return my_dict

    """
    Creates an issue with the given dictionary
    :param - Dictionary with keys ['Service Tag', 'Asset Owner', 'Display Name', 'Reporter']
    :return - newly created issue ID
    """

    def create_ticket(self, asset_details):
        hostName = asset_details.get('Hostname')
        assetOwner = asset_details.get('Asset Owner')
        name = asset_details.get('Display Name')
        reporter = asset_details.get('Reporter')
        if assetOwner:
            issue_dict = {
                'project': {'id': 10098},
                'summary': f'SentinelOne - Agent disabled - Machine {hostName} ',
                'description': f'Sentinel One Agent has been disabled in {name}\'s machine',
                'issuetype': {'id': 10020},
                'reporter': {'id': reporter},
                'assignee': {"id": assetOwner},
                'component': ''
            }
        else:
            issue_dict = {
                'project': {'id': 10098},
                'summary': f'SentinelOne - Agent disabled - Machine {hostName} ',
                'description': f'Sentinel One Agent has been disabled in Machine {hostName}',
                'issuetype': {'id': 10020},
                'reporter': {'id': reporter},
                'assignee': {"id": reporter}
            }
        try:
            new_issue = self.jira.create_issue(fields=issue_dict)
            logger.info(f'Ticket created for Service Tag {hostName}. Issue - {new_issue}')
        except Exception as e:
            logger.error(f'Cannot Create a ticket for Service Tag {hostName}. Error - {e}')

def GenerateFilename():
    dateTime = str(datetime.datetime.today()).split(" ")
    date = dateTime[0]
    if not os.path.exists("Logs-SentinelOne-Jira-Integration"):
        os.mkdir("Logs-SentinelOne-Jira-Integration")
    if os.path.exists(f"Logs-SentinelOne-Jira-Integration/Log-{date}"):
        return f"Logs-SentinelOne-Jira-Integration/Log-{date}/Log-{dateTime[1]}.log"
    else:
        os.mkdir(f"Logs-SentinelOne-Jira-Integration/Log-{date}")
        return f"Logs-SentinelOne-Jira-Integration/Log-{date}/Log-{dateTime[1]}.log"


if __name__ == "__main__":
    logfile = GenerateFilename()
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
    handler = logging.handlers.RotatingFileHandler(filename=logfile, mode='w',
                                                   backupCount=1)
    handler.setLevel(logging.INFO)
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.info(f'Mail Scan has Started, Time - {datetime.datetime.now()}')
    obj = SentinelOneIssueCreator()
    stags = obj.run_scan()
    if stags:
        logger.info(f'Mail Scan has ended, Time - {datetime.datetime.now()}')
    else:
        logger.error(f'Mail Scan Encountered an Error - {stags}. Time - {datetime.datetime.now()}')
