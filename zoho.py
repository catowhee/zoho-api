import json
import requests
from datetime import datetime
import time
import os
import csv
from urllib.parse import urlencode

class Zoho:
    """
    Todo: endpoint --> base URL (or URL?)
    """

    def __init__(self, zoho_auth_path):

        self._zoho_auth_path = zoho_auth_path

        with open(zoho_auth_path, 'r') as f:
            zoho_auth = json.load(f)

        # Zoho auth props
        self._creds = {
            "client_id": zoho_auth["client_id"],
            "client_secret": zoho_auth["client_secret"],
            "org_id": zoho_auth["org_id"],
            "access_token": zoho_auth["access_token"],
            "expiry_time": int(zoho_auth["expiry_time"]),
            "refresh_token": zoho_auth["refresh_token"],
            "scopes": zoho_auth["scopes"]
        }

        # API request headers
        # Todo: rename. Does not seem specific to GET.
        self._headers = {
            "orgId": self._creds["org_id"],
            "Authorization": "Zoho-oauthtoken " + self._creds["access_token"]
        }

        # API logs
        self._logs = []
        
    def _is_token_expired(self):
        return self._creds["expiry_time"] <= time.time()
    
    def _refresh_access_token(self):

        print("Before refresh:")
        print(self._creds)

        query = {
            "refresh_token": self._creds["refresh_token"],
            "client_id": self._creds["client_id"],
            "client_secret": self._creds["client_secret"],
            "grant_type": "refresh_token"
        }
        query_string = urlencode(query)

        endpoint = "https://accounts.zoho.com/oauth/v2/token?" + query_string
        r = requests.post(endpoint, headers=self._headers)
        content = json.loads(r.content.decode('utf-8'))

        try:
            new_access_token = content["access_token"]
            print("Access token refreshed.")
        except:
            raise Exception("Could not refresh access token")
        
        # Update props
        self._creds["access_token"] = new_access_token
        self._creds["expiry_time"] = time.time() + 60 * 60 - 30

        self._headers = {
            "orgId": self._creds["org_id"], 
            "Authorization": "Zoho-oauthtoken " + self._creds["access_token"]
        }
        
        # Dump updated creds to file
        with open(self._zoho_auth_path, 'w') as f:
            json.dump(self._creds, f, indent=4)

        time.sleep(5)

        
    def check_token(func):
        def wrapper(self, *args, **kwargs):
            if self._is_token_expired():
                self._refresh_access_token()
            else:
                #print("Access token not refreshed.")
                pass
            return func(self, *args, **kwargs)
        return wrapper

    @check_token
    def _base_api_call(self, method, endpoint, payload={}):

        # Todo: Should refresh token use this?

        count = 0
        content = "" 
        headers = self._headers # need to reconsider headers

        if method == "GET":
            content = None
            r = requests.get(endpoint, headers=headers)
            if r.status_code == 200:
                content = json.loads(r.content.decode('utf-8'))
                count = content.get("count")
            else:
                print(f"Warning\nHTTP Response: {r.status_code}\nEndpoint:{endpoint}")
        
        elif method == 'POST':
            r = requests.post(endpoint, headers=headers, json=payload)

        elif method == 'PUT':
            r = requests.put(endpoint, headers=headers, json=payload)

        
        # Log it
        log = [datetime.now().isoformat(), method, endpoint, payload, r.status_code, count]
        print(log)
        self._logs.append(log)


        # Return
        if method == 'GET':
            return (r.status_code, content)
        elif method == 'POST' or method == 'PUT':
            return r.status_code


    def _paginate(self, endpoint, query={}, results_limit = None):

        items = []
        query.setdefault('from', 0)
        query.setdefault('limit', 100)

        items_requested = 0

        while True:
            query_string = urlencode(query)
            #endpoint = f"{endpoint}?{query_string}"
            #print(endpoint + "?" + query_string)
            status_code, content = self._base_api_call("GET", endpoint + "?" + query_string)

            if status_code == 200:
                if isinstance(content, list):
                    items.extend(content)
                else:
                    items.extend(content["data"])
            elif status_code == 204:
                break
            elif 400 <= status_code < 600: # temp for recycle bin download
                break

            if len(content["data"]) < 100:
                break
            else:
                from_ = int(query["from"]) + int(query["limit"])
                query["from"] = from_

            # Update items requested
            items_requested += int(query["limit"])

            # If results_limit passed, check if it was reached, break if needed
            if results_limit is not None and items_requested >= results_limit:
                break

        query["from"] = 0
        return items
    
    def get_logs(self):
        return self._logs
    
    def write_logs_to_csv(self, output_dir):
        file_name = "zoho_api_logs_" + str(int(time.time())) + '.csv'
        with open(os.path.join(output_dir, file_name), 'w', newline='') as outfile:
            csvoutfile = csv.writer(outfile, quoting=csv.QUOTE_ALL)
            csvoutfile.writerow(["time", "method", "url", "payload" "http_code", "count" ])
            for log in self._logs:
                csvoutfile.writerow(log)
        
    def get_ticket(self, ticket_id):
        endpoint = f"https://desk.zoho.com/api/v1/tickets/{ticket_id}"
        return self._base_api_call("GET", endpoint)
    
    def trash_tickets(self, ticket_ids):
        endpoint = "https://desk.zoho.com/api/v1/tickets/moveToTrash"
        # todo: validate ticket_ids
        payload = {
            "ticketIds": ticket_ids
        }
        return self._base_api_call("POST", endpoint, payload)
    
    def close_tickets(self, ticket_ids):
        endpoint = "https://desk.zoho.com/api/v1/closeTickets"
        # todo: validate ticket_ids
        payload = {
            "ids": ticket_ids
        }
        return self._base_api_call("POST", endpoint, payload)

    def update_ticket(self, ticket_id, payload):
        endpoint = "https://desk.zoho.com/api/v1/tickets/" + ticket_id
        # todo: validate ticket_ids
        return self._base_api_call("PUT", endpoint, payload)
    
    def get_task(self, task_id):
        endpoint = f"https://desk.zoho.com/api/v1/tasks/{task_id}"
        return self._base_api_call("GET", endpoint)
    
    def get_comment(self, ticket_id, comment_id):  
        endpoint = f"https://desk.zoho.com/api/v1/tickets/{ticket_id}/comments/{comment_id}?include=plainText,mentions"
        status_code, content = self._base_api_call("GET", endpoint)

        # Replace "zsu" agent tags with the agent name
        # This can be done for other mention types if needed
        content["content"] = self._replace_mention_tags(content)

        return (status_code, content)
    
    def get_thread(self, ticket_id, thread_id):
        endpoint = f"https://desk.zoho.com/api/v1/tickets/{ticket_id}/threads/{thread_id}?include=plainText"
        return self._base_api_call("GET", endpoint)
    
    def get_task_comment(self, task_id, comment_id):
        endpoint = f"https://desk.zoho.com/api/v1/tasks/{task_id}/comments/{comment_id}?include=mentions"
        status_code, content = self._base_api_call("GET", endpoint)

    def get_organization_fields(self, module):
        endpoint = f"https://desk.zoho.com/api/v1/organizationFields?module={module}"    
        return self._base_api_call('GET', endpoint)

        # Replace "zsu" agent tags with the agent name
        # This can be done for other mention types if needed        
        content["content"] = self._replace_mention_tags(content)

        return (status_code, content)
    
    def _replace_mention_tags(self, content):
        if content.get("mention") is not None:
            mentions = content.get("mention")
            for mention in mentions:      
                if mention["type"] == "AGENT":
                    replace = "zsu[@user:" + mention["zuid"] + "]zsu"
                    replace_with = mention["firstName"] + " " + mention["lastName"]
                    content["content"] = content["content"].replace(replace, replace_with)
        return content["content"]

    
    @check_token
    def download_attachment(self, attachment, output_dir):
  
        url = attachment["href"]
        file_name = attachment["name"]

        r = requests.get(url, headers=self._headers)

        with open(os.path.join(output_dir, file_name), 'wb') as outfile:
            for chunk in r.iter_content(1024):
                outfile.write(chunk)
    
    def list_conversations(self, ticket_id):

        endpoint = f"https://desk.zoho.com/api/v1/tickets/{ticket_id}/conversations"
        convos = self._paginate(endpoint, {"limit": 100})

        # For now, keep only the id and type of each convo
        # To do - look into list comprehension
        for i in range(len(convos)):
            convos[i] = {
                "id": convos[i]["id"],
                "type": convos[i]["type"]
            }
        return convos

    def list_tasks_by_ticket(self, ticket_id):
        endpoint = f"https://desk.zoho.com/api/v1/tickets/{ticket_id}/tasks"
        tasks = self._paginate(endpoint)
        return tasks
    
    def get_task_comments(self, task_id, query = {"sortBy": "-commentedTime","limit": 100}):
        endpoint = f"https://desk.zoho.com/api/v1/tasks/{task_id}/comments"    
        query = {"sortBy": "-commentedTime","limit": 100}
        task_comments = self._paginate(endpoint, query)       
        return task_comments

    def search_tickets(self, query = {"limit": 100}):
        endpoint = "https://desk.zoho.com/api/v1/tickets/search"    
        #query = {"sortBy": "-commentedTime","limit": 100}
        results = self._paginate(endpoint, query)       
        return results
    
    def list_recycled(self, query = {"limit": 100}, results_limit = 500):
        endpoint = "https://desk.zoho.com/api/v1/recycleBin"    
        #query = {"sortBy": "-commentedTime","limit": 100}
        results = self._paginate(endpoint, query, results_limit)       
        return results

    def get_convo_details(self, ticket_id, convo_id, convo_type):

        details = {}
        details["id"] = convo_id
        details["type"] = convo_type

        if convo_type == "comment":
            status_code, content = self.get_comment(ticket_id, convo_id)
        elif convo_type == "thread":
            status_code, content = self.get_thread(ticket_id, convo_id)

        if status_code == 200:
            details["created_time"] = content.get("commentedTime", content.get("createdTime"))
            details["modified_time"] = content.get("modifiedTime")
            details["content"] = content.get("plainText")
            details["commenter"] = content.get("commenter", {}).get("name")
            details["commenter_photo"] = content.get("commenter", {}).get("photoURL")
            details["attachments"] = content.get("attachments")
            details["content_html"] = content.get("content")
            details["author_name"] = content.get("author", {}).get("name")
            details["from"] = content.get("fromEmailAddress")
            details["to"] = content.get("to")
            details["replyTo"] = content.get("replyTo")
            details["cc"] = content.get("cc")
            details["bcc"] = content.get("bcc")
            details["author_photo"] = content.get("author", {}).get("photoURL")
      
        return details

    def write_convo_to_csv(self, outfile, convo):
        row = []
        row.append("'" + convo.get("id")) # Formatting char for Excel
        row.append(convo.get("type"))
        row.append(convo.get("created_time"))
        row.append(convo.get("modified_time"))
        row.append(convo.get("content")[:30000])
        row.append(convo.get("commenter"))
        row.append(convo.get("from"))
        row.append(convo.get("to"))
        row.append(convo.get("cc"))
        row.append(convo.get("bcc"))
        row.append(convo.get("attachments"))     
        outfile.writerow(row)
        
    def get_ticket_attachments(self, ticket_id):
        endpoint = f"https://desk.zoho.com/api/v1/tickets/{ticket_id}/attachments"
        attachments = self._paginate(endpoint)       
        return attachments

    def get_task_attachments(self, task_id):
        endpoint = f"https://desk.zoho.com/api/v1/tasks/{task_id}/attachments"
        attachments = self._paginate(endpoint)       
        return attachments

    def get_ticket_history(self, ticket_id):
        endpoint = f"https://desk.zoho.com/api/v1/tickets/{ticket_id}/History"    
        query = {"limit": 50}
        history = self._paginate(endpoint, query)       
        return history