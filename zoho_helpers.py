import csv
import os
import json
import time
import sys
import re

#zoho_api_dir = '/Users/JPMartyn/python/zoho/api'
#sys.path.append(os.path.abspath(zoho_api_dir))

from zoho import Zoho

def ticket_to_json(z, ticket_id, output_path):
    status_code, ticket = z.get_ticket(ticket_id)
    if status_code == 200:
        with open(output_path, "w", encoding='utf-8') as outfile:
            json.dump(ticket, outfile, indent=4)
    else:    
        sys.exit("Error retrieving ticket " + ticket_id)

def ticket_attachments_to_dir(z, ticket_id, output_dir):
    attachments = z.get_ticket_attachments(ticket_id)
    if len(attachments) > 0:
        if not os.path.isdir(output_dir):
            os.makedirs(output_dir)
        for attachment in attachments:
            z.download_attachment(attachment, output_dir)

def ticket_history_to_csv(z, ticket_id, output_path):
    history = z.get_ticket_history(ticket_id)
    if len(history) > 0:
        with open(output_path, "w", newline='', encoding='utf-8') as outfile:
            csvoutfile = csv.writer(outfile, quoting=csv.QUOTE_ALL)
            # Write header to file
            header_row = ["eventTime","eventName","eventInfo", "actor", "actorInfo", "source"]    
            csvoutfile.writerow(header_row)
            for event in history:
                row = []
                row.append(event.get("eventTime"))
                row.append(event.get("eventName"))
                row.append(event.get("eventInfo"))
                row.append(event.get("actor"))
                row.append(event.get("actorInfo"))
                row.append(event.get("source"))
                csvoutfile.writerow(row)

def get_ticket_convos(z, ticket_id):
    full_convos = []
    # Get list of conversations (id + type for each)
    conversations = z.list_conversations(ticket_id)
    if len(conversations) > 0:
        for convo in conversations:
            # Get convo details
            details = z.get_convo_details(ticket_id, convo["id"], convo["type"])
            full_convos.append(details)
    return full_convos

def ticket_convos_to_csv(ticket_convos, output_path):
    with open(output_path, "w", newline='', encoding='utf-8') as outfile:
        csvoutfile = csv.writer(outfile, quoting=csv.QUOTE_ALL)
        header_row = ["id","type","created_time","modified_time","content","commenter","from","to","cc","bcc","attachments"]
        csvoutfile.writerow(header_row)
        for convo in ticket_convos:
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
            csvoutfile.writerow(row)

def convos_attachments_to_dir(z, ticket_convos, output_path):
    for convo in ticket_convos:
        if len(convo['attachments']) > 0:
            if not os.path.isdir(output_path):
                os.makedirs(output_path)
            convo_attachment_path = "convo_" + convo["id"]
            convo_attachment_path = os.path.join(output_path, convo_attachment_path)
            if not os.path.isdir(convo_attachment_path):
                os.makedirs(convo_attachment_path)
            attachments_to_dir(z, convo["attachments"], convo_attachment_path)

def task_to_json(z, task_id, output_path):
    status_code, task_attributes = z.get_task(task_id)
    if status_code == 200:
        with open(output_path, "w", encoding='utf-8') as outfile:
            json.dump(task_attributes, outfile, indent=4)
    else:    
        sys.exit("Error retrieving task " + task_id)

def task_attachments_to_dir(z, task_id, output_dir):
    attachments = z.get_task_attachments(task_id)
    if len(attachments) > 0:
        if not os.path.isdir(output_dir):
            os.makedirs(output_dir)
            for attachment in attachments:
                z.download_attachment(attachment, output_dir)

def get_full_task_comments(z, task_id):
    task_comments = z.get_task_comments(task_id)
    full_task_comments = []
    for comment in task_comments:
        status_code, comment_details = z.get_task_comment(task_id, comment["id"])
        full_task_comments.append(comment_details)
    return full_task_comments

def task_comments_to_csv(full_task_comments, output_path):
    with open(output_path, "w", newline='', encoding='utf-8') as outfile:
        csvoutfile = csv.writer(outfile, quoting=csv.QUOTE_ALL)
        header_row = ["id","content_type","created_time","content","commenter_email","attachments"]
        csvoutfile.writerow(header_row)
        for comment in full_task_comments:
            row = []
            row.append("'" + comment.get("id")) # Format char for Excel
            row.append(comment.get("contentType"))    
            row.append(comment.get("commentedTime"))
            row.append(comment.get("content")[:30000])    
            row.append(comment.get("commenter").get("email")) 
            row.append(comment.get("attachments"))    
            csvoutfile.writerow(row)

def task_comment_attachments_to_dirs(z, task_comments, output_dir):
    for comment in task_comments:
        if len(comment['attachments']) > 0:
            if not os.path.isdir(output_dir):
                os.makedirs(output_dir)
            comment_attachment_path = "comment_" + comment["id"]
            comment_attachment_path = os.path.join(output_dir, comment_attachment_path)
            if not os.path.isdir(comment_attachment_path):
                os.makedirs(comment_attachment_path)
            attachments_to_dir(z, comment["attachments"], comment_attachment_path)

def attachments_to_dir(z, attachments, output_dir):
    if len(attachments) > 0:
        for attachment in attachments:
            z.download_attachment(attachment, output_dir)